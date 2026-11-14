"""PowerPoint 発表者ノート自動動画化パイプライン.

PowerPoint (.pptx) の各スライドを画像として高解像度エクスポートし、
発表者ノートを edge-tts で音声化して同期・結合し、MP4 動画を生成する。

環境構築 (uv):
    1. uv venv .venv
    2. .venv\\Scripts\\activate
    3. uv pip install python-pptx edge-tts moviepy comtypes
"""

from __future__ import annotations

import argparse
import asyncio
import gc
import json
import logging
import os
import random
import re
import struct
import subprocess
import sys
import tempfile
import threading
import unicodedata
import urllib.parse
import urllib.request
import wave
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pronunciation import (
    DEFAULT_AI_BASE_URL,
    DEFAULT_AI_MODEL,
    load_pronunciation_dict,
    prepare_tts_text,
)

# ==============================================================================
# 話者 (Voice ID) 設定エリア
# ==============================================================================
VOICE_PRESETS: dict[str, str] = {
    "nanami": "ja-JP-NanamiNeural",   # 女性・落ち着いた標準的な声
    "keita": "ja-JP-KeitaNeural",     # 男性・明瞭でビジネス向け
    "aoi": "ja-JP-AoiNeural",         # 女性・明るく柔らかい声
    "daichi": "ja-JP-DaichiNeural",   # 男性・深みのあるナレーション向け
}

DEFAULT_VOICE_KEY = "nanami"
DEFAULT_SLIDE_WIDTH = 1920
DEFAULT_SLIDE_HEIGHT = 1080
DEFAULT_FPS = 30
DEFAULT_PAUSE_SECONDS = 0.8

# ==============================================================================
# ロガー設定
# ==============================================================================
LOGGER_NAME = "pptx_voice_pipeline"


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """タイムスタンプ付きの構造化ロガーを初期化する."""
    logger = logging.getLogger(LOGGER_NAME)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(
            fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    logger.setLevel(level)
    logger.propagate = False
    return logger


# ==============================================================================
# 例外・設定・データモデル
# ==============================================================================
class PipelineError(Exception):
    """パイプライン処理全体の基底例外."""


class InputValidationError(PipelineError):
    """入力ファイルやパラメータのバリデーションエラー."""


class SlideExportError(PipelineError):
    """PowerPoint COM によるスライドエクスポートエラー."""


class PipelineCancelledError(PipelineError):
    """ユーザーによるパイプライン中止."""


@dataclass(frozen=True)
class PipelineConfig:
    """パイプラインの動作設定."""

    voice: str = DEFAULT_VOICE_KEY
    slide_width: int = DEFAULT_SLIDE_WIDTH
    slide_height: int = DEFAULT_SLIDE_HEIGHT
    fps: int = DEFAULT_FPS
    pause_seconds: float = DEFAULT_PAUSE_SECONDS
    empty_note_template: str = "スライド {slide_num} の説明です。"
    max_retries: int = 3
    retry_base_delay: float = 1.0
    circuit_breaker_threshold: int = 3
    fallback_silence_seconds: float = 3.0
    tts_rate: str = "+0%"
    tts_volume: str = "+0%"
    tts_pitch: str = "+0Hz"
    voicevox_url: str = "http://127.0.0.1:50021"
    voicevox_speed: float = 1.0
    voicevox_pitch: float = 0.0
    voicevox_intonation: float = 1.0
    voicevox_volume: float = 1.0
    enable_pronunciation_dict: bool = True
    pronunciation_dict_path: Path | None = None
    ai_reading_enabled: bool = False
    ai_reading_base_url: str = DEFAULT_AI_BASE_URL
    ai_reading_model: str = DEFAULT_AI_MODEL
    ai_reading_api_key: str = ""

    def resolve_voice_id(self) -> str:
        """設定された話者キーまたは Voice ID を解決する."""
        raw = self.voice.strip()
        key = raw.lower()
        if key in VOICE_PRESETS:
            return VOICE_PRESETS[key]
        if key.startswith("sapi:") and len(raw) > 5:
            return raw
        if key.startswith("winrt:") and len(raw) > 6:
            return raw
        if key.startswith("voicevox:"):
            style_part = raw.split(":", 1)[1].strip()
            if style_part.isdigit():
                return f"voicevox:{int(style_part)}"
            raise InputValidationError(
                f"VOICEVOX 話者 ID は整数である必要があります: '{self.voice}'"
            )
        if "-" in raw:
            return raw
        available = ", ".join(sorted(VOICE_PRESETS.keys()))
        raise InputValidationError(
            f"未知の話者キー '{self.voice}' です。利用可能: {available}"
        )

    def resolve_voicevox_url(self) -> str:
        """VOICEVOX Engine URL を検証して正規化する."""
        url = self.voicevox_url.strip().rstrip("/")
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme not in ("http", "https") or not parsed.netloc:
            raise InputValidationError(
                f"不正な VOICEVOX Engine URL です: '{self.voicevox_url}'"
            )
        return url


@dataclass
class SlideAsset:
    """各スライドの中間成果物情報."""

    index: int
    image_path: Path
    note_text: str
    audio_path: Path | None = None
    is_fallback_audio: bool = False


# ==============================================================================
# サーキットブレイカー
# ==============================================================================
class CircuitBreaker:
    """連続失敗時に外部 API 呼び出しを遮断するサーキットブレイカー."""

    def __init__(self, failure_threshold: int = 3) -> None:
        if failure_threshold < 1:
            raise ValueError("failure_threshold must be >= 1")
        self.failure_threshold = failure_threshold
        self.consecutive_failures = 0
        self.is_open = False

    def record_success(self) -> None:
        """成功を記録し、連続失敗カウンタをリセットする."""
        self.consecutive_failures = 0
        self.is_open = False

    def record_failure(self) -> bool:
        """失敗を記録し、閾値に達した場合はオープン状態へ遷移する."""
        self.consecutive_failures += 1
        if self.consecutive_failures >= self.failure_threshold:
            self.is_open = True
        return self.is_open


# ==============================================================================
# 入力バリデーション & テキストサニタイジング
# ==============================================================================
_EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001F9FF"
    "\U0001FA00-\U0001FAFF"
    "\U00002600-\U000027BF"
    "\U0000FE00-\U0000FE0F"
    "\U0000200D"
    "]+",
    flags=re.UNICODE,
)
_BULLET_LINE_PATTERN = re.compile(r"^[\s・•●○■□◆◇▪▫‣⁃\-*+]+")
_DECORATIVE_SYMBOLS_PATTERN = re.compile(r"[★☆※†‡§¶◆◇■□●○▲△▼▽▶▷◀◁→←↑↓⇒⇐]")
_MARKDOWN_SYMBOLS_PATTERN = re.compile(r"[`*_~#>|]")


def validate_pptx_path(
    raw_path: str | Path,
    allowed_base_dir: Path | None = None,
) -> Path:
    """入力 .pptx ファイルのパスと形式を厳格に検証する."""
    raw_str = str(raw_path).strip()
    if not raw_str:
        raise InputValidationError("入力ファイルパスが空です。")

    path_obj = Path(raw_str)
    if ".." in path_obj.parts:
        raise InputValidationError(
            f"不正な相対パス指定 ('..') が含まれています: {raw_path}"
        )

    try:
        resolved = path_obj.resolve(strict=True)
    except FileNotFoundError as exc:
        raise InputValidationError(
            f"入力ファイルが存在しません: {raw_path}"
        ) from exc
    except OSError as exc:
        raise InputValidationError(
            f"入力パスの解決に失敗しました: {raw_path}"
        ) from exc

    if allowed_base_dir is not None:
        base_resolved = allowed_base_dir.resolve()
        if not resolved.is_relative_to(base_resolved):
            raise InputValidationError(
                f"許可されたディレクトリ外のファイルは指定できません: {resolved}"
            )

    if not resolved.is_file():
        raise InputValidationError(
            f"指定されたパスは通常ファイルではありません: {resolved}"
        )

    if resolved.suffix.lower() != ".pptx":
        raise InputValidationError(
            f"拡張子は .pptx である必要があります: {resolved.name}"
        )

    return resolved


def validate_output_path(raw_path: str | Path) -> Path:
    """出力 .mp4 ファイルのパスを検証し、親ディレクトリの存在を確認する."""
    raw_str = str(raw_path).strip()
    if not raw_str:
        raise InputValidationError("出力ファイルパスが空です。")

    path_obj = Path(raw_str)
    if ".." in path_obj.parts:
        raise InputValidationError(
            f"不正な相対パス指定 ('..') が含まれています: {raw_path}"
        )

    resolved = path_obj.resolve()
    if resolved.suffix.lower() != ".mp4":
        raise InputValidationError(
            f"出力ファイルの拡張子は .mp4 である必要があります: {resolved.name}"
        )

    if not resolved.parent.exists():
        raise InputValidationError(
            f"出力先ディレクトリが存在しません: {resolved.parent}"
        )

    return resolved


def sanitize_speaker_notes(
    raw_text: str | None,
    slide_num: int,
    empty_template: str = "スライド {slide_num} の説明です。",
) -> str:
    """発表者ノートを音声合成に適したプレーンテキストに正規化する."""
    if not raw_text:
        return empty_template.format(slide_num=slide_num)

    # Unicode 正規化 (全角英数・半角カナ等の揺らぎを統一)
    text = unicodedata.normalize("NFKC", raw_text)
    # 制御文字 (改行・タブ以外) を除去
    text = "".join(
        ch for ch in text
        if ch in ("\n", "\r", "\t") or unicodedata.category(ch)[0] != "C"
    )
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u3000", " ").replace("\t", " ")

    # 絵文字・装飾記号・Markdown 記号を除去
    text = _EMOJI_PATTERN.sub("", text)
    text = _DECORATIVE_SYMBOLS_PATTERN.sub("", text)
    text = _MARKDOWN_SYMBOLS_PATTERN.sub("", text)

    cleaned_lines: list[str] = []
    for line in text.split("\n"):
        stripped = _BULLET_LINE_PATTERN.sub("", line).strip()
        stripped = re.sub(r"\s+", " ", stripped)
        if not stripped:
            continue
        # 文末に句読点がなければ読点/句点を補い、自然な間を作る
        if stripped[-1] not in ("。", "．", "、", "！", "？", "!", "?", "."):
            stripped += "。"
        cleaned_lines.append(stripped)

    normalized = " ".join(cleaned_lines).strip()
    if not normalized:
        return empty_template.format(slide_num=slide_num)

    return normalized


# ==============================================================================
# 補助ユーティリティ (無音生成 & MoviePy 互換レイヤ)
# ==============================================================================
def create_silent_wav(
    output_path: Path,
    duration_seconds: float,
    sample_rate: int = 24000,
) -> Path:
    """外部ライブラリ不要で無音 WAV ファイルを生成する."""
    num_frames = max(1, int(sample_rate * max(0.1, duration_seconds)))
    silence_frame = struct.pack("<h", 0)
    with wave.open(str(output_path), "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        wf.writeframes(silence_frame * num_frames)
    return output_path


def _import_moviepy() -> tuple[Any, Any, Any]:
    """MoviePy v2.x と v1.x の両方に対応する互換インポート."""
    try:
        from moviepy import (  # type: ignore[import-not-found]
            AudioFileClip,
            ImageClip,
            concatenate_audioclips,
            concatenate_videoclips,
        )
        return ImageClip, AudioFileClip, (concatenate_videoclips, concatenate_audioclips)
    except ImportError:
        from moviepy.editor import (  # type: ignore[import-not-found,no-redef]
            AudioFileClip,
            ImageClip,
            concatenate_audioclips,
            concatenate_videoclips,
        )
        return ImageClip, AudioFileClip, (concatenate_videoclips, concatenate_audioclips)


def _with_duration(clip: Any, duration: float) -> Any:
    """MoviePy v1/v2 互換で duration を設定する."""
    if hasattr(clip, "with_duration"):
        return clip.with_duration(duration)
    return clip.set_duration(duration)


def _with_audio(clip: Any, audio: Any) -> Any:
    """MoviePy v1/v2 互換で audio を設定する."""
    if hasattr(clip, "with_audio"):
        return clip.with_audio(audio)
    return clip.set_audio(audio)


# ==============================================================================
# メインパイプラインクラス
# ==============================================================================
class PowerPointVideoPipeline:
    """PowerPoint 発表者ノートから MP4 動画を自動生成するパイプライン."""

    def __init__(
        self,
        config: PipelineConfig | None = None,
        logger: logging.Logger | None = None,
        cancel_event: threading.Event | None = None,
    ) -> None:
        self.config = config or PipelineConfig()
        self.logger = logger or setup_logging()
        self.voice_id = self.config.resolve_voice_id()
        self.cancel_event = cancel_event

    def _check_cancelled(self) -> None:
        """中止要求が発行されていれば PipelineCancelledError を送出する."""
        if self.cancel_event is not None and self.cancel_event.is_set():
            raise PipelineCancelledError("処理がユーザーにより中止されました。")

    def run(self, input_pptx: str | Path, output_mp4: str | Path) -> Path:
        """全3フェーズを実行し、MP4 動画を生成する."""
        pptx_path = validate_pptx_path(input_pptx)
        mp4_path = validate_output_path(output_mp4)

        self.logger.info("パイプライン開始: input=%s, output=%s, voice=%s",
                         pptx_path, mp4_path, self.voice_id)

        with tempfile.TemporaryDirectory(prefix="pptx_voice_") as tmp_dir_str:
            work_dir = Path(tmp_dir_str)
            self.logger.info("一時作業ディレクトリを作成しました: %s", work_dir)

            # フェーズ1: 同期処理 (PowerPoint COM で画像出力 + python-pptx でノート取得)
            assets = self._phase1_export_slides_and_notes(pptx_path, work_dir)
            if not assets:
                raise PipelineError("処理対象のスライドが見つかりませんでした。")

            # フェーズ2: 非同期処理 (edge-tts で全スライド音声を一括生成)
            asyncio.run(self._phase2_synthesize_all_audio(assets, work_dir))

            # フェーズ3: 動画・音声の同期結合と MP4 書き出し
            self._phase3_compose_video(assets, mp4_path, work_dir)

        self.logger.info("パイプライン正常終了: %s", mp4_path)
        return mp4_path

    # --------------------------------------------------------------------------
    # フェーズ 1: PowerPoint COM 制御 & ノート抽出
    # --------------------------------------------------------------------------
    def _phase1_export_slides_and_notes(
        self,
        pptx_path: Path,
        work_dir: Path,
    ) -> list[SlideAsset]:
        """COM (STA) を安全に隔離し、全スライドの高解像度 PNG とノートを取得する."""
        self.logger.info("[フェーズ1] スライド画像エクスポートとノート抽出を開始します。")
        notes = self._extract_notes_via_python_pptx(pptx_path)
        image_paths = self._export_slides_via_com(pptx_path, work_dir)

        if len(notes) != len(image_paths):
            self.logger.warning(
                "スライド数に差異があります (python-pptx=%d, COM=%d)。最小件数で同期します。",
                len(notes), len(image_paths),
            )

        count = min(len(notes), len(image_paths))
        assets: list[SlideAsset] = []
        for idx in range(count):
            slide_num = idx + 1
            sanitized = sanitize_speaker_notes(
                notes[idx],
                slide_num=slide_num,
                empty_template=self.config.empty_note_template,
            )
            assets.append(
                SlideAsset(
                    index=slide_num,
                    image_path=image_paths[idx],
                    note_text=sanitized,
                )
            )
            self.logger.info("スライド %d 準備完了 (ノート文字数=%d)", slide_num, len(sanitized))

        self.logger.info("[フェーズ1] 完了 (%d スライド)", len(assets))
        return assets

    def _extract_notes_via_python_pptx(self, pptx_path: Path) -> list[str]:
        """python-pptx を用いて各スライドの発表者ノートを読み込む."""
        from pptx import Presentation  # type: ignore[import-not-found]

        prs = Presentation(str(pptx_path))
        notes: list[str] = []
        for slide in prs.slides:
            note_text = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                note_text = slide.notes_slide.notes_text_frame.text or ""
            notes.append(note_text)
        return notes

    def _export_slides_via_com(self, pptx_path: Path, work_dir: Path) -> list[Path]:
        """PowerPoint COM オートメーションで全スライドを高解像度 PNG 出力する."""
        import comtypes  # type: ignore[import-not-found]
        import comtypes.client  # type: ignore[import-not-found]

        comtypes.CoInitialize()
        ppt_app: Any = None
        presentation: Any = None
        exported_paths: list[Path] = []

        try:
            ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
            # ダイアログ抑止 (ppAlertsNone = 1)
            try:
                ppt_app.DisplayAlerts = 1
            except Exception as alert_exc:  # noqa: BLE001
                self.logger.debug("DisplayAlerts 設定をスキップしました: %s", alert_exc)

            # WithWindow=0 (msoFalse) でバックグラウンドオープン
            presentation = ppt_app.Presentations.Open(
                str(pptx_path),
                ReadOnly=1,
                Untitled=0,
                WithWindow=0,
            )

            slide_count = int(presentation.Slides.Count)
            self.logger.info(
                "PowerPoint COM で %d 枚のスライドを検出しました (%dx%d PNG)。",
                slide_count, self.config.slide_width, self.config.slide_height,
            )

            for i in range(1, slide_count + 1):
                self._check_cancelled()
                out_png = work_dir / f"slide_{i:04d}.png"
                slide = presentation.Slides(i)
                slide.Export(
                    str(out_png),
                    "PNG",
                    self.config.slide_width,
                    self.config.slide_height,
                )
                if not out_png.exists() or out_png.stat().st_size == 0:
                    raise SlideExportError(f"スライド {i} の画像出力に失敗しました: {out_png}")
                exported_paths.append(out_png)

        except PipelineCancelledError:
            raise
        except Exception as exc:
            raise SlideExportError(f"PowerPoint COM 処理中にエラーが発生しました: {exc}") from exc
        finally:
            # ゾンビプロセス (POWERPNT.EXE) 残存を防ぐための確実な解放処理
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception as close_exc:  # noqa: BLE001
                    self.logger.warning("presentation.Close() 失敗: %s", close_exc)
                del presentation

            if ppt_app is not None:
                try:
                    ppt_app.Quit()
                except Exception as quit_exc:  # noqa: BLE001
                    self.logger.warning("ppt_app.Quit() 失敗: %s", quit_exc)
                del ppt_app

            gc.collect()
            comtypes.CoUninitialize()
            self.logger.info("PowerPoint COM リソースを解放しました。")

        return exported_paths

    # --------------------------------------------------------------------------
    # フェーズ 2: edge-tts 非同期音声合成 (リトライ & サーキットブレイカー)
    # --------------------------------------------------------------------------
    async def _phase2_synthesize_all_audio(
        self,
        assets: list[SlideAsset],
        work_dir: Path,
    ) -> None:
        """全スライドの音声を生成する (COM 解放後に実行)."""
        voice_lower = self.voice_id.lower()
        is_sapi = voice_lower.startswith("sapi:")
        is_winrt = voice_lower.startswith("winrt:")
        is_voicevox = voice_lower.startswith("voicevox:")
        engine_label = (
            "VOICEVOX Engine" if is_voicevox
            else "Windows OneCore (WinRT)" if is_winrt
            else "Windows SAPI5" if is_sapi
            else "edge-tts"
        )
        self.logger.info("[フェーズ2] %s による音声合成を開始します。", engine_label)
        breaker = CircuitBreaker(failure_threshold=self.config.circuit_breaker_threshold)
        dict_mapping = (
            load_pronunciation_dict(self.config.pronunciation_dict_path)
            if self.config.enable_pronunciation_dict
            else None
        )
        reading_cache: dict[str, str] = {}

        for asset in assets:
            self._check_cancelled()
            ext = "wav" if (is_sapi or is_winrt or is_voicevox) else "mp3"
            out_audio = work_dir / f"audio_{asset.index:04d}.{ext}"

            if breaker.is_open:
                self.logger.error(
                    "サーキットブレイカーがオープン状態です。スライド %d に無音フォールバックを適用します。",
                    asset.index,
                )
                self._apply_silent_fallback(asset, work_dir)
                continue

            tts_text = prepare_tts_text(
                asset.note_text,
                dict_mapping=dict_mapping,
                ai_enabled=self.config.ai_reading_enabled,
                ai_base_url=self.config.ai_reading_base_url,
                ai_model=self.config.ai_reading_model,
                ai_api_key=self.config.ai_reading_api_key,
                cache=reading_cache,
            )

            if is_voicevox:
                speaker_id = int(self.voice_id.split(":", 1)[1])
                success = self._synthesize_voicevox_wav(
                    text=tts_text,
                    output_path=out_audio,
                    speaker_id=speaker_id,
                    slide_num=asset.index,
                )
            elif is_winrt:
                winrt_voice_name = self.voice_id.split(":", 1)[1].strip()
                success = self._synthesize_winrt_wav(
                    text=tts_text,
                    output_path=out_audio,
                    voice_name=winrt_voice_name,
                    slide_num=asset.index,
                )
            elif is_sapi:
                sapi_voice_name = self.voice_id.split(":", 1)[1].strip()
                success = self._synthesize_sapi_wav(
                    text=tts_text,
                    output_path=out_audio,
                    voice_name=sapi_voice_name,
                    slide_num=asset.index,
                )
            else:
                success = await self._synthesize_with_retry(
                    text=tts_text,
                    output_path=out_audio,
                    slide_num=asset.index,
                )

            if success:
                breaker.record_success()
                asset.audio_path = out_audio
            else:
                opened = breaker.record_failure()
                self.logger.error(
                    "スライド %d の音声合成に失敗しました (連続失敗=%d)。無音フォールバックを適用します。",
                    asset.index, breaker.consecutive_failures,
                )
                if opened:
                    self.logger.error(
                        "連続失敗数が閾値 (%d) に達したため、サーキットブレイカーをオープンしました。",
                        self.config.circuit_breaker_threshold,
                    )
                self._apply_silent_fallback(asset, work_dir)

        self.logger.info("[フェーズ2] 完了")

    def _synthesize_sapi_wav(
        self,
        text: str,
        output_path: Path,
        voice_name: str,
        slide_num: int,
    ) -> bool:
        """Windows 標準 SAPI5 (System.Speech) を用いて WAV を生成する."""
        ps_script = (
            "Add-Type -AssemblyName System.Speech; "
            "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            "try { "
            "$s.SelectVoice($env:PPTX_TTS_VOICE); "
            "$s.SetOutputToWaveFile($env:PPTX_TTS_OUT); "
            "$s.Speak($env:PPTX_TTS_TEXT); "
            "} finally { $s.Dispose() }"
        )
        env = os.environ.copy()
        env["PPTX_TTS_VOICE"] = voice_name
        env["PPTX_TTS_OUT"] = str(output_path)
        env["PPTX_TTS_TEXT"] = text

        try:
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                env=env,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
            if proc.returncode == 0 and output_path.exists() and output_path.stat().st_size > 44:
                self.logger.info("スライド %d SAPI5 音声生成成功 (%s)", slide_num, voice_name)
                return True
            self.logger.error(
                "スライド %d SAPI5 音声生成失敗 (code=%d): %s",
                slide_num, proc.returncode, proc.stderr.strip(),
            )
            return False
        except Exception as exc:  # noqa: BLE001
            self.logger.error("スライド %d SAPI5 実行エラー: %s", slide_num, exc)
            return False

    def _synthesize_winrt_wav(
        self,
        text: str,
        output_path: Path,
        voice_name: str,
        slide_num: int,
    ) -> bool:
        """Windows OneCore (WinRT SpeechSynthesizer) を用いて WAV を生成する."""
        ps_script = (
            "Add-Type -AssemblyName System.Runtime.WindowsRuntime; "
            "$null = [Windows.Media.SpeechSynthesis.SpeechSynthesizer, Windows.Media.SpeechSynthesis, ContentType=WindowsRuntime]; "
            "$null = [Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType=WindowsRuntime]; "
            "$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | "
            "Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]; "
            "Function Await($WinRtTask, $ResultType) { "
            "$m = $asTaskGeneric.MakeGenericMethod($ResultType); "
            "$t = $m.Invoke($null, @($WinRtTask)); $t.Wait(-1) | Out-Null; $t.Result }; "
            "$s = New-Object Windows.Media.SpeechSynthesis.SpeechSynthesizer; "
            "try { "
            "$v = [Windows.Media.SpeechSynthesis.SpeechSynthesizer]::AllVoices | "
            "Where-Object { $_.DisplayName -eq $env:PPTX_TTS_VOICE } | Select-Object -First 1; "
            "if ($null -eq $v) { throw ('Voice not found: ' + $env:PPTX_TTS_VOICE) }; "
            "$s.Voice = $v; "
            "$stream = Await ($s.SynthesizeTextToStreamAsync($env:PPTX_TTS_TEXT)) ([Windows.Media.SpeechSynthesis.SpeechSynthesisStream]); "
            "$reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0)); "
            "$null = Await ($reader.LoadAsync([uint32]$stream.Size)) ([uint32]); "
            "$bytes = New-Object byte[] $stream.Size; "
            "$reader.ReadBytes($bytes); "
            "[System.IO.File]::WriteAllBytes($env:PPTX_TTS_OUT, $bytes); "
            "$reader.Dispose(); $stream.Dispose(); "
            "} finally { $s.Dispose() }"
        )
        env = os.environ.copy()
        env["PPTX_TTS_VOICE"] = voice_name
        env["PPTX_TTS_OUT"] = str(output_path)
        env["PPTX_TTS_TEXT"] = text

        try:
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                env=env,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
            if proc.returncode == 0 and output_path.exists() and output_path.stat().st_size > 44:
                self.logger.info("スライド %d OneCore 音声生成成功 (%s)", slide_num, voice_name)
                return True
            self.logger.error(
                "スライド %d OneCore 音声生成失敗 (code=%d): %s",
                slide_num, proc.returncode, proc.stderr.strip(),
            )
            return False
        except Exception as exc:  # noqa: BLE001
            self.logger.error("スライド %d OneCore 実行エラー: %s", slide_num, exc)
            return False

    def _synthesize_voicevox_wav(
        self,
        text: str,
        output_path: Path,
        speaker_id: int,
        slide_num: int,
    ) -> bool:
        """VOICEVOX Engine REST API (audio_query -> synthesis) で WAV を生成する."""
        base_url = self.config.resolve_voicevox_url()
        query_params = urllib.parse.urlencode({"text": text, "speaker": speaker_id})
        query_url = f"{base_url}/audio_query?{query_params}"
        synth_url = f"{base_url}/synthesis?speaker={speaker_id}"

        try:
            req_query = urllib.request.Request(query_url, method="POST", data=b"")
            with urllib.request.urlopen(req_query, timeout=30) as resp:
                audio_query = json.loads(resp.read().decode("utf-8"))

            audio_query["speedScale"] = float(self.config.voicevox_speed)
            audio_query["pitchScale"] = float(self.config.voicevox_pitch)
            audio_query["intonationScale"] = float(self.config.voicevox_intonation)
            audio_query["volumeScale"] = float(self.config.voicevox_volume)

            req_synth = urllib.request.Request(
                synth_url,
                method="POST",
                data=json.dumps(audio_query, ensure_ascii=False).encode("utf-8"),
                headers={"Content-Type": "application/json"},
            )
            with urllib.request.urlopen(req_synth, timeout=60) as resp:
                wav_bytes = resp.read()

            if len(wav_bytes) <= 44:
                raise RuntimeError("VOICEVOX から有効な WAV データが返されませんでした。")

            output_path.write_bytes(wav_bytes)
            self.logger.info("スライド %d VOICEVOX 音声生成成功 (speaker=%d)", slide_num, speaker_id)
            return True
        except Exception as exc:  # noqa: BLE001
            self.logger.error("スライド %d VOICEVOX 音声生成エラー: %s", slide_num, exc)
            return False

    async def _synthesize_with_retry(
        self,
        text: str,
        output_path: Path,
        slide_num: int,
    ) -> bool:
        """エクスポネンシャル・バックオフ付きで edge-tts を呼び出す."""
        import edge_tts  # type: ignore[import-not-found]

        total_attempts = self.config.max_retries + 1
        for attempt in range(total_attempts):
            try:
                communicate = edge_tts.Communicate(
                    text=text,
                    voice=self.voice_id,
                    rate=self.config.tts_rate,
                    volume=self.config.tts_volume,
                    pitch=self.config.tts_pitch,
                )
                await asyncio.wait_for(communicate.save(str(output_path)), timeout=60.0)

                if output_path.exists() and output_path.stat().st_size > 0:
                    self.logger.info("スライド %d 音声生成成功", slide_num)
                    return True

                raise RuntimeError("生成された音声ファイルが空です。")

            except Exception as exc:  # noqa: BLE001
                if attempt < self.config.max_retries:
                    delay = self.config.retry_base_delay * (2 ** attempt) + random.uniform(0.0, 0.3)
                    self.logger.warning(
                        "スライド %d 音声生成リトライ (%d/%d) %.2f秒後に再試行: %s",
                        slide_num, attempt + 1, self.config.max_retries, delay, exc,
                    )
                    await asyncio.sleep(delay)
                else:
                    self.logger.error(
                        "スライド %d 音声生成が最大リトライ回数 (%d) を超えて失敗しました: %s",
                        slide_num, self.config.max_retries, exc,
                    )
        return False

    def _apply_silent_fallback(self, asset: SlideAsset, work_dir: Path) -> None:
        """音声合成失敗時の安全な無音プレースホルダーを割り当てる."""
        fallback_wav = work_dir / f"fallback_{asset.index:04d}.wav"
        create_silent_wav(fallback_wav, duration_seconds=self.config.fallback_silence_seconds)
        asset.audio_path = fallback_wav
        asset.is_fallback_audio = True

    # --------------------------------------------------------------------------
    # フェーズ 3: 動画・音声の同期結合と MP4 書き出し
    # --------------------------------------------------------------------------
    def _phase3_compose_video(
        self,
        assets: list[SlideAsset],
        output_mp4: Path,
        work_dir: Path,
    ) -> None:
        """スライド画像と音声を動的尺合わせ + 無音パディングで高速結合し MP4 出力する."""
        import imageio_ffmpeg  # type: ignore[import-not-found]

        self.logger.info("[フェーズ3] 動画と音声の結合・高速エンコードを開始します。")
        _, AudioFileClip, _ = _import_moviepy()
        ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()

        segment_paths: list[Path] = []
        pause_sec = max(0.0, float(self.config.pause_seconds))

        for asset in assets:
            self._check_cancelled()
            if asset.audio_path is None or not asset.audio_path.exists():
                raise PipelineError(f"スライド {asset.index} の音声ファイルが見つかりません。")

            speech_clip = AudioFileClip(str(asset.audio_path))
            try:
                speech_duration = float(speech_clip.duration)
            finally:
                speech_clip.close()

            total_duration = speech_duration + pause_sec
            seg_path = work_dir / f"segment_{asset.index:04d}.mp4"

            cmd = [
                ffmpeg_exe,
                "-y",
                "-loop", "1",
                "-framerate", "1",
                "-i", str(asset.image_path),
                "-i", str(asset.audio_path),
                "-c:v", "libx264",
                "-tune", "stillimage",
                "-preset", "veryfast",
                "-pix_fmt", "yuv420p",
                "-r", str(self.config.fps),
                "-c:a", "aac",
                "-ar", "44100",
                "-ac", "2",
            ]
            if pause_sec > 0:
                cmd.extend(["-af", f"apad=pad_dur={pause_sec:.3f}"])
            cmd.extend(["-t", f"{total_duration:.3f}", str(seg_path)])

            self._run_ffmpeg(cmd)
            segment_paths.append(seg_path)

            self.logger.info(
                "スライド %d エンコード完了 (尺=%.2f秒%s)",
                asset.index,
                total_duration,
                ", 無音フォールバック" if asset.is_fallback_audio else "",
            )

        self._check_cancelled()
        concat_list = work_dir / "concat.txt"
        concat_lines = [
            f"file '{seg.resolve().as_posix()}'" for seg in segment_paths
        ]
        concat_list.write_text("\n".join(concat_lines), encoding="utf-8")

        self.logger.info("MP4 セグメントを結合中 (libx264 / aac, fps=%d)...", self.config.fps)
        concat_cmd = [
            ffmpeg_exe,
            "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(concat_list),
            "-c", "copy",
            "-movflags", "+faststart",
            str(output_mp4),
        ]
        self._run_ffmpeg(concat_cmd)
        self.logger.info("[フェーズ3] 完了: %s", output_mp4)

    def _run_ffmpeg(self, cmd: list[str], timeout_seconds: float = 600.0) -> None:
        """中止イベントとタイムアウトを監視しながら ffmpeg を実行する (stderr パイプ詰まりを防止)."""
        import time

        full_cmd = [cmd[0], "-nostdin", "-loglevel", "error", *cmd[1:]]
        deadline = time.monotonic() + timeout_seconds
        with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as err_file:
            proc = subprocess.Popen(
                full_cmd,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=err_file,
            )
            while proc.poll() is None:
                if self.cancel_event is not None and self.cancel_event.is_set():
                    proc.terminate()
                    try:
                        proc.wait(timeout=3)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    raise PipelineCancelledError("処理がユーザーにより中止されました。")
                if time.monotonic() > deadline:
                    proc.terminate()
                    try:
                        proc.wait(timeout=3)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    raise PipelineError(
                        f"ffmpeg 処理がタイムアウト ({timeout_seconds:.0f}秒) しました。"
                    )
                time.sleep(0.1)

            if proc.returncode != 0:
                err_file.seek(0)
                err = err_file.read()
                raise PipelineError(
                    f"ffmpeg エンコードに失敗しました (code={proc.returncode}): {err[-500:]}"
                )


# ==============================================================================
# CLI エントリポイント
# ==============================================================================
def build_arg_parser() -> argparse.ArgumentParser:
    """コマンドライン引数パーサを構築する."""
    parser = argparse.ArgumentParser(
        description="PowerPoint (.pptx) の発表者ノートを音声化し、MP4 動画を自動生成します。"
    )
    parser.add_argument("input_pptx", type=Path, help="入力 .pptx ファイルのパス")
    parser.add_argument("output_mp4", type=Path, help="出力 .mp4 ファイルのパス")
    parser.add_argument(
        "--voice",
        default=DEFAULT_VOICE_KEY,
        help=f"話者キーまたは Voice ID (既定: {DEFAULT_VOICE_KEY} / 選択肢: {', '.join(VOICE_PRESETS)})",
    )
    parser.add_argument("--width", type=int, default=DEFAULT_SLIDE_WIDTH, help="出力幅 (px)")
    parser.add_argument("--height", type=int, default=DEFAULT_SLIDE_HEIGHT, help="出力高さ (px)")
    parser.add_argument("--fps", type=int, default=DEFAULT_FPS, help="動画 FPS")
    parser.add_argument(
        "--pause",
        type=float,
        default=DEFAULT_PAUSE_SECONDS,
        help="各スライド末尾の無音ポーズ秒数 (既定: 0.8)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI メイン関数."""
    logger = setup_logging()
    args = build_arg_parser().parse_args(argv)

    try:
        config = PipelineConfig(
            voice=args.voice,
            slide_width=args.width,
            slide_height=args.height,
            fps=args.fps,
            pause_seconds=args.pause,
        )
        pipeline = PowerPointVideoPipeline(config=config, logger=logger)
        pipeline.run(args.input_pptx, args.output_mp4)
        return 0
    except PipelineError as exc:
        logger.error("パイプラインエラー: %s", exc)
        return 1
    except Exception:
        logger.exception("予期しないエラーが発生しました")
        return 2


if __name__ == "__main__":
    sys.exit(main())
