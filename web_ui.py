"""PPTX Voice Studio Web UI サーバー.

標準ライブラリのみで動作するローカル Web UI サーバー。
`python web_ui.py` で起動し、ブラウザで http://127.0.0.1:8000 を開きます。
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import re
import shutil
import sys
import tempfile
import threading
import urllib.parse
import urllib.request
import uuid
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from email.parser import BytesParser
from email.policy import default as email_policy
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from loopback_recorder import WasapiLoopbackRecorder
from pptx_to_video import (
    DEFAULT_FPS,
    DEFAULT_PAUSE_SECONDS,
    DEFAULT_SLIDE_HEIGHT,
    DEFAULT_SLIDE_WIDTH,
    DEFAULT_VOICE_KEY,
    VOICE_PRESETS,
    PipelineCancelledError,
    PipelineConfig,
    PowerPointVideoPipeline,
    SlideAsset,
    sanitize_speaker_notes,
    setup_logging,
)
from pronunciation import (
    DEFAULT_AI_BASE_URL,
    DEFAULT_AI_MODEL,
    fetch_lmproxy_models,
    load_pronunciation_dict,
    prepare_tts_text,
    validate_openai_base_url,
)

MAX_UPLOAD_BYTES = 200 * 1024 * 1024  # 200 MB
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")
PROJECT_ROOT = Path(__file__).resolve().parent
STATIC_INDEX = PROJECT_ROOT / "web" / "index.html"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"

_VOICE_CACHE_LOCK = threading.Lock()
_VOICE_CACHE: dict[str, object] | None = None

_GENDER_JA: dict[str, str] = {
    "Female": "女性",
    "Male": "男性",
}


def fetch_japanese_voices(force_refresh: bool = False) -> dict[str, object]:
    """edge-tts から高品質な日本語専用話者一覧を取得する（失敗時はプリセットにフォールバック）."""
    global _VOICE_CACHE
    with _VOICE_CACHE_LOCK:
        if _VOICE_CACHE is not None and not force_refresh:
            return _VOICE_CACHE

    default_id = VOICE_PRESETS[DEFAULT_VOICE_KEY]
    try:
        import edge_tts  # type: ignore[import-not-found]

        raw_voices = asyncio.run(edge_tts.list_voices())
        ja_voices: list[dict[str, str]] = []
        for item in raw_voices:
            locale = str(item.get("Locale", ""))
            short_name = str(item.get("ShortName", "")).strip()
            if not short_name or not locale.lower().startswith("ja"):
                continue

            gender_raw = str(item.get("Gender", ""))
            gender_label = _GENDER_JA.get(gender_raw, gender_raw or "不明")
            display_name = short_name.split("-")[-1]
            ja_voices.append(
                {
                    "id": short_name,
                    "label": f"{display_name} ({gender_label}) — {short_name}",
                    "gender": gender_raw,
                    "locale": locale,
                }
            )

        ja_voices.sort(key=lambda v: (0 if v["id"] == default_id else 1, v["id"]))
        if ja_voices:
            result: dict[str, object] = {
                "voices": ja_voices,
                "default": default_id,
                "source": "edge-tts",
                "count": len(ja_voices),
            }
            with _VOICE_CACHE_LOCK:
                _VOICE_CACHE = result
            return result
    except Exception as exc:  # noqa: BLE001
        logging.getLogger("pptx_voice_pipeline").debug("edge-tts 話者取得失敗: %s", exc)

    fallback_voices = [
        {
            "id": voice_id,
            "label": f"{key} — {voice_id}",
            "gender": "",
            "locale": "ja-JP",
        }
        for key, voice_id in (
            ("nanami", VOICE_PRESETS["nanami"]),
            ("keita", VOICE_PRESETS["keita"]),
        )
    ]
    return {
        "voices": fallback_voices,
        "default": default_id,
        "source": "preset-fallback",
        "count": len(fallback_voices),
    }


def fetch_voicevox_speakers(base_url: str = "http://127.0.0.1:50021") -> dict[str, object]:
    """VOICEVOX Engine (/speakers) から話者とスタイル一覧を取得する."""
    url = PipelineConfig(voicevox_url=base_url).resolve_voicevox_url()
    req = urllib.request.Request(f"{url}/speakers", method="GET")
    with urllib.request.urlopen(req, timeout=5) as resp:
        raw_speakers = json.loads(resp.read().decode("utf-8"))

    speakers: list[dict[str, object]] = []
    for item in raw_speakers:
        styles = [
            {"id": int(s["id"]), "name": str(s["name"])}
            for s in item.get("styles", [])
            if "id" in s and "name" in s
        ]
        if styles:
            speakers.append(
                {
                    "name": str(item.get("name", "Unknown")),
                    "speaker_uuid": str(item.get("speaker_uuid", "")),
                    "styles": styles,
                }
            )
    return {"url": url, "speakers": speakers}


def synthesize_preview_audio(
    voice: str,
    raw_text: str,
    *,
    voicevox_url: str = "http://127.0.0.1:50021",
    voicevox_speed: float = 1.0,
    voicevox_pitch: float = 0.0,
    voicevox_intonation: float = 1.0,
    voicevox_volume: float = 1.0,
    enable_pronunciation_dict: bool = True,
    ai_reading_enabled: bool = False,
    ai_reading_base_url: str = DEFAULT_AI_BASE_URL,
    ai_reading_model: str = DEFAULT_AI_MODEL,
    ai_reading_api_key: str = "",
) -> tuple[bytes, str]:
    """UI 試聴用に単発テキストを指定話者で音声合成し、(bytes, content_type) を返す."""
    cleaned = sanitize_speaker_notes(raw_text, slide_num=1, empty_template="音声プレビューのテストです。")
    if len(cleaned) > 500:
        cleaned = cleaned[:500]

    dict_mapping = load_pronunciation_dict() if enable_pronunciation_dict else None
    cleaned = prepare_tts_text(
        cleaned,
        dict_mapping=dict_mapping,
        ai_enabled=ai_reading_enabled,
        ai_base_url=ai_reading_base_url,
        ai_model=ai_reading_model,
        ai_api_key=ai_reading_api_key,
    )

    config = PipelineConfig(
        voice=voice,
        voicevox_url=voicevox_url,
        voicevox_speed=voicevox_speed,
        voicevox_pitch=voicevox_pitch,
        voicevox_intonation=voicevox_intonation,
        voicevox_volume=voicevox_volume,
    )
    voice_id = config.resolve_voice_id()
    pipeline = PowerPointVideoPipeline(config=config)
    voice_lower = voice_id.lower()

    with tempfile.TemporaryDirectory(prefix="pptx_preview_") as tmp_dir:
        tmp = Path(tmp_dir)
        if voice_lower.startswith("voicevox:"):
            out_path = tmp / "preview.wav"
            speaker_id = int(voice_id.split(":", 1)[1].strip())
            if not pipeline._synthesize_voicevox_wav(cleaned, out_path, speaker_id, slide_num=1):
                raise RuntimeError(f"VOICEVOX 音声合成に失敗しました: speaker_id={speaker_id}")
            return out_path.read_bytes(), "audio/wav"

        if voice_lower.startswith("winrt:"):
            out_path = tmp / "preview.wav"
            name = voice_id.split(":", 1)[1].strip()
            if not pipeline._synthesize_winrt_wav(cleaned, out_path, name, slide_num=1):
                raise RuntimeError(f"OneCore 音声合成に失敗しました: {name}")
            return out_path.read_bytes(), "audio/wav"

        if voice_lower.startswith("sapi:"):
            out_path = tmp / "preview.wav"
            name = voice_id.split(":", 1)[1].strip()
            if not pipeline._synthesize_sapi_wav(cleaned, out_path, name, slide_num=1):
                raise RuntimeError(f"SAPI5 音声合成に失敗しました: {name}")
            return out_path.read_bytes(), "audio/wav"

        out_path = tmp / "preview.mp3"
        ok = asyncio.run(pipeline._synthesize_with_retry(cleaned, out_path, slide_num=1))
        if not ok or not out_path.exists():
            raise RuntimeError(f"edge-tts 音声合成に失敗しました: {voice_id}")
        return out_path.read_bytes(), "audio/mpeg"


@dataclass
class JobState:
    """Web UI 上の変換ジョブ状態."""

    job_id: str
    work_dir: Path
    input_path: Path
    output_path: Path
    output_filename: str
    config: PipelineConfig
    mode: str = "edge-tts"  # edge-tts | edge-record
    status: str = "queued"  # queued | running | recording_needed | completed | failed
    phase: str = "待機中"
    progress: int = 0
    logs: list[str] = field(default_factory=list)
    error: str | None = None
    assets: list[SlideAsset] = field(default_factory=list)
    recorder: WasapiLoopbackRecorder | None = None
    recording_slide_index: int | None = None
    cancel_event: threading.Event = field(default_factory=threading.Event)
    lock: threading.Lock = field(default_factory=threading.Lock)


class JobLogHandler(logging.Handler):
    """ジョブ単位でログをメモリに収集し、フェーズ進捗を更新するハンドラ."""

    def __init__(self, job: JobState) -> None:
        super().__init__()
        self.job = job
        self.setFormatter(
            logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
        )

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        with self.job.lock:
            self.job.logs.append(msg)
            raw = record.getMessage()
            if "[フェーズ1]" in raw:
                self.job.phase = "フェーズ1: スライド画像エクスポート中"
                self.job.progress = 20
            elif "[フェーズ2]" in raw:
                self.job.phase = "フェーズ2: edge-tts 音声合成中"
                self.job.progress = 55
            elif "[フェーズ3]" in raw:
                self.job.phase = "フェーズ3: MP4 動画エンコード中"
                self.job.progress = 80


class JobManager:
    """ジョブの登録・直列実行 (COM 安全) ・状態管理を行う."""

    def __init__(
        self,
        base_temp_dir: Path,
        output_dir: Path = DEFAULT_OUTPUT_DIR,
    ) -> None:
        self.base_temp_dir = base_temp_dir
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.jobs: dict[str, JobState] = {}
        self._jobs_lock = threading.Lock()
        self._execution_lock = threading.Lock()
        self._server_logger = setup_logging()

    def _resolve_unique_output_path(self, safe_pptx_name: str) -> tuple[str, Path]:
        """output フォルダ内の重複しない出力ファイル名とパスを決定する."""
        stem = Path(safe_pptx_name).stem
        candidate_name = f"{stem}.mp4"
        candidate_path = self.output_dir / candidate_name
        if candidate_path.exists():
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            candidate_name = f"{stem}_{stamp}.mp4"
            candidate_path = self.output_dir / candidate_name
        return candidate_name, candidate_path

    def create_job(
        self,
        filename: str,
        file_bytes: bytes,
        config: PipelineConfig,
        mode: str = "edge-tts",
    ) -> JobState:
        """アップロードされた .pptx を検証・保存し、バックグラウンド実行を開始する."""
        safe_name = Path(filename).name
        if not safe_name.lower().endswith(".pptx"):
            raise ValueError("拡張子が .pptx のファイルのみアップロードできます。")
        if not file_bytes.startswith(b"PK\x03\x04"):
            raise ValueError("有効な .pptx (Office Open XML) ファイルではありません。")

        job_id = uuid.uuid4().hex
        job_dir = self.base_temp_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        input_path = job_dir / safe_name
        input_path.write_bytes(file_bytes)

        output_filename, output_path = self._resolve_unique_output_path(safe_name)

        job = JobState(
            job_id=job_id,
            work_dir=job_dir,
            input_path=input_path,
            output_path=output_path,
            output_filename=output_filename,
            config=config,
            mode=mode,
        )
        with self._jobs_lock:
            self.jobs[job_id] = job

        thread = threading.Thread(target=self._run_job, args=(job,), daemon=True)
        thread.start()
        return job

    def get_job(self, job_id: str) -> JobState | None:
        """ジョブ ID から状態を取得する."""
        if not JOB_ID_PATTERN.match(job_id):
            return None
        with self._jobs_lock:
            return self.jobs.get(job_id)

    def _run_job(self, job: JobState) -> None:
        job_logger = logging.getLogger(f"pptx_voice_job_{job.job_id}")
        job_logger.setLevel(logging.INFO)
        job_logger.propagate = False
        handler = JobLogHandler(job)
        job_logger.addHandler(handler)

        try:
            with self._execution_lock:
                if job.cancel_event.is_set():
                    raise PipelineCancelledError("処理がユーザーにより中止されました。")
                with job.lock:
                    job.status = "running"
                    job.phase = "パイプライン開始"
                    job.progress = 10

                pipeline = PowerPointVideoPipeline(
                    config=job.config,
                    logger=job_logger,
                    cancel_event=job.cancel_event,
                )
                if job.mode == "edge-record":
                    assets = pipeline._phase1_export_slides_and_notes(job.input_path, job.work_dir)
                    if not assets:
                        raise RuntimeError("処理対象のスライドが見つかりませんでした。")
                    with job.lock:
                        if job.cancel_event.is_set():
                            raise PipelineCancelledError("処理がユーザーにより中止されました。")
                        job.assets = assets
                        job.status = "recording_needed"
                        job.phase = "ブラウザ音声録音待機中"
                        job.progress = 30
                else:
                    pipeline.run(job.input_path, job.output_path)
                    with job.lock:
                        job.status = "completed"
                        job.phase = "完了"
                        job.progress = 100
        except PipelineCancelledError:
            self._cleanup_partial_output(job)
            with job.lock:
                job.status = "cancelled"
                job.phase = "中止されました"
                job.logs.append("処理が中止されました。")
        except Exception as exc:  # noqa: BLE001
            self._server_logger.error("ジョブ %s 失敗: %s", job.job_id, exc)
            with job.lock:
                job.status = "failed"
                job.phase = "失敗"
                job.error = str(exc)
        finally:
            job_logger.removeHandler(handler)

    def start_slide_recording(self, job: JobState, slide_index: int) -> None:
        """指定スライドの WASAPI ループバック録音を開始する."""
        with job.lock:
            if job.status != "recording_needed":
                raise RuntimeError("録音待機状態ではありません。")
            if job.recorder is not None:
                raise RuntimeError("既に別の録音が進行中です。")
            recorder = WasapiLoopbackRecorder()
            recorder.start()
            job.recorder = recorder
            job.recording_slide_index = slide_index
            job.phase = f"フェーズ2: スライド {slide_index} を録音中..."

    def stop_slide_recording(self, job: JobState, slide_index: int) -> dict[str, object]:
        """指定スライドの録音を停止して WAV 保存する."""
        with job.lock:
            recorder = job.recorder
            active_idx = job.recording_slide_index
            if recorder is None or active_idx != slide_index:
                raise RuntimeError("一致する録音が開始されていません。")
            job.recorder = None
            job.recording_slide_index = None

        wav_path = job.work_dir / f"audio_{slide_index:04d}.wav"
        recorder.stop_and_save(wav_path)

        with job.lock:
            for asset in job.assets:
                if asset.index == slide_index:
                    asset.audio_path = wav_path
                    break
            done_count = sum(1 for a in job.assets if a.audio_path is not None)
            total = max(1, len(job.assets))
            job.progress = 30 + int(45 * done_count / total)
            job.logs.append(f"録音完了: スライド {slide_index}/{total}")

        return {"slide_index": slide_index, "recorded": done_count, "total": len(job.assets)}

    def start_compose(self, job: JobState) -> None:
        """録音済み音声を用いてフェーズ3 (MP4 書き出し) をバックグラウンドで開始する."""
        with job.lock:
            if job.status != "recording_needed":
                raise RuntimeError("動画結合を開始できる状態ではありません。")
            job.status = "running"
            job.phase = "フェーズ3: MP4 動画エンコード中"
            job.progress = 80

        thread = threading.Thread(target=self._run_compose, args=(job,), daemon=True)
        thread.start()

    def _run_compose(self, job: JobState) -> None:
        job_logger = logging.getLogger(f"pptx_voice_job_{job.job_id}")
        job_logger.setLevel(logging.INFO)
        job_logger.propagate = False
        handler = JobLogHandler(job)
        job_logger.addHandler(handler)
        try:
            with self._execution_lock:
                if job.cancel_event.is_set():
                    raise PipelineCancelledError("処理がユーザーにより中止されました。")
                pipeline = PowerPointVideoPipeline(
                    config=job.config,
                    logger=job_logger,
                    cancel_event=job.cancel_event,
                )
                pipeline._phase3_compose_video(job.assets, job.output_path, job.work_dir)
                with job.lock:
                    job.status = "completed"
                    job.phase = "完了"
                    job.progress = 100
        except PipelineCancelledError:
            self._cleanup_partial_output(job)
            with job.lock:
                job.status = "cancelled"
                job.phase = "中止されました"
                job.logs.append("動画結合が中止されました。")
        except Exception as exc:  # noqa: BLE001
            self._server_logger.error("ジョブ %s 結合失敗: %s", job.job_id, exc)
            with job.lock:
                job.status = "failed"
                job.phase = "失敗"
                job.error = str(exc)
        finally:
            job_logger.removeHandler(handler)

    def _cleanup_partial_output(self, job: JobState) -> None:
        try:
            if job.output_path.exists():
                job.output_path.unlink()
        except OSError:
            pass

    def cancel_job(self, job: JobState) -> None:
        """実行中または録音中のジョブを中止する."""
        job.cancel_event.set()
        recorder: WasapiLoopbackRecorder | None = None
        with job.lock:
            recorder = job.recorder
            job.recorder = None
            job.recording_slide_index = None
            if job.status in ("queued", "running", "recording_needed"):
                job.status = "cancelled"
                job.phase = "中止されました"
                job.logs.append("ユーザーにより中止が要求されました。")

        if recorder is not None:
            try:
                recorder.stop_and_save(job.work_dir / "_cancelled.wav")
            except Exception as exc:  # noqa: BLE001
                self._server_logger.warning("中止時の録音停止で警告: %s", exc)
        self._cleanup_partial_output(job)


def parse_multipart_form(
    content_type: str,
    body: bytes,
) -> tuple[dict[str, str], tuple[str, bytes] | None]:
    """標準ライブラリ email.parser を用いて multipart/form-data を解析する."""
    header = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode()
    msg = BytesParser(policy=email_policy).parsebytes(header + body)

    fields: dict[str, str] = {}
    uploaded_file: tuple[str, bytes] | None = None

    if not msg.is_multipart():
        raise ValueError("multipart/form-data 形式ではありません。")

    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not isinstance(name, str) or not name:
            continue
        filename = part.get_filename()
        raw_payload = part.get_payload(decode=True)
        payload = raw_payload if isinstance(raw_payload, bytes) else b""
        if filename is not None:
            uploaded_file = (filename, payload)
        else:
            fields[name] = payload.decode("utf-8", errors="replace")

    return fields, uploaded_file


def make_handler(manager: JobManager) -> type[BaseHTTPRequestHandler]:
    """JobManager をバインドした HTTP リクエストハンドラを生成する."""

    class WebUIHandler(BaseHTTPRequestHandler):
        def log_message(self, format: str, *args: object) -> None:
            return  # コンソールのアクセスログノイズを抑制

        def _send_json(self, status: int, payload: dict[str, object]) -> None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self) -> None:
            path = self.path.split("?", 1)[0]

            if path == "/":
                if not STATIC_INDEX.exists():
                    self.send_error(HTTPStatus.NOT_FOUND, "index.html not found")
                    return
                html = STATIC_INDEX.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(html)))
                self.end_headers()
                self.wfile.write(html)
                return

            if path == "/api/voices":
                force_refresh = "refresh=1" in self.path
                self._send_json(
                    HTTPStatus.OK,
                    fetch_japanese_voices(force_refresh=force_refresh),
                )
                return

            if path == "/api/voicevox-speakers":
                query = urllib.parse.urlsplit(self.path).query
                params = urllib.parse.parse_qs(query)
                base_url = params.get("url", ["http://127.0.0.1:50021"])[0]
                try:
                    self._send_json(HTTPStatus.OK, fetch_voicevox_speakers(base_url))
                except Exception as exc:  # noqa: BLE001
                    self._send_json(HTTPStatus.BAD_GATEWAY, {"error": f"VOICEVOX Engine 接続エラー: {exc}"})
                return

            parts = [p for p in path.split("/") if p]
            if len(parts) >= 3 and parts[0] == "api" and parts[1] == "jobs":
                job_id = parts[2]
                job = manager.get_job(job_id)
                if job is None:
                    self._send_json(HTTPStatus.NOT_FOUND, {"error": "ジョブが見つかりません。"})
                    return

                if len(parts) == 3:
                    with job.lock:
                        self._send_json(
                            HTTPStatus.OK,
                            {
                                "job_id": job.job_id,
                                "status": job.status,
                                "phase": job.phase,
                                "progress": job.progress,
                                "logs": list(job.logs),
                                "error": job.error,
                                "output_filename": job.output_filename,
                                "output_path": str(job.output_path),
                                "mode": job.mode,
                                "slides": (
                                    [{"index": a.index, "note_text": a.note_text} for a in job.assets]
                                    if job.status == "recording_needed"
                                    else []
                                ),
                            },
                        )
                    return

                if len(parts) == 4 and parts[3] in ("video", "download"):
                    if job.status != "completed" or not job.output_path.exists():
                        self._send_json(HTTPStatus.CONFLICT, {"error": "動画はまだ準備できていません。"})
                        return
                    self._serve_mp4(job.output_path, job.output_filename, download=(parts[3] == "download"))
                    return

            self._send_json(HTTPStatus.NOT_FOUND, {"error": "Not Found"})

        def do_POST(self) -> None:
            path = self.path.split("?", 1)[0]
            if path == "/api/tts-preview":
                self._handle_tts_preview()
                return
            if path == "/api/lmproxy-models":
                self._handle_lmproxy_models()
                return

            parts = [p for p in path.split("/") if p]
            if len(parts) == 4 and parts[0] == "api" and parts[1] == "jobs":
                self._handle_record_action(parts[2], parts[3])
                return

            if path != "/api/jobs":
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "Not Found"})
                return

            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "不正な Content-Length です。"})
                return

            if content_length <= 0 or content_length > MAX_UPLOAD_BYTES:
                self._send_json(
                    HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                    {"error": "ファイルサイズが上限 (200 MB) を超えているか空です。"},
                )
                return

            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)

            try:
                fields, uploaded = parse_multipart_form(content_type, body)
                if uploaded is None:
                    raise ValueError(".pptx ファイルが添付されていません。")

                filename, file_bytes = uploaded
                mode = fields.get("mode", "edge-tts").strip()
                if mode not in ("edge-tts", "edge-record", "voicevox"):
                    raise ValueError(f"未知の生成方式です: {mode}")

                voice_value = fields.get("voice", DEFAULT_VOICE_KEY)
                ai_enabled = fields.get("ai_reading_enabled", "0").strip().lower() in ("1", "true", "on", "yes")
                ai_base_url = fields.get("ai_reading_base_url", DEFAULT_AI_BASE_URL).strip() or DEFAULT_AI_BASE_URL
                if ai_enabled:
                    ai_base_url = validate_openai_base_url(ai_base_url)

                config = PipelineConfig(
                    voice=DEFAULT_VOICE_KEY if mode == "edge-record" else voice_value,
                    slide_width=int(fields.get("width", str(DEFAULT_SLIDE_WIDTH))),
                    slide_height=int(fields.get("height", str(DEFAULT_SLIDE_HEIGHT))),
                    fps=int(fields.get("fps", str(DEFAULT_FPS))),
                    pause_seconds=float(fields.get("pause", str(DEFAULT_PAUSE_SECONDS))),
                    voicevox_url=fields.get("voicevox_url", "http://127.0.0.1:50021").strip() or "http://127.0.0.1:50021",
                    voicevox_speed=float(fields.get("voicevox_speed", "1.0")),
                    voicevox_pitch=float(fields.get("voicevox_pitch", "0.0")),
                    voicevox_intonation=float(fields.get("voicevox_intonation", "1.0")),
                    voicevox_volume=float(fields.get("voicevox_volume", "1.0")),
                    enable_pronunciation_dict=fields.get("enable_pronunciation_dict", "1").strip().lower()
                    in ("1", "true", "on", "yes"),
                    ai_reading_enabled=ai_enabled,
                    ai_reading_base_url=ai_base_url,
                    ai_reading_model=fields.get("ai_reading_model", DEFAULT_AI_MODEL).strip() or DEFAULT_AI_MODEL,
                    ai_reading_api_key=fields.get("ai_reading_api_key", "").strip(),
                )
                config.resolve_voice_id()
                if mode == "voicevox":
                    config.resolve_voicevox_url()
                job = manager.create_job(filename, file_bytes, config, mode=mode)
                self._send_json(HTTPStatus.ACCEPTED, {"job_id": job.job_id})
            except Exception as exc:  # noqa: BLE001
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})

        def _handle_tts_preview(self) -> None:
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "不正な Content-Length です。"})
                return
            if content_length <= 0 or content_length > 65536:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "リクエストサイズが無効です。"})
                return

            try:
                payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
                voice = str(payload.get("voice", DEFAULT_VOICE_KEY)).strip()
                text = str(payload.get("text", "")).strip()
                if not text:
                    raise ValueError("サンプルテキストを入力してください。")
                audio_bytes, content_type = synthesize_preview_audio(
                    voice,
                    text,
                    voicevox_url=str(payload.get("voicevox_url", "http://127.0.0.1:50021")).strip()
                    or "http://127.0.0.1:50021",
                    voicevox_speed=float(payload.get("voicevox_speed", 1.0)),
                    voicevox_pitch=float(payload.get("voicevox_pitch", 0.0)),
                    voicevox_intonation=float(payload.get("voicevox_intonation", 1.0)),
                    voicevox_volume=float(payload.get("voicevox_volume", 1.0)),
                    enable_pronunciation_dict=bool(payload.get("enable_pronunciation_dict", True)),
                    ai_reading_enabled=bool(payload.get("ai_reading_enabled", False)),
                    ai_reading_base_url=str(payload.get("ai_reading_base_url", DEFAULT_AI_BASE_URL)).strip()
                    or DEFAULT_AI_BASE_URL,
                    ai_reading_model=str(payload.get("ai_reading_model", DEFAULT_AI_MODEL)).strip()
                    or DEFAULT_AI_MODEL,
                    ai_reading_api_key=str(payload.get("ai_reading_api_key", "")).strip(),
                )
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(audio_bytes)))
                self.end_headers()
                self.wfile.write(audio_bytes)
            except Exception as exc:  # noqa: BLE001
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})

        def _handle_lmproxy_models(self) -> None:
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "不正な Content-Length です。"})
                return
            if content_length < 0 or content_length > 16384:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "リクエストサイズが無効です。"})
                return

            try:
                raw_body = self.rfile.read(content_length).decode("utf-8") if content_length > 0 else "{}"
                payload = json.loads(raw_body or "{}")
                base_url = str(payload.get("base_url", DEFAULT_AI_BASE_URL)).strip() or DEFAULT_AI_BASE_URL
                api_key = str(payload.get("api_key", "")).strip()
                self._send_json(
                    HTTPStatus.OK,
                    fetch_lmproxy_models(base_url, api_key=api_key),
                )
            except Exception as exc:  # noqa: BLE001
                self._send_json(HTTPStatus.BAD_GATEWAY, {"error": f"LMproxy 接続エラー: {exc}"})

        def _handle_record_action(self, job_id: str, action: str) -> None:
            job = manager.get_job(job_id)
            if job is None:
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "ジョブが見つかりません。"})
                return

            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                content_length = 0

            payload: dict[str, object] = {}
            if content_length > 0:
                if content_length > 65536:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "リクエストサイズ過大"})
                    return
                payload = json.loads(self.rfile.read(content_length).decode("utf-8"))

            try:
                if action == "record-start":
                    slide_idx = int(str(payload.get("slide_index", 0)))
                    manager.start_slide_recording(job, slide_idx)
                    self._send_json(HTTPStatus.OK, {"status": "recording", "slide_index": slide_idx})
                elif action == "record-stop":
                    slide_idx = int(str(payload.get("slide_index", 0)))
                    info = manager.stop_slide_recording(job, slide_idx)
                    self._send_json(HTTPStatus.OK, info)
                elif action == "compose":
                    manager.start_compose(job)
                    self._send_json(HTTPStatus.ACCEPTED, {"status": "composing"})
                elif action == "cancel":
                    manager.cancel_job(job)
                    self._send_json(HTTPStatus.OK, {"status": "cancelled"})
                else:
                    self._send_json(HTTPStatus.NOT_FOUND, {"error": "Not Found"})
            except Exception as exc:  # noqa: BLE001
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})

        def _serve_mp4(self, file_path: Path, filename: str, download: bool) -> None:
            file_size = file_path.stat().st_size
            range_header = self.headers.get("Range")

            start = 0
            end = file_size - 1
            status = HTTPStatus.OK

            if range_header and range_header.startswith("bytes="):
                spec = range_header[len("bytes="):].split(",")[0].strip()
                if "-" in spec:
                    s_str, e_str = spec.split("-", 1)
                    if s_str:
                        start = int(s_str)
                    if e_str:
                        end = min(int(e_str), file_size - 1)
                    if start <= end < file_size:
                        status = HTTPStatus.PARTIAL_CONTENT

            length = end - start + 1
            self.send_response(status)
            self.send_header("Content-Type", "video/mp4")
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Length", str(length))
            if status == HTTPStatus.PARTIAL_CONTENT:
                self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
            if download:
                self.send_header(
                    "Content-Disposition",
                    f'attachment; filename="{filename}"',
                )
            self.end_headers()

            with file_path.open("rb") as fh:
                fh.seek(start)
                remaining = length
                while remaining > 0:
                    chunk = fh.read(min(65536, remaining))
                    if not chunk:
                        break
                    self.wfile.write(chunk)
                    remaining -= len(chunk)

    return WebUIHandler


def main(argv: list[str] | None = None) -> int:
    """Web UI サーバーを起動する."""
    parser = argparse.ArgumentParser(description="PPTX Voice Studio Web UI サーバー")
    parser.add_argument("--host", default="127.0.0.1", help="バインドホスト (既定: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="ポート番号 (既定: 8000)")
    parser.add_argument("--no-browser", action="store_true", help="ブラウザを自動で開かない")
    args = parser.parse_args(argv)

    logger = setup_logging()
    temp_root = Path(tempfile.mkdtemp(prefix="pptx_voice_webui_"))
    DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    manager = JobManager(temp_root, output_dir=DEFAULT_OUTPUT_DIR)
    server = ThreadingHTTPServer((args.host, args.port), make_handler(manager))
    server.daemon_threads = True

    url = f"http://{args.host}:{args.port}"
    logger.info("Web UI サーバーを起動しました: %s (出力先: %s)", url, DEFAULT_OUTPUT_DIR)
    if not args.no_browser:
        webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("サーバーを停止しています...")
    finally:
        server.server_close()
        shutil.rmtree(temp_root, ignore_errors=True)
        logger.info("一時作業領域を削除しました。")

    return 0


if __name__ == "__main__":
    sys.exit(main())
