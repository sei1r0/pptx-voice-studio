"""pptx_to_video の単体テスト."""

from __future__ import annotations

import asyncio
import logging
import sys
import tempfile
import types
import unittest
import wave
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx_to_video import (
    CircuitBreaker,
    InputValidationError,
    PipelineConfig,
    PowerPointVideoPipeline,
    SlideAsset,
    SlideExportError,
    create_silent_wav,
    sanitize_speaker_notes,
    validate_output_path,
    validate_pptx_path,
)


class TestValidation(unittest.TestCase):
    """入力・出力パスバリデーションのテスト."""

    def test_validate_pptx_path_success(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "deck.pptx"
            pptx.write_bytes(b"PK\x03\x04dummy")
            resolved = validate_pptx_path(pptx)
            self.assertEqual(resolved, pptx.resolve())

    def test_validate_pptx_path_rejects_traversal(self) -> None:
        with self.assertRaises(InputValidationError):
            validate_pptx_path("../secret.pptx")

    def test_validate_pptx_path_rejects_wrong_extension(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            txt = Path(tmp) / "deck.txt"
            txt.write_text("hello", encoding="utf-8")
            with self.assertRaises(InputValidationError):
                validate_pptx_path(txt)

    def test_validate_pptx_path_rejects_missing_file(self) -> None:
        with self.assertRaises(InputValidationError):
            validate_pptx_path("nonexistent_file_12345.pptx")

    def test_validate_output_path_checks_extension_and_parent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "out.mp4"
            self.assertEqual(validate_output_path(out), out.resolve())
            with self.assertRaises(InputValidationError):
                validate_output_path(Path(tmp) / "out.avi")
            with self.assertRaises(InputValidationError):
                validate_output_path(Path(tmp) / "missing_dir" / "out.mp4")


class TestSanitizeNotes(unittest.TestCase):
    """発表者ノートサニタイジングのテスト."""

    def test_empty_and_whitespace_fallback(self) -> None:
        self.assertEqual(sanitize_speaker_notes("", 1), "スライド 1 の説明です。")
        self.assertEqual(sanitize_speaker_notes("   \n\u3000\t  ", 3), "スライド 3 の説明です。")
        self.assertEqual(sanitize_speaker_notes(None, 2), "スライド 2 の説明です。")

    def test_cleans_bullets_emojis_and_whitespace(self) -> None:
        raw = "・最初のポイント🎉\n- 二番目の\u3000ポイント★\n**強調**テキスト"
        cleaned = sanitize_speaker_notes(raw, 1)
        self.assertNotIn("🎉", cleaned)
        self.assertNotIn("★", cleaned)
        self.assertNotIn("**", cleaned)
        self.assertNotIn("\u3000", cleaned)
        self.assertIn("最初のポイント。", cleaned)
        self.assertIn("二番目の ポイント。", cleaned)
        self.assertIn("強調テキスト。", cleaned)


class TestVoiceAndCircuitBreaker(unittest.TestCase):
    """Voice ID 解決とサーキットブレイカーのテスト."""

    def test_voice_preset_resolution(self) -> None:
        self.assertEqual(PipelineConfig(voice="nanami").resolve_voice_id(), "ja-JP-NanamiNeural")
        self.assertEqual(PipelineConfig(voice="keita").resolve_voice_id(), "ja-JP-KeitaNeural")
        self.assertEqual(PipelineConfig(voice="en-US-AriaNeural").resolve_voice_id(), "en-US-AriaNeural")
        with self.assertRaises(InputValidationError):
            PipelineConfig(voice="unknown_voice").resolve_voice_id()

    def test_circuit_breaker_transitions(self) -> None:
        cb = CircuitBreaker(failure_threshold=2)
        self.assertFalse(cb.is_open)
        self.assertFalse(cb.record_failure())
        self.assertFalse(cb.is_open)
        self.assertTrue(cb.record_failure())
        self.assertTrue(cb.is_open)
        cb.record_success()
        self.assertFalse(cb.is_open)
        self.assertEqual(cb.consecutive_failures, 0)

    def test_create_silent_wav(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            wav_path = Path(tmp) / "silence.wav"
            create_silent_wav(wav_path, duration_seconds=1.0, sample_rate=16000)
            with wave.open(str(wav_path), "rb") as wf:
                self.assertEqual(wf.getnchannels(), 1)
                self.assertEqual(wf.getframerate(), 16000)
                self.assertEqual(wf.getnframes(), 16000)


class TestPipelinePhases(unittest.TestCase):
    """COM 解放と edge-tts リトライ/フォールバックのテスト."""

    def test_phase2_retry_and_circuit_breaker_fallback(self) -> None:
        config = PipelineConfig(
            max_retries=1,
            retry_base_delay=0.01,
            circuit_breaker_threshold=2,
            fallback_silence_seconds=0.5,
        )
        logger = logging.getLogger("test_pipeline")
        logger.addHandler(logging.NullHandler())
        pipeline = PowerPointVideoPipeline(config=config, logger=logger)

        fake_edge_tts = types.ModuleType("edge_tts")
        call_count = 0

        class FailingCommunicate:
            def __init__(self, **kwargs: object) -> None:
                pass

            async def save(self, path: str) -> None:
                nonlocal call_count
                call_count += 1
                raise ConnectionError("Simulated network failure")

        fake_edge_tts.Communicate = FailingCommunicate  # type: ignore[attr-defined]

        with tempfile.TemporaryDirectory() as tmp:
            work_dir = Path(tmp)
            assets = [
                SlideAsset(index=1, image_path=work_dir / "1.png", note_text="テスト1。"),
                SlideAsset(index=2, image_path=work_dir / "2.png", note_text="テスト2。"),
                SlideAsset(index=3, image_path=work_dir / "3.png", note_text="テスト3。"),
            ]
            with patch.dict(sys.modules, {"edge_tts": fake_edge_tts}):
                asyncio.run(pipeline._phase2_synthesize_all_audio(assets, work_dir))

            # スライド1,2で各2回(初回+リトライ1回)=4回試行後、サーキットブレイカーが開きスライド3は即フォールバック
            self.assertEqual(call_count, 4)
            for asset in assets:
                self.assertTrue(asset.is_fallback_audio)
                self.assertIsNotNone(asset.audio_path)
                assert asset.audio_path is not None
                self.assertTrue(asset.audio_path.exists())

    def test_com_quit_called_on_export_error(self) -> None:
        config = PipelineConfig()
        logger = logging.getLogger("test_com")
        logger.addHandler(logging.NullHandler())
        pipeline = PowerPointVideoPipeline(config=config, logger=logger)

        mock_presentation = MagicMock()
        mock_presentation.Slides.Count = 1
        mock_slide = MagicMock()
        mock_slide.Export.side_effect = RuntimeError("COM export failed")
        mock_presentation.Slides.return_value = mock_slide

        mock_ppt_app = MagicMock()
        mock_ppt_app.Presentations.Open.return_value = mock_presentation

        fake_comtypes = types.ModuleType("comtypes")
        fake_comtypes.CoInitialize = MagicMock()  # type: ignore[attr-defined]
        fake_comtypes.CoUninitialize = MagicMock()  # type: ignore[attr-defined]
        fake_comtypes_client = types.ModuleType("comtypes.client")
        fake_comtypes_client.CreateObject = MagicMock(return_value=mock_ppt_app)  # type: ignore[attr-defined]
        fake_comtypes.client = fake_comtypes_client  # type: ignore[attr-defined]

        with tempfile.TemporaryDirectory() as tmp:
            work_dir = Path(tmp)
            pptx_path = work_dir / "test.pptx"
            pptx_path.write_bytes(b"PK\x03\x04")
            with patch.dict(
                sys.modules,
                {"comtypes": fake_comtypes, "comtypes.client": fake_comtypes_client},
            ), self.assertRaises(SlideExportError):
                pipeline._export_slides_via_com(pptx_path, work_dir)

        mock_presentation.Close.assert_called_once()
        mock_ppt_app.Quit.assert_called_once()
        fake_comtypes.CoUninitialize.assert_called_once()


class TestWebUIVoicesAndMultipart(unittest.TestCase):
    """Web UI の日本語話者一覧取得および multipart 解析のテスト."""

    def test_fetch_japanese_voices_filters_ja_locales(self) -> None:
        import web_ui

        web_ui._VOICE_CACHE = None
        fake_edge_tts = types.ModuleType("edge_tts")

        async def fake_list_voices() -> list[dict[str, str]]:
            return [
                {"ShortName": "en-US-AriaNeural", "Gender": "Female", "Locale": "en-US"},
                {"ShortName": "en-US-AvaMultilingualNeural", "Gender": "Female", "Locale": "en-US"},
                {"ShortName": "ja-JP-KeitaNeural", "Gender": "Male", "Locale": "ja-JP"},
                {"ShortName": "ja-JP-NanamiNeural", "Gender": "Female", "Locale": "ja-JP"},
            ]

        fake_edge_tts.list_voices = fake_list_voices  # type: ignore[attr-defined]
        with patch.dict(sys.modules, {"edge_tts": fake_edge_tts}):
            data = web_ui.fetch_japanese_voices(force_refresh=True)

        self.assertEqual(data["source"], "edge-tts")
        self.assertEqual(data["count"], 2)
        voices = data["voices"]
        assert isinstance(voices, list)
        ids = [v["id"] for v in voices]
        self.assertEqual(ids, ["ja-JP-NanamiNeural", "ja-JP-KeitaNeural"])
        web_ui._VOICE_CACHE = None

    def test_parse_multipart_form(self) -> None:
        import web_ui

        boundary = "----TestBoundary7MA4YWxkTrZu0gW"
        body = (
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="voice"\r\n\r\n'
            "ja-JP-MayuNeural\r\n"
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="file"; filename="sample.pptx"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n"
            "PK\x03\x04dummy\r\n"
            f"--{boundary}--\r\n"
        ).encode()
        fields, uploaded = web_ui.parse_multipart_form(
            f"multipart/form-data; boundary={boundary}",
            body,
        )
        self.assertEqual(fields["voice"], "ja-JP-MayuNeural")
        self.assertIsNotNone(uploaded)
        assert uploaded is not None
        self.assertEqual(uploaded[0], "sample.pptx")
        self.assertTrue(uploaded[1].startswith(b"PK\x03\x04"))

    def test_http_tts_preview_endpoint(self) -> None:
        import json
        import threading
        import urllib.request
        from http.server import ThreadingHTTPServer

        import web_ui

        with tempfile.TemporaryDirectory() as tmp:
            manager = web_ui.JobManager(Path(tmp), output_dir=Path(tmp) / "out")
            server = ThreadingHTTPServer(("127.0.0.1", 0), web_ui.make_handler(manager))
            port = server.server_address[1]
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                with patch.object(
                    web_ui, "synthesize_preview_audio", return_value=(b"RIFFfakewav", "audio/wav")
                ):
                    req = urllib.request.Request(
                        f"http://127.0.0.1:{port}/api/tts-preview",
                        data=json.dumps({"voice": "winrt:Microsoft Ayumi", "text": "テスト"}).encode("utf-8"),
                        headers={"Content-Type": "application/json"},
                        method="POST",
                    )
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        self.assertEqual(resp.status, 200)
                        self.assertEqual(resp.headers.get("Content-Type"), "audio/wav")
                        self.assertEqual(resp.read(), b"RIFFfakewav")
            finally:
                server.shutdown()
                server.server_close()

    def test_trim_pcm_silence(self) -> None:
        import struct

        from loopback_recorder import trim_pcm_silence

        # 0.5秒の無音 + 0.1秒の音声 + 0.5秒の無音 (1ch, 1000Hz)
        silence = struct.pack("<h", 0) * 500
        signal = struct.pack("<h", 2000) * 100
        trimmed = trim_pcm_silence(silence + signal + silence, channels=1, sample_rate=1000, pad_seconds=0.05)
        # 100 + 50*2 = 200サンプル (400 bytes) 前後になること
        self.assertEqual(len(trimmed), 200 * 2)

    def test_voicevox_config_validation(self) -> None:
        self.assertEqual(PipelineConfig(voice="voicevox:3").resolve_voice_id(), "voicevox:3")
        with self.assertRaises(InputValidationError):
            PipelineConfig(voice="voicevox:abc").resolve_voice_id()
        self.assertEqual(
            PipelineConfig(voicevox_url="http://127.0.0.1:50021/").resolve_voicevox_url(),
            "http://127.0.0.1:50021",
        )
        with self.assertRaises(InputValidationError):
            PipelineConfig(voicevox_url="ftp://127.0.0.1:50021").resolve_voicevox_url()

    def test_synthesize_voicevox_wav(self) -> None:
        import io
        import json

        captured_requests: list[tuple[str, bytes | None]] = []
        fake_wav = b"RIFF" + b"\x00" * 44

        def fake_urlopen(req, timeout=0):
            captured_requests.append((req.full_url, req.data))
            if "/audio_query" in req.full_url:
                payload = json.dumps({"accent_phrases": [], "speedScale": 1.0}).encode("utf-8")
                return io.BytesIO(payload)
            return io.BytesIO(fake_wav)

        config = PipelineConfig(
            voice="voicevox:3",
            voicevox_url="http://127.0.0.1:50021",
            voicevox_speed=1.2,
            voicevox_pitch=0.05,
            voicevox_intonation=1.3,
            voicevox_volume=0.9,
        )
        pipeline = PowerPointVideoPipeline(config=config)
        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "slide_001.wav"
            with patch("urllib.request.urlopen", side_effect=fake_urlopen):
                ok = pipeline._synthesize_voicevox_wav("テスト発話", out_path, speaker_id=3, slide_num=1)
            self.assertTrue(ok)
            self.assertEqual(out_path.read_bytes(), fake_wav)

        self.assertEqual(len(captured_requests), 2)
        self.assertIn("/audio_query?text=", captured_requests[0][0])
        self.assertIn("speaker=3", captured_requests[0][0])
        self.assertIn("/synthesis?speaker=3", captured_requests[1][0])
        synth_body = json.loads((captured_requests[1][1] or b"{}").decode("utf-8"))
        self.assertEqual(synth_body["speedScale"], 1.2)
        self.assertEqual(synth_body["pitchScale"], 0.05)
        self.assertEqual(synth_body["intonationScale"], 1.3)
        self.assertEqual(synth_body["volumeScale"], 0.9)

    def test_fetch_voicevox_speakers(self) -> None:
        import io
        import json

        import web_ui

        fake_speakers = [
            {
                "name": "ずんだもん",
                "speaker_uuid": "uuid-1",
                "styles": [{"id": 3, "name": "ノーマル"}, {"id": 1, "name": "あまあま"}],
            }
        ]
        with patch(
            "urllib.request.urlopen",
            return_value=io.BytesIO(json.dumps(fake_speakers).encode("utf-8")),
        ):
            result = web_ui.fetch_voicevox_speakers("http://127.0.0.1:50021/")
        self.assertEqual(result["url"], "http://127.0.0.1:50021")
        speakers = result["speakers"]
        assert isinstance(speakers, list)
        self.assertEqual(len(speakers), 1)
        self.assertEqual(speakers[0]["name"], "ずんだもん")
        self.assertEqual(speakers[0]["styles"][0]["id"], 3)

    def test_pronunciation_dict_word_boundaries(self) -> None:
        from pronunciation import apply_pronunciation_dict, load_pronunciation_dict

        mapping = load_pronunciation_dict()
        self.assertGreaterEqual(len(mapping), 100)
        text = "VSCodeでpptxを作成し、macOSのmain関数とAIを確認する。"
        converted = apply_pronunciation_dict(text, mapping)
        self.assertIn("ブイエスコード", converted)
        self.assertIn("ピーピーティーエックス", converted)
        self.assertIn("マックオーエス", converted)
        self.assertIn("main関数", converted)
        self.assertIn("エーアイ", converted)

    def test_prepare_tts_text_with_ai_and_fallback(self) -> None:
        import io
        import json

        from pronunciation import prepare_tts_text

        mapping = {"VSCode": "ブイエスコード", "pptx": "ピーピーティーエックス"}
        fake_response = {
            "choices": [{"message": {"content": "ブイエスコードでピーピーティーエックスとクロノスを起動"}}]
        }
        cache: dict[str, str] = {}
        with patch(
            "urllib.request.urlopen",
            return_value=io.BytesIO(json.dumps(fake_response).encode("utf-8")),
        ) as mock_open:
            out = prepare_tts_text(
                "VSCodeでpptxとChronosを起動",
                dict_mapping=mapping,
                ai_enabled=True,
                ai_base_url="http://127.0.0.1:4000/v1",
                ai_model="gpt-4o-mini",
                cache=cache,
            )
            self.assertEqual(out, "ブイエスコードでピーピーティーエックスとクロノスを起動")
            # 2回目はキャッシュヒットして urlopen が増えないこと
            out2 = prepare_tts_text(
                "VSCodeでpptxとChronosを起動",
                dict_mapping=mapping,
                ai_enabled=True,
                ai_base_url="http://127.0.0.1:4000/v1",
                ai_model="gpt-4o-mini",
                cache=cache,
            )
            self.assertEqual(out2, out)
            self.assertEqual(mock_open.call_count, 1)

        # AI 接続エラー時は辞書適用済みテキストに安全にフォールバックすること
        with patch("urllib.request.urlopen", side_effect=OSError("connection refused")):
            fallback = prepare_tts_text(
                "VSCodeでUnknownToolを起動",
                dict_mapping=mapping,
                ai_enabled=True,
                ai_base_url="http://127.0.0.1:4000/v1",
                ai_model="gpt-4o-mini",
            )
            self.assertEqual(fallback, "ブイエスコードでUnknownToolを起動")

    def test_fetch_lmproxy_models(self) -> None:
        import io
        import json

        from pronunciation import fetch_lmproxy_models

        fake_models = {"data": [{"id": "gpt-4o-mini"}, {"id": "claude-3-5-sonnet"}]}
        with patch(
            "urllib.request.urlopen",
            return_value=io.BytesIO(json.dumps(fake_models).encode("utf-8")),
        ):
            res = fetch_lmproxy_models("http://127.0.0.1:4000/v1/")
        self.assertEqual(res["url"], "http://127.0.0.1:4000/v1")
        self.assertEqual(res["models"], ["claude-3-5-sonnet", "gpt-4o-mini"])


if __name__ == "__main__":
    unittest.main()
