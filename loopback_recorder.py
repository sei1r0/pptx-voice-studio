"""Windows WASAPI ループバック録音モジュール (PyAudioWPatch 使用)."""

from __future__ import annotations

import threading
import wave
from pathlib import Path

import numpy as np


def trim_pcm_silence(
    pcm_bytes: bytes,
    channels: int,
    sample_rate: int,
    threshold: int = 120,
    pad_seconds: float = 0.08,
) -> bytes:
    """16-bit PCM バイト列の前後の無音区間をトリミングする."""
    if not pcm_bytes or channels < 1:
        return pcm_bytes

    samples = np.frombuffer(pcm_bytes, dtype=np.int16)
    usable = (len(samples) // channels) * channels
    if usable == 0:
        return pcm_bytes

    frames = samples[:usable].reshape(-1, channels)
    peak_per_frame = np.max(np.abs(frames), axis=1)
    active = np.where(peak_per_frame > threshold)[0]
    if len(active) == 0:
        return pcm_bytes

    pad_frames = int(max(0.0, pad_seconds) * sample_rate)
    start = max(0, int(active[0]) - pad_frames)
    end = min(len(frames), int(active[-1]) + 1 + pad_frames)
    return frames[start:end].tobytes()


class WasapiLoopbackRecorder:
    """デフォルト再生デバイスの出力を WASAPI ループバックで WAV 録音する."""

    def __init__(self) -> None:
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._chunks: list[bytes] = []
        self._channels = 2
        self._sample_rate = 48000
        self._error: Exception | None = None

    def start(self) -> None:
        """バックグラウンドスレッドで録音を開始する."""
        import pyaudiowpatch as pyaudio  # type: ignore[import-not-found]

        pa = pyaudio.PyAudio()
        try:
            wasapi = pa.get_host_api_info_by_type(pyaudio.paWASAPI)
            default_out = pa.get_device_info_by_index(wasapi["defaultOutputDevice"])
            target_device = default_out
            if not default_out.get("isLoopbackDevice", False):
                matched = None
                for lb in pa.get_loopback_device_info_generator():
                    if default_out["name"] in lb["name"]:
                        matched = lb
                        break
                if matched is None:
                    pa.terminate()
                    raise RuntimeError("WASAPI ループバックデバイスが見つかりません。")
                target_device = matched

            self._channels = max(1, int(target_device["maxInputChannels"]))
            self._sample_rate = int(target_device["defaultSampleRate"])
            stream = pa.open(
                format=pyaudio.paInt16,
                channels=self._channels,
                rate=self._sample_rate,
                input=True,
                input_device_index=int(target_device["index"]),
                frames_per_buffer=1024,
            )
        except Exception:
            pa.terminate()
            raise

        self._stop_event.clear()
        self._chunks = []
        self._error = None

        def _record_loop() -> None:
            try:
                while not self._stop_event.is_set():
                    data = stream.read(1024, exception_on_overflow=False)
                    self._chunks.append(data)
            except Exception as exc:  # noqa: BLE001
                self._error = exc
            finally:
                try:
                    stream.stop_stream()
                    stream.close()
                finally:
                    pa.terminate()

        self._thread = threading.Thread(target=_record_loop, daemon=True)
        self._thread.start()

    def stop_and_save(self, output_wav: Path) -> Path:
        """録音を停止し、前後無音をトリミングして WAV ファイルに保存する."""
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=5.0)
            self._thread = None

        if self._error is not None:
            raise RuntimeError(f"WASAPI 録音中にエラーが発生しました: {self._error}") from self._error

        raw_pcm = b"".join(self._chunks)
        trimmed_pcm = trim_pcm_silence(raw_pcm, self._channels, self._sample_rate)

        output_wav.parent.mkdir(parents=True, exist_ok=True)
        with wave.open(str(output_wav), "wb") as wf:
            wf.setnchannels(self._channels)
            wf.setsampwidth(2)
            wf.setframerate(self._sample_rate)
            wf.writeframes(trimmed_pcm)

        return output_wav
