# PowerPoint 発表者ノート自動動画化パイプライン

Windows 11（Microsoft Office インストール済み環境）において、PowerPoint (`.pptx`) ファイルの各スライドの「発表者ノート」を音声化し、高解像度スライド画像と同期させて MP4 動画 (`.mp4`) を自動生成する Python パイプラインです。

音声生成方式として **標準モード (`edge-tts`)**、**VOICEVOX Engine モード（CPU モード / ローカル API・話者/スタイル選択）**、**Edge 再生＆録音モード（ブラウザ音声＋WASAPI 録音）** に対応しています。

---

## 環境構築手順 (`uv` 仮想環境)

高速なパッケージマネージャ **`uv`** を使用して仮想環境を構築します。

```powershell
# 1. 仮想環境の作成
uv venv .venv

# 2. 仮想環境の有効化
.venv\Scripts\activate

# 3. 依存パッケージのインストール
uv pip install python-pptx edge-tts moviepy comtypes PyAudioWPatch
```

> ※ `PyAudioWPatch` は「Edge 再生＆録音モード」を使用する場合に必要です。

---

## VOICEVOX Engine の導入手順（CPU モード）

ローカルで VOICEVOX 音声（ずんだもん・四国めたん等）を使いたい場合は、VOICEVOX Engine を別途起動します。

1. **ダウンロード**
   - [VOICEVOX/voicevox_engine Releases](https://github.com/VOICEVOX/voicevox_engine/releases) から最新バージョンの **`voicevox_engine-windows-cpu-...`** をダウンロードします（`.vvpp` ではなく `windows-cpu` を選択してください）。
   - `.7z.001` / `.7z.002` のような分割アーカイブになっている場合は、すべての分割ファイルを同じフォルダに置き、**7-Zip** などで `.001` ファイルを展開します。
2. **Engine の起動**
   - 展開したフォルダ内の **`run.exe`** を実行します。
   - コンソールに起動ログが表示され、`http://127.0.0.1:50021` で待ち受け状態になれば準備完了です。
   - ※ すでに **通常の VOICEVOX デスクトップアプリ** をインストール済みの場合は、アプリを起動しておくだけでも同様に `http://127.0.0.1:50021` で利用できます（AivisSpeech 等、互換 API を提供するローカルエンジンでも Engine URL を変更して接続可能です）。
3. **Web UI での設定**
   - 音声生成方式で **「VOICEVOX Engine モード」** を選択し、**「話者・スタイルを取得」** をクリックします。
   - キャラクター（話者）・スタイル（ノーマル／あまあま等）と、**話速 (`speedScale`)・音高 (`pitchScale`)・抑揚 (`intonationScale`)・音量 (`volumeScale`)** を調整し、サンプル音声で試聴したうえで動画を生成できます。

---

## 使い方

### Web UI 実行

追加パッケージ不要（Python 標準ライブラリのみで動作）でローカル Web UI を起動できます。

```powershell
python web_ui.py
```

起動すると自動でブラウザが開き（`http://127.0.0.1:8000`）、`.pptx` のドラッグ＆ドロップ、音声生成方式の切り替え（`edge-tts` / `VOICEVOX Engine` / `Edge 再生＆録音`）、サンプルテキスト試聴、進捗ログ表示、生成 MP4 のプレビューおよびダウンロードが可能です。

### CLI 実行

```powershell
python pptx_to_video.py input.pptx output.mp4 --voice nanami

# VOICEVOX Engine のスタイル ID を直接指定する場合（例: ずんだもん ノーマル = 3）
python pptx_to_video.py input.pptx output.mp4 --voice voicevox:3
```

主なオプション:

| オプション | 既定値 | 説明 |
| :--- | :--- | :--- |
| `--voice` | `nanami` | 話者プリセット (`nanami`, `keita`, `aoi`, `daichi`)、Voice ID、または `voicevox:<style_id>` |
| `--width` | `1920` | スライド出力画像の幅 (px) |
| `--height` | `1080` | スライド出力画像の高さ (px) |
| `--fps` | `30` | 出力動画のフレームレート |
| `--pause` | `0.8` | 各スライド末尾の無音ポーズ秒数 |

### 話者プリセット一覧 (`VOICE_PRESETS`)

`pptx_to_video.py` 冒頭の `VOICE_PRESETS` で設定・拡張できます。

- `nanami`: `ja-JP-NanamiNeural`（女性・落ち着いた標準的な声）
- `keita`: `ja-JP-KeitaNeural`（男性・明瞭でビジネス向け）
- `aoi`: `ja-JP-AoiNeural`（女性・明るく柔らかい声）
- `daichi`: `ja-JP-DaichiNeural`（男性・深みのあるナレーション向け）

### Python コードからの利用

```python
from pptx_to_video import PipelineConfig, PowerPointVideoPipeline

# edge-tts を使用する場合
config = PipelineConfig(voice="keita", pause_seconds=0.8)
pipeline = PowerPointVideoPipeline(config=config)
pipeline.run("presentation.pptx", "output.mp4")

# VOICEVOX Engine を使用する場合
vv_config = PipelineConfig(
    voice="voicevox:3",
    voicevox_url="http://127.0.0.1:50021",
    voicevox_speed=1.05,
    voicevox_pitch=0.0,
    voicevox_intonation=1.0,
    voicevox_volume=1.0,
)
PowerPointVideoPipeline(config=vv_config).run("presentation.pptx", "output_vv.mp4")
```

---

## アーキテクチャ設計のポイント

1. **COM (STA) と `asyncio` の完全分離**
   - **フェーズ1 (同期)**: `comtypes` で PowerPoint をバックグラウンド起動し、FHD (1920x1080) PNG をエクスポート。`try ... finally` で `presentation.Close()` / `ppt_app.Quit()` / `CoUninitialize()` を必ず実行し、`POWERPNT.EXE` のゾンビプロセス化を防止します。
   - **フェーズ2 (非同期)**: COM 完全解放後に `asyncio.run()` で `edge-tts` を呼び出し、エクスポネンシャル・バックオフ付きリトライおよびサーキットブレイカー（連続失敗時は無音 WAV フォールバック）により安全に音声を生成します。
   - **フェーズ3 (動画合成)**: MoviePy を用いて音声の実尺＋無音パディングに画像クリップの尺を一致させ、`libx264` / `aac` で MP4 を書き出します。
2. **セキュリティと一時ファイル管理**
   - 入力 `.pptx` の存在確認・拡張子チェック・`..` を含むパストラバーサル拒否を実施します。
   - `tempfile.TemporaryDirectory()` を使用し、中間ファイルは処理完了後または例外発生時に自動削除されます。
