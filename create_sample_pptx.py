"""テスト用の音声合成ノート付き PowerPoint (.pptx) を生成するスクリプト."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def create_sample_pptx(output_path: Path) -> Path:
    """16:9 のサンプル PowerPoint ファイルを作成する."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    slides_data: list[dict[str, Any]] = [
        {
            "title": "PPTX Voice Studio へようこそ",
            "body": "PowerPoint の発表者ノートを\n自動で高品質な音声付き動画に変換します。",
            "bg": RGBColor(19, 26, 46),
            "accent": RGBColor(99, 102, 241),
            "note": (
                "こんにちは。PPTX Voice Studio のテストプレゼンテーションへようこそ。\n"
                "・このスライドは音声合成パイプラインの動作確認用です。\n"
                "・発表者ノートの内容が自動で読み上げられます。"
            ),
        },
        {
            "title": "主な機能とアーキテクチャ",
            "body": "• PowerPoint COM による FHD 高解像度エクスポート\n• edge-tts による自然な日本語 Neural 音声合成\n• MoviePy による動的尺合わせと MP4 エンコード",
            "bg": RGBColor(15, 23, 42),
            "accent": RGBColor(34, 197, 94),
            "note": (
                "続いて、主な機能をご紹介します。\n"
                "- PowerPoint COM を使い、フルHD画質でスライドを画像化します。\n"
                "- 次に edge-tts で非同期に音声を生成し、最後に動画を合成します。"
            ),
        },
        {
            "title": "ご清聴ありがとうございました",
            "body": "Web UI または CLI からいつでも動画を生成できます。",
            "bg": RGBColor(30, 27, 75),
            "accent": RGBColor(245, 158, 11),
            "note": "以上でサンプルプレゼンテーションを終わります。ご清聴ありがとうございました。",
        },
    ]

    for idx, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank_layout)

        # 背景色を設定
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = item["bg"]

        # 上部アクセントバー
        bar = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            Inches(0.8),
            Inches(0.9),
            Inches(1.6),
            Inches(0.12),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = item["accent"]
        bar.line.fill.background()

        # タイトル
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.7), Inches(1.2))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = item["title"]
        p_title.font.size = Pt(40)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(248, 250, 252)

        # 本文
        body_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.7), Inches(3.8))
        tf_body = body_box.text_frame
        tf_body.word_wrap = True
        for line_i, line in enumerate(item["body"].split("\n")):
            p = tf_body.paragraphs[0] if line_i == 0 else tf_body.add_paragraph()
            p.text = line
            p.font.size = Pt(24)
            p.font.color.rgb = RGBColor(203, 213, 225)
            p.space_after = Pt(14)

        # ページ番号
        footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.7), Inches(1.2), Inches(0.5))
        p_footer = footer_box.text_frame.paragraphs[0]
        p_footer.text = f"{idx} / {len(slides_data)}"
        p_footer.font.size = Pt(16)
        p_footer.font.color.rgb = RGBColor(148, 163, 184)

        # 発表者ノート
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = item["note"]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


if __name__ == "__main__":
    out = create_sample_pptx(Path("samples/sample_presentation.pptx"))
    print(f"Generated: {out.resolve()}")
