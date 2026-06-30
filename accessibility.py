from pptx import Presentation

MIN_FONT_SIZE = 18
GENERIC_LINK_TEXT = ["click here", "here", "link"]

def get_relative_luminance(color_rgb):
    r, g, b = color_rgb[0] / 255.0, color_rgb[1] / 255.0, color_rgb[2] / 255.0
    def adjust(val):
        return val / 12.92 if val <= 0.03928 else ((val + 0.055) / 1.055) ** 2.4
    return 0.2126 * adjust(r) + 0.7152 * adjust(g) + 0.0722 * adjust(b)

def get_contrast_ratio(rgb1, rgb2):
    l1 = get_relative_luminance(rgb1)
    l2 = get_relative_luminance(rgb2)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)

def check_ppt_accessibility(file_path):
    prs = Presentation(file_path)
    issues = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        # Check slide title
        if not slide.shapes.title or not slide.shapes.title.text.strip():
            issues.append({"slide": slide_index, "category": "Missing Slide Title", "severity": "Error", "detail": "Slide is missing a title shape."})

        for shape in slide.shapes:

            # Check empty text boxes
            if hasattr(shape, "text"):
                if not shape.text.strip():
                    issues.append({"slide": slide_index, "category": "Empty Text Box", "severity": "Tip", "detail": "An empty text box was found on the slide."})

            # Check font sizes
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size:
                            font_size = run.font.size.pt
                            if font_size < MIN_FONT_SIZE:
                                issues.append({"slide": slide_index, "category": "Small Font Size", "severity": "Tip", "detail": f"Small font ({font_size}pt) used in text '{run.text}'"})

                        # Check generic hyperlink text
                        if run.hyperlink.address:
                            if run.text.lower() in GENERIC_LINK_TEXT:
                                issues.append({"slide": slide_index, "category": "Generic Hyperlink Text", "severity": "Warning", "detail": f"Generic hyperlink text '{run.text}' used."})

            # Check reading order (title should be read first)
            if slide.shapes.title and len(slide.shapes) > 0:
                if slide.shapes[0] != slide.shapes.title:
                    issues.append({"slide": slide_index, "category": "Reading Order Issue", "severity": "Warning", "detail": "Title is not the first element read by screen readers."})

            # Check images alt text
            if shape.shape_type == 13:  # Picture
                alt_text = shape.name
                if not alt_text or alt_text.startswith("Picture"):
                    issues.append({"slide": slide_index, "category": "Missing Alt Text", "severity": "Error", "detail": "Image is missing meaningful alt text."})

            # Check tables
            if shape.has_table:
                table = shape.table
                first_row_empty = all(
                    cell.text.strip() == "" for cell in table.rows[0].cells
                )
                header_style_disabled = not getattr(table, "first_row", True)
                if first_row_empty:
                    issues.append({"slide": slide_index, "category": "Missing Table Header", "severity": "Error", "detail": "Table header row is missing or empty."})
                elif header_style_disabled:
                    issues.append({"slide": slide_index, "category": "Missing Table Header", "severity": "Warning", "detail": "Table has header text, but header row styling is disabled."})
                    
            # Color Contrast Warning (WCAG 2.1 Contrast Ratio Check)
            if hasattr(shape, "fill") and shape.fill.type == 1: # SOLID
                try:
                    bg_color = shape.fill.fore_color.rgb
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.color and run.font.color.type == 1: # RGB
                                    fg_color = run.font.color.rgb
                                    font_size = run.font.size.pt if run.font.size else 14
                                    is_large_text = font_size >= 18 or (font_size >= 14 and run.font.bold)
                                    ratio = get_contrast_ratio(bg_color, fg_color)
                                    required = 3.0 if is_large_text else 4.5
                                    if ratio < required:
                                        issues.append({
                                            "slide": slide_index,
                                            "category": "Poor Color Contrast",
                                            "severity": "Error",
                                            "detail": f"Contrast ratio is {ratio:.1f}:1, which is below the WCAG AA requirement of {required}:1 for text '{run.text}'."
                                        })
                except Exception:
                    pass

            # Check for audio/video media subtitles
            is_media = (
                (hasattr(shape, "shape_type") and int(shape.shape_type) in [16, 24]) or 
                "video" in shape.name.lower() or 
                "audio" in shape.name.lower() or 
                "media" in shape.name.lower()
            )
            if is_media:
                issues.append({
                    "slide": slide_index,
                    "category": "Missing Media Subtitles",
                    "severity": "Warning",
                    "detail": f"Media element '{shape.name}' was found. Ensure it has subtitles, closed captions, or a text transcript."
                })

    # Check for duplicate slide titles
    title_map = {}
    for slide_index, slide in enumerate(prs.slides, start=1):
        if slide.shapes.title and slide.shapes.title.text.strip():
            title_text = slide.shapes.title.text.strip().lower()
            if title_text not in title_map:
                title_map[title_text] = []
            title_map[title_text].append(slide_index)

    for title_text, slide_indices in title_map.items():
        if len(slide_indices) > 1:
            for idx in slide_indices:
                other_slides = [str(i) for i in slide_indices if i != idx]
                issues.append({
                    "slide": idx,
                    "category": "Duplicate Slide Title",
                    "severity": "Warning",
                    "detail": f"This slide has the same title as slide(s): {', '.join(other_slides)}."
                })

    # Check for duplicate slides (identically matched content)
    slide_contents = []
    for slide_index, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text.strip())
            elif shape.shape_type == 13:  # Picture
                texts.append(f"[Image:{shape.name}]")
        content_sig = "||".join(sorted(texts))
        slide_contents.append((slide_index, content_sig))

    content_map = {}
    for idx, sig in slide_contents:
        if sig:
            if sig not in content_map:
                content_map[sig] = []
            content_map[sig].append(idx)

    for sig, slide_indices in content_map.items():
        if len(slide_indices) > 1:
            for idx in slide_indices:
                other_slides = [str(i) for i in slide_indices if i != idx]
                issues.append({
                    "slide": idx,
                    "category": "Duplicate Slide Content",
                    "severity": "Warning",
                    "detail": f"This slide has identical content to slide(s): {', '.join(other_slides)}."
                })

    return issues

if __name__ == "__main__":
    file_path = "sample.pptx"
    results = check_ppt_accessibility(file_path)

    if results:
        print("Accessibility Issues Found:")
        for issue in results:
            print(f"- Slide {issue['slide']}: {issue['category']} - {issue['detail']}")
    else:
        print("No accessibility issues found.")
