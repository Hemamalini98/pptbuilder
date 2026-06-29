from pptx import Presentation

MIN_FONT_SIZE = 18
GENERIC_LINK_TEXT = ["click here", "here", "link"]

def check_ppt_accessibility(file_path):
    prs = Presentation(file_path)
    issues = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        # Check slide title
        if not slide.shapes.title or not slide.shapes.title.text.strip():
            issues.append({"slide": slide_index, "category": "Missing Slide Title", "detail": "Slide is missing a title shape."})

        for shape in slide.shapes:

            # Check empty text boxes
            if hasattr(shape, "text"):
                if not shape.text.strip():
                    issues.append({"slide": slide_index, "category": "Empty Text Box", "detail": "An empty text box was found on the slide."})

            # Check font sizes
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size:
                            font_size = run.font.size.pt
                            if font_size < MIN_FONT_SIZE:
                                issues.append({"slide": slide_index, "category": "Small Font Size", "detail": f"Small font ({font_size}pt) used in text '{run.text}'"})

                        # Check generic hyperlink text
                        if run.hyperlink.address:
                            if run.text.lower() in GENERIC_LINK_TEXT:
                                issues.append({"slide": slide_index, "category": "Generic Hyperlink Text", "detail": f"Generic hyperlink text '{run.text}' used."})

            # Check reading order (title should be read first)
            if slide.shapes.title and len(slide.shapes) > 0:
                if slide.shapes[0] != slide.shapes.title:
                    issues.append({"slide": slide_index, "category": "Reading Order Issue", "detail": "Title is not the first element read by screen readers."})

            # Check images alt text
            if shape.shape_type == 13:  # Picture
                alt_text = shape.name
                if not alt_text or alt_text.startswith("Picture"):
                    issues.append({"slide": slide_index, "category": "Missing Alt Text", "detail": "Image is missing meaningful alt text."})

            # Check tables
            if shape.has_table:
                table = shape.table
                first_row_empty = all(
                    cell.text.strip() == "" for cell in table.rows[0].cells
                )
                if first_row_empty:
                    issues.append({"slide": slide_index, "category": "Missing Table Header", "detail": "Table header row is missing or empty."})
                    
            # Basic Color Contrast Warning
            if hasattr(shape, "fill") and shape.fill.type == 1: # SOLID
                try:
                    bg_color = shape.fill.fore_color.rgb
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.color and run.font.color.type == 1: # RGB
                                    fg_color = run.font.color.rgb
                                    if bg_color == fg_color:
                                        issues.append({"slide": slide_index, "category": "Poor Color Contrast", "detail": "Text color matches the background color exactly."})
                except Exception:
                    pass

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
