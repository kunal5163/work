import os
import json
import fitz  # PyMuPDF

def emu_to_points(emu):
    return emu / 12700.0

def extract_pdf_layout(pdf_path):
    doc = fitz.open(pdf_path)
    layout = []
    for page_num, page in enumerate(doc):
        pdf_width = page.rect.width
        pdf_height = page.rect.height
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                line_text = "".join([span["text"] for span in line["spans"]])
                bbox = line["bbox"]  # (x0, y0, x1, y1)
                layout.append({
                    "slide_number": page_num + 1,
                    "text": line_text.strip(),
                    "bbox": bbox,
                    "pdf_width": pdf_width,
                    "pdf_height": pdf_height
                })
    return layout

def scale_bbox(bbox, pdf_width, pdf_height, pptx_width_pt, pptx_height_pt):
    scale_x = pptx_width_pt / pdf_width
    scale_y = pptx_height_pt / pdf_height
    x0, y0, x1, y1 = bbox
    return [x0 * scale_x, y0 * scale_y, x1 * scale_x, y1 * scale_y]

def attach_rendered_lines(json_path, pdf_path, output_json="output_with_layout.json"):
    with open(json_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)

    pptx_width_pt = emu_to_points(json_data["slide_width_emu"])
    pptx_height_pt = emu_to_points(json_data["slide_height_emu"])
    layout_lines = extract_pdf_layout(pdf_path)

    for slide in json_data["slides"]:
        slide_num = slide["slide_number"]
        for shape in slide["shapes"]:
            if shape["type"] != "text":
                continue
            x0 = shape["position"]["x_pt"]
            y0 = shape["position"]["y_pt"]
            x1 = x0 + shape["size"]["width_pt"]
            y1 = y0 + shape["size"]["height_pt"]

            lines_in_shape = []
            for line in layout_lines:
                if line["slide_number"] != slide_num:
                    continue
                scaled_bbox = scale_bbox(line["bbox"], line["pdf_width"], line["pdf_height"],
                                         pptx_width_pt, pptx_height_pt)
                line_x, line_y = scaled_bbox[0], scaled_bbox[1]
                if x0 <= line_x <= x1 and y0 <= line_y <= y1:
                    lines_in_shape.append(line["text"])

            if lines_in_shape:
                shape["rendered_lines"] = lines_in_shape

    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=4, ensure_ascii=False)

    print(f"\n✅ Enhanced JSON with rendered layout saved to: {output_json}")

# To run:
# attach_rendered_lines("output_data.json", "input.pdf")
# Example usage
if __name__ == "__main__":
    json_file = "output_data.json"   # your JSON file
    pdf_file = "input.pdf"           # your PDF file
    output_json = "output_with_layout.json"  # output file

    attach_rendered_lines(json_file, pdf_file, output_json)      