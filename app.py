from flask import Flask, render_template, request, send_file
import pandas as pd
import fitz  # PyMuPDF
import os

app = Flask(__name__)

# Create an uploads directory if it doesn't exist
if not os.path.exists('uploads'):
    os.makedirs('uploads')

def add_links_to_pdf(pdf_path, excel_path, output_pdf_path):
    try:
        excel_data = pd.read_excel(excel_path)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    excel_data.columns = excel_data.columns.str.strip()
    required_columns = {'URL', 'Discription'}

    if not required_columns.issubset(excel_data.columns):
        return f"Excel file must contain 'URL' and 'Discription' columns."

    try:
        pdf_document = fitz.open(pdf_path)
    except Exception as e:
        return f"Error opening PDF file: {e}"
    
    for index, row in excel_data.iterrows():
        url = row['URL']
        description_text = str(row['Discription'])

        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            description_areas = page.search_for(description_text)
            if description_areas:
                combined_area = fitz.Rect()
                for area in description_areas:
                    combined_area |= area

                nearby_area = fitz.Rect(combined_area.x0 - 100, combined_area.y0 - 100,
                                         combined_area.x1 + 100, combined_area.y1 + 100)
                images = page.get_images(full=True)

                image_bboxes = []
                for img_index, img in enumerate(images):
                    xref = img[0]
                    try:
                        bbox = page.get_image_bbox(xref)
                    except ValueError:
                        continue

                    if nearby_area.intersects(bbox):
                        image_bboxes.append(bbox)

                for img_bbox in image_bboxes:
                    combined_area |= img_bbox

                if combined_area:
                    page.insert_link({
                        "kind": fitz.LINK_URI,
                        "from": combined_area,
                        "uri": url
                    })

    try:
        pdf_document.save(output_pdf_path)
        pdf_document.close()
        return output_pdf_path
    except Exception as e:
        return f"Error saving the modified PDF: {e}"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        pdf_file = request.files['pdf_file']
        excel_file = request.files['excel_file']
        
        if pdf_file and excel_file:
            pdf_path = os.path.join('uploads', pdf_file.filename)
            excel_path = os.path.join('uploads', excel_file.filename)
            output_pdf_path = os.path.join('uploads', 'output.pdf')

            pdf_file.save(pdf_path)
            excel_file.save(excel_path)

            result = add_links_to_pdf(pdf_path, excel_path, output_pdf_path)

            if os.path.exists(result):
                return send_file(result, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
