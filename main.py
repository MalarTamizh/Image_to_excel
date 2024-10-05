import os
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.drawing.image import Image

def extract_images_to_excel(pdf_folder, output_excel):
    workbook = Workbook()
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            images = convert_from_path(pdf_path)
            sheet = workbook.create_sheet(title=filename)
            sheet.sheet_view.showGridLines = False
            for i, image in enumerate(images, start=1):
                img_path = os.path.join(pdf_folder, f"{filename}_image_{i}.png")
                image.save(img_path)
                img = Image(img_path)
                sheet.add_image(img, f"A{i}")

    workbook.remove(workbook.active)
    workbook.save(output_excel)
    print(f"Images extracted and saved to {output_excel} successfully.")
pdf_folder = "C:/Users/tamiz/Downloads/oo"
output_excel = "C:/Users/tamiz/Downloads/output.xlsx"
print(output_excel)
extract_images_to_excel(pdf_folder, output_excel)
