import os
from datetime import datetime
from io import BytesIO

import qrcode.image.svg
from barcode import Code128
from pdfrw import PdfReader, PdfWriter, IndirectPdfDict
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.graphics import renderPDF
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from svglib.svglib import svg2rlg

outfile = "result"
data = {
    "region": "16",
    "code_oo": "591054",
    "grade": "11",
    "title": "Пробный ЕГЭ в IT-лицее КФУ",
    "subj_code": "01",
    "subj_name": "мат",
    "blank_footer": "При нехватке места используйте",
    "blank_footer_2": "на оборотной стороне листа",
    "blank_register": "Бланк регистрации",
    "blank_answ1": "Бланк ответов №1",
    "blank_answ2": "Бланк ответов №2",
    "blank_answ2_sub1": "Лист 1",
    "blank_answ2_sub2": "Лист 2",
    "blank_answ3": "Дополнительный бланк ответов №2",
    "qrcode": "https://vk.com/iisdf"
}
base_height = 297 * mm
base_char_space = 6.05

image = qrcode.QRCode(box_size=18, border=0)
image.add_data(data["qrcode"])
image.make_image(image_factory=qrcode.image.svg.SvgPathImage).save("qr.svg")

rv = BytesIO()
Code128(str(int(datetime.now().timestamp() * 1000))).save("barcode", {
    "module_width": 0.3,
    "module_height": 10.0,
    "font_size": 5.0,
    "text_distance": 2.0,
    "quiet_zone": 0.0,
    "center_text": True
})

pages = PdfReader("template.pdf", decompress=False).pages
pdfWriter = PdfWriter()

for page_num in range(len(pages)):
    template = pages[page_num]
    template_obj = pagexobj(template)

    canvas = Canvas(outfile + "_" + str(page_num) + ".pdf")

    xobj_name = makerl(canvas, template_obj)
    canvas.doForm(xobj_name)

    pdfmetrics.registerFont(TTFont('GlobalFont', 'ttcommons.ttf'))
    canvas.setFont("GlobalFont", 20)

    # Заголовок
    canvas.drawString(54 * mm + ((173 - 54) * mm - stringWidth(data["title"], "GlobalFont", 20)) / 2.0,
                      base_height - 11 * mm, data["title"])

    barcode = svg2rlg("barcode.svg")
    if barcode.height > 14 * mm or barcode.height < 14 * mm:
        coef = (14 * mm) / barcode.height
        barcode.height = 14 * mm
        barcode.width *= coef
        barcode.scale(coef, coef)
    qrcode = svg2rlg("qr.svg")
    if qrcode.height > 28.5 * mm or barcode.height < 28.5 * mm:
        coef = (28.5 * mm) / qrcode.height
        qrcode.height = 28 * mm
        qrcode.width *= coef
        qrcode.scale(coef, coef)
    renderPDF.draw(barcode, canvas, 20 * mm + ((70.5 - 20) * mm - barcode.width) / 2.0, base_height - 57 * mm)

    if barcode.height > 13 * mm or barcode.height < 13 * mm:
        coef = (13 * mm) / barcode.height
        barcode.height = 13 * mm
        barcode.width *= coef
        barcode.scale(coef, coef)
    barcode.rotate(90)
    renderPDF.draw(barcode, canvas, 203.25 * mm, base_height - 57 * mm + ((57 - 9) * mm - barcode.width) / 2.0)
    renderPDF.draw(qrcode, canvas, 8.5 * mm, base_height - 38 * mm)
    if page_num == 0:
        # Подзаголовок
        canvas.drawString(54 * mm + ((173 - 54) * mm - stringWidth(data["blank_register"], "GlobalFont", 20)) / 2.0,
                          base_height - 19.5 * mm, data["blank_register"])

        # Регион
        canvas.drawString(47.5 * mm, base_height - 35.5 * mm, data["region"], charSpace=base_char_space)
        # Код ОО
        canvas.drawString(65.5 * mm, base_height - 35.5 * mm, data["code_oo"], charSpace=base_char_space)
        # Класс
        canvas.drawString(106.5 * mm, base_height - 35.5 * mm, data["grade"], charSpace=base_char_space)

        # Название предмета
        canvas.drawString(96 * mm, base_height - 54.5 * mm, data["subj_name"], charSpace=base_char_space + 1)
        # Код предмета
        canvas.drawString(80 * mm, base_height - 54.5 * mm, data["subj_code"], charSpace=base_char_space)
    elif page_num == 1:
        # Подзаголовок
        canvas.drawString(52 * mm + ((173 - 54) * mm - stringWidth(data["blank_answ1"], "GlobalFont", 20)) / 2.0,
                          base_height - 18.5 * mm, data["blank_answ1"])

        # Регион
        canvas.drawString(50 * mm, base_height - 31.5 * mm, data["region"], charSpace=base_char_space)
        # Код предмета
        canvas.drawString(66.25 * mm, base_height - 31.5 * mm, data["subj_code"], charSpace=base_char_space)
        # Название предмета
        canvas.drawString(83.5 * mm, base_height - 31.5 * mm, data["subj_name"], charSpace=base_char_space + 1)

    elif page_num == 2:
        # Подзаголовок
        canvas.drawString(53 * mm + ((129 - 53) * mm - stringWidth(data["blank_answ2"], "GlobalFont", 20)) / 2.0,
                          base_height - 18.5 * mm, data["blank_answ2"])
        # Подзаголовок 2
        canvas.drawString(133 * mm, base_height - 18.5 * mm, data["blank_answ2_sub1"])

        # Регион
        canvas.drawString(56.5 * mm, base_height - 30 * mm, data["region"], charSpace=base_char_space)
        # Код предмета
        canvas.drawString(79.5 * mm, base_height - 30 * mm, data["subj_code"], charSpace=base_char_space)
        # Название предмета
        canvas.drawString(103.25 * mm, base_height - 30 * mm, data["subj_name"], charSpace=base_char_space + 1)

        # Предупреждение
        canvas.setFont("GlobalFont", 10)
        footer = data["blank_footer"] + " " + data["blank_answ2"] + " (" + data["blank_answ2_sub2"] + ") " + data[
            "blank_footer_2"]
        canvas.drawString(15 * mm + ((177 - 15) * mm - stringWidth(footer, "GlobalFont", 10)) / 2.0,
                          base_height - 289 * mm, footer)
    elif page_num == 3:
        # Подзаголовок
        canvas.drawString(53 * mm + ((129 - 53) * mm - stringWidth(data["blank_answ2"], "GlobalFont", 20)) / 2.0,
                          base_height - 18.5 * mm, data["blank_answ2"])
        # Подзаголовок 2
        canvas.drawString(133 * mm, base_height - 18.5 * mm, data["blank_answ2_sub2"])

        # Регион
        canvas.drawString(56.5 * mm, base_height - 30 * mm, data["region"], charSpace=base_char_space)
        # Код предмета
        canvas.drawString(79.5 * mm, base_height - 30 * mm, data["subj_code"], charSpace=base_char_space)
        # Название предмета
        canvas.drawString(103.25 * mm, base_height - 30 * mm, data["subj_name"], charSpace=base_char_space + 1)

        # Предупреждение
        canvas.setFont("GlobalFont", 10)
        footer = data["blank_footer"] + " " + data["blank_answ3"]
        canvas.drawString(44 * mm + ((197 - 44) * mm - stringWidth(footer, "GlobalFont", 10)) / 2.0,
                          base_height - 289 * mm, footer)
    elif page_num == 4:
        # Подзаголовок
        canvas.drawString(52 * mm + ((173 - 54) * mm - stringWidth(data["blank_answ3"], "GlobalFont", 20)) / 2.0,
                          base_height - 18 * mm, data["blank_answ3"])

        # Регион
        canvas.drawString(56.5 * mm, base_height - 30 * mm, data["region"], charSpace=base_char_space)
        # Код предмета
        canvas.drawString(79.5 * mm, base_height - 30 * mm, data["subj_code"], charSpace=base_char_space)
        # Название предмета
        canvas.drawString(103.25 * mm, base_height - 30 * mm, data["subj_name"], charSpace=base_char_space + 1)

        # Предупреждение
        canvas.setFont("GlobalFont", 10)
        footer = data["blank_footer"] + " " + data["blank_answ3"] + " " + data["blank_footer_2"]
        canvas.drawString(28 * mm + ((197 - 28) * mm - stringWidth(footer, "GlobalFont", 10)) / 2.0,
                          base_height - 289 * mm, footer)

    if os.path.exists(outfile + "_" + str(page_num) + ".pdf"):
        os.remove(outfile + "_" + str(page_num) + ".pdf")
    canvas.save()

    pdfReader = PdfReader(outfile + "_" + str(page_num) + ".pdf")
    pdfWriter.addpages(pdfReader.pages)

pdfWriter.trailer.Info = IndirectPdfDict(
    Title=data["title"],
    Author="Ivan Petrov",
    Subject=data["title"],
    Creator="Ivan Petrov"
)

if os.path.exists("result.pdf"):
    os.remove("result.pdf")
pdfWriter.write(outfile + ".pdf")

for i in range(len(pages)):
    if os.path.exists(outfile + "_" + str(i) + ".pdf"):
        os.remove(outfile + "_" + str(i) + ".pdf")

if os.path.exists("barcode.svg"):
    os.remove("barcode.svg")
if os.path.exists("qr.svg"):
    os.remove("qr.svg")
