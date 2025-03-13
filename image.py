from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/template-1.docx")
tpl.replace_pic("backgrond_image", "images/template-2.jpg")
tpl.render({})
tpl.save("output.docx")