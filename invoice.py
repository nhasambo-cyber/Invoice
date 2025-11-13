from docxtpl import DocxTemplate

doc = DocxTemplate("INVOICE1264.docx_020919 - Copy.docx")

invoice_list = [["PAINT TRAILABLE CYLINDER", 4, 343.97, 1375.88],
                ["SANDBLAST TRAILABLE CYLINDER", 4, 214.82, 859.28]]

doc.render({"Description":"SANDBLAST TRAILABLE CYLINDER", "invoice_list": invoice_list})
doc.save("new_invoice.docx")
