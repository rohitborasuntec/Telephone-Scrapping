from docxtpl import DocxTemplate

doc = DocxTemplate(r"c:\Users\hp\Downloads\letter one final draft 09.08.2026.docx")
Address = "1 Linden Way, Kingfield, Surrey, Woking, GU22 9BS"
Address1 = Address.replace(", ", ",\n")
doc.render({
    # "Applicant" : "Tina",
    "Reference": "PA/26/00689/S",
    "Address1": Address1,
    "Address": Address,
})
doc.save("test_output_long.docx")
print("Done - check test_output_long.docx")