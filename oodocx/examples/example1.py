from oodocx import oodocx

d = oodocx.Docx()
body = d.get_body()
body.append(oodocx.paragraph('Hello World!'))
d.save('Example1.docx')