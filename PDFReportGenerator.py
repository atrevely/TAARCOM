from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


canvas = canvas.Canvas("testing.pdf", pagesize=letter)
canvas.setFont('Helvetica', 12)

canvas.drawString(20, 750, 'QUARTERLY SALES REPORT')
canvas.drawString(20, 735, 'TAARCOM, INC.')
canvas.drawString(535, 750, "12/12/2010")

canvas.drawString(20, 705, 'SALESPERSON: J. Wickware')
canvas.setLineWidth(2)
canvas.line(20, 700, 600, 700)

canvas.setFont('Helvetica-Bold', 10)
canvas.drawString(20, 680, 'Balance due prior month:')
canvas.drawString(400, 680, 'Commissions paid last cycle:')
canvas.drawString(20, 645, 'Manufacturer')
canvas.drawString(200, 645, 'Months')
canvas.drawString(540, 655, 'Salesperson')
canvas.drawString(540, 645, 'Commission')
canvas.setFont('Helvetica', 10)
canvas.drawString(145, 680, '$1,000,000')
canvas.drawString(545, 680, '$2,500,000')


canvas.setLineWidth(0.5)
canvas.line(20, 640, 600, 640)
canvas.line(20, 639, 600, 639)

canvas.save()
