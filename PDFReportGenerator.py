from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import locale


def pdfReport(salesperson, data):
    locale.setlocale(locale.LC_ALL, '')

    # Initialize report pdf.
    report = canvas.Canvas("Test Report.pdf", pagesize=letter)
    report.setFont('Helvetica', 12)

    # Header.
    report.drawString(20, 750, 'QUARTERLY SALES REPORT')
    report.drawString(20, 735, 'TAARCOM, INC.')
    report.drawString(535, 750, "8/31/2019")
    report.drawString(20, 705, 'SALESPERSON: ' + salesperson)
    report.setLineWidth(2)
    report.line(20, 700, 600, 700)
    report.setFont('Helvetica-Bold', 10)
    report.drawString(20, 680, 'Balance due prior month:')
    report.drawString(400, 680, 'Commissions paid last cycle:')
    report.drawString(20, 645, 'Manufacturer')
    report.drawString(250, 645, 'Months')
    report.drawString(540, 655, 'Salesperson')
    report.drawString(540, 645, 'Commission')
    report.setFont('Helvetica', 10)
    report.drawString(145, 680, '$1,000,000')
    report.drawString(545, 680, '$2,500,000')

    # Keep track of the full principal names.
    princDict = {'ABR': 'Abracon', 'ATS': 'Advanced Thermal Solutions',
                 'COS': 'Cosel', 'GLO': 'Globtek', 'INF': 'Infineon',
                 'ISS': 'ISSI', 'LAT': 'Lattice Semiconductor', 'OSR': 'Osram',
                 'QRF': 'RF360 Qualcomm', 'TRU': 'Truly',
                 'TRI': 'Triad Semiconductor', 'XMO': 'XMOS'}

    # Iterate over each principal and fill in data.
    report.setLineWidth(0.5)
    shift = 0
    for princ in princDict.keys():
        princData = data[data['Principal'] == princ]
        princComm = sum(princData['Sales Commission'])
        report.line(20, 600-shift, 600, 600-shift)
        commStr = locale.currency(princComm, grouping=True)
        commWidth = stringWidth(commStr, 'Helvetica', 10)
        report.drawString(600-commWidth, 605-shift, commStr)
        report.drawString(20, 605-shift, princDict[princ])

    # Commission total.
    report.line(20, 640, 600, 640)
    report.line(20, 639, 600, 639)
    report.line(20, 600-shift+39, 600, 600-shift+39)
    report.line(20, 600-shift-1, 600, 600-shift-1)
    report.line(20, 600-shift, 600, 600-shift)
    report.setFont('Helvetica-Bold', 10)
    report.drawString(20, 605 - shift, 'TOTAL COMMISSIONS DUE:')

    # Footer.
    report.setFont('Helvetica', 10)
    report.drawString(20, 605-shift-40, 'Sales draw:')
    report.setFillColorRGB(0.9, 0.9, 0.9)
    report.rect(20, 605-shift-80, 580, 20, fill=1)
    report.setFont('Helvetica-Bold', 10)
    report.setFillColorRGB(0, 0, 0)
    report.drawString(25, 605-shift-73, 'BALANCE DUE:')

    report.save()
