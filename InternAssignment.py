import tempfile
import segno
import xlwings as xw


book = xw.Book('data.xlsx')
sheet = xw.sheets[1]

start_cell = sheet['A1']
urls = start_cell.options(expand='down', ndim=1).value

for ix, url in enumerate(urls):
    # Generate the QR code
    qr = segno.make(url)
    with tempfile.TemporaryDirectory() as td:
        
        filepath = f'{td}/qr.png'
        qr.save(filepath, scale=5, border=0, finder_dark='#15a43a')
        # Insert the QR code to the right of the URL
        destination_cell = start_cell.offset(row_offset=ix, column_offset=1)
        sheet.pictures.add(filepath,
                           left=destination_cell.left,
                           top=destination_cell.top)

