from django.views.generic import TemplateView
from django.http import HttpResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from itertools import cycle
import math


class Home(TemplateView):
    template_name = 'index.html'


def reportXLS(request):
    wb = load_workbook(filename='templates/20200522_Fichas_(PROPUESTA).xlsx')
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename=Ficha_Campo.xlsx'
    # ws = wb.active
    list_photos = ['media/foto_1.jpg', 'media/foto_2.jpg', 'media/foto_3.jpg', 'media/foto_4.jpg',
                   'media/foto_5.jpg', 'media/foto_6.jpg', 'media/foto_7.jpg', 'media/foto_8.jpg', 'media/foto_9.jpg',
                   'media/foto_10.jpg', 'media/foto_11.jpg', 'media/foto_12.jpg', 'media/foto_13.jpg',
                   'media/foto_14.jpg', 'media/foto_15.jpg', 'media/foto_16.jpg', 'media/foto_17.jpg']
    sheets_positions = ['B7', 'Y7', 'B24', 'Y24']

    list_photos_positions = list(zip(cycle(sheets_positions), list_photos))

    sheets = []

    total_sheets = (len(list_photos) // 4)

    for i in range(int(total_sheets)):
        sheet = wb.get_sheet_by_name('Registro_fotografico')
        sheets.append(wb.copy_worksheet(sheet))

    for index, sheet_position_photo in enumerate(list_photos_positions):
        if index < 4:
            sheet = wb.get_sheet_by_name('Registro_fotografico')
            img = Image(sheet_position_photo[1])
            img.width = 350
            img.height = 300
            sheet.add_image(img, sheet_position_photo[0])
        elif index < 8:
            sheet = wb.get_sheet_by_name('Registro_fotografico Copy')
            img = Image(sheet_position_photo[1])
            img.width = 350
            img.height = 300
            sheet.add_image(img, sheet_position_photo[0])
        elif index >= 8:
            sheet_name = "Registro_fotografico Copy" + str((index // 4)-1)
            sheet = wb.get_sheet_by_name(sheet_name)
            img = Image(sheet_position_photo[1])
            img.width = 350
            img.height = 300
            sheet.add_image(img, sheet_position_photo[0])

    wb.save(response)

    return response

