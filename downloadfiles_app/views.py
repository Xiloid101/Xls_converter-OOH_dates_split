from django.shortcuts import render
from django.http import HttpResponse
import openpyxl
import mimetypes


def index(request):
    if request.method == 'GET':
        return render(request, 'downloadfiles_app/index.html', {})
    else:
        excel_file = request.FILES["excel_file"]

        # you may put validations here to check extension or file size
        wb = openpyxl.load_workbook(excel_file)

        # getting a particular sheet by name out of many sheets
        worksheet = wb["Sheet1"]
        print(worksheet)

        excel_data = list()
        # iterating over the rows and
        # getting value from each cell in row
        for row in worksheet.iter_rows():
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
            excel_data.append(row_data)

        with open('media/media/test.txt', 'w', encoding='utf-8') as file:
            file.write('Python and beegeek forever')

        return render(request, 'downloadfiles_app/result.html')


def result(request):
    return render(request, 'downloadfiles_app/result.html')


def download_file(request):
    fl_path = 'media/media/test.txt'
    filename = 'test.txt'
    fl = open(fl_path, 'r')
    mime_type, _ = mimetypes.guess_type(fl_path)
    response = HttpResponse(fl, content_type=mime_type)
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response
