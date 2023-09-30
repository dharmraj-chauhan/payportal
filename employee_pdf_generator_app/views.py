from django.shortcuts import render
from django.http import HttpResponse, JsonResponse, FileResponse
import pandas as pd

from decouple import config

try:
    url = config('NGINX_URL')
except:
    url = config('URL')

filename = r'Datas/Employee_data.xlsx'

def home_page(request):
    context = {'url': url}
    return render(request, 'employee_pdf_generator_app/home_page.html', context)


def get_employee_id(request):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    data = df['E.CODE'].to_json(orient='values')
    employee_data = {'data': data}

    return JsonResponse(employee_data)


def get_emp_data_by_emp_id(request, employee_id):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    filtered_row = df[df['E.CODE'] == str(employee_id)]
    employee_data_as_dict = filtered_row.to_dict(orient='records')[0]

    new_refine_dict = {
        'employee_name' : employee_data_as_dict['Name of Employe'],
        'employee_dob' : employee_data_as_dict['DOB'],
        'employee_doj' : employee_data_as_dict['DOJ'],
        'employee_department' : employee_data_as_dict['DEPARTMENT'],
        'employee_mnumber' : employee_data_as_dict['MOBILE'],
        'employee_address' : str(employee_data_as_dict['PERMENANT ADDRESS']).strip('AT:'),
    }
    
    return JsonResponse(new_refine_dict)

def emp_data_by_emp_id(employee_id):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    filtered_row = df[df['E.CODE'] == str(employee_id)]
    employee_data_as_dict = filtered_row.to_dict(orient='records')[0]

    new_refine_dict = {
        'employee_full_name'        : str(employee_data_as_dict['Name of Employe']),
        'employee_father_name'      : str(employee_data_as_dict['FATHER NAME']),
        'employee_dob'              : str(employee_data_as_dict['DOB']),
        'employee_sex'              : 'MALE',           # const value passing
        'employee_maritual_status'  : 'UNMARRIED',      # const value passing
        'employee_account_number'   : str(employee_data_as_dict['PF NO']),
        'employee_temp_address'     : str(employee_data_as_dict['PRESENT ADDRESS']).strip('AT:'),
        'employee_const_address'    : str(employee_data_as_dict['PERMENANT ADDRESS']).strip('AT:'),
    }
    
    return new_refine_dict


from django.http import HttpResponse
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont("Helvetica-Bold", r"Fonts/Helvetica-Bold-Font.ttf"))

def generate_pdf(request, employee_id):
    page_width, page_height = letter
    buffer = BytesIO()

    template = PdfReader(r"Forms/Form_2_Revised.pdf", decompress=False)
    template_page_count = len(template.pages)
    canvas = Canvas(buffer, pagesize=(page_width, page_height))

    font_size = 8
    font_name = "Helvetica"

    employee_data = emp_data_by_emp_id(employee_id)

    for page_num in range(template_page_count):
        template_obj = pagexobj(template.pages[page_num])
        xobj_name = makerl(canvas, template_obj)
        canvas.doForm(xobj_name)
        canvas.setFont(font_name, font_size)

        if(page_num == 0):
            canvas.drawString(180, 576, employee_data['employee_full_name'])
            canvas.drawString(180, 558, employee_data['employee_father_name'])
            canvas.drawString(180, 540, employee_data['employee_dob'])
            canvas.drawString(180, 521, employee_data['employee_sex'])
            canvas.drawString(180, 502, employee_data['employee_maritual_status'])
            canvas.drawString(180, 484, employee_data['employee_account_number'])

            canvas.drawString(250, 466, employee_data['employee_temp_address'])
            canvas.drawString(250, 448, employee_data['employee_const_address'])

            # canvas.drawString(250, 429, employee_data[''])
            # canvas.drawString(250, 411, employee_data[''])

        if(page_num == 1):
            canvas.drawString(61, 143, "Vadodara")
            employee_full_name = employee_data['employee_full_name'].split(" ")

            canvas.drawString(464, 107, ":")

            canvas.setFont("Helvetica-Bold", font_size)
            canvas.drawString(443, 198, employee_full_name[0] + " " + employee_full_name[1])
            canvas.drawString(32, 180, employee_full_name[2])

            canvas.drawString(415, 154, "PARMAR KAILASHBEN KISHORSINH")
            canvas.drawString(469, 107, "PROPRIETOR")

        canvas.showPage()

    canvas.save()
    buffer.seek(0)

    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="result.pdf"'

    buffer.close()

    return response