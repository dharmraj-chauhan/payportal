from django.shortcuts import render
from django.http import HttpResponse, JsonResponse, FileResponse
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import tempfile

from io import BytesIO
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl

filename = r'Datas/Employee_data.xlsx'

def home_page(request):
    return render(request, 'employee_pdf_generator_app/home_page.html')


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
        'employee_name' : str(employee_data_as_dict['Name of Employe']),
        'employee_dob' : str(employee_data_as_dict['DOB']),
        'employee_doj' : str(employee_data_as_dict['DOJ']),        
        'employee_adhar_no' : str(employee_data_as_dict['AADHARCARD']),
        'employee_id' : str(employee_data_as_dict['E.CODE']),
        'employee_department' : str(employee_data_as_dict['DEPARTMENT']),
        'employee_mnumber' : str(employee_data_as_dict['MOBILE']),
        'employee_address' : str(employee_data_as_dict['PERMENANT ADDRESS']).strip('AT:'),
    }
    print(new_refine_dict)
    return JsonResponse(new_refine_dict)

def emp_data_by_emp_id(employee_id):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    filtered_row = df[df['E.CODE'] == str(employee_id)]
    employee_data_as_dict = filtered_row.to_dict(orient='records')[0]

    new_refine_dict = {
        'employee_name' : str(employee_data_as_dict['Name of Employe']),
        'employee_father_name' : str(employee_data_as_dict['FATHER NAME']),
        'employee_dob' : str(employee_data_as_dict['DOB']),
        'employee_doj' : str(employee_data_as_dict['DOJ']),
        'employee_pf_no' : str(employee_data_as_dict['PF NO']),
        'employee_uan_no' : str(employee_data_as_dict['UAN NO']),
        'employee_maritual_status' : str('MARRIED'),
        'employee_sex' : str('MALE'),
        'employee_esic_no' : str(employee_data_as_dict['ESIC NO']),
        'employee_adhar_no' : str(employee_data_as_dict['AADHARCARD']),
        'employee_id' : str(employee_data_as_dict['E.CODE']),
        'employee_BD' : str(employee_data_as_dict['Basic+DA']),
        'employee_HRA' : str(employee_data_as_dict['HRA']),
        'employee_bank_ac_no' : str(employee_data_as_dict['A/C NUMBER']),
        'employee_ifsc_no' : str(employee_data_as_dict['IFSC CODE']),
        'employee_department' : str(employee_data_as_dict['DEPARTMENT']),
        'employee_mnumber' : str(employee_data_as_dict['MOBILE']),
        'employee_temp_address' : str(employee_data_as_dict['PRESENT ADDRESS']).strip('AT:'),
        'employee_const_address' : str(employee_data_as_dict['PERMENANT ADDRESS']).strip('AT:'),
    }
    print(new_refine_dict)
    
    return new_refine_dict

pdfmetrics.registerFont(TTFont("Helvetica-Bold", r"Fonts/Helvetica-Bold-Font.ttf"))

def form_2_generate_pdf(request, employee_id):
    buffer = BytesIO()

    template = PdfReader(r"Forms/Form_2_Revised.pdf", decompress=False)
    template_page_count = len(template.pages)
    page_width, page_height = A4
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
            canvas.drawString(157, 596, employee_data['employee_name'])
            canvas.drawString(157, 578, employee_data['employee_father_name'])
            canvas.drawString(157, 560, employee_data['employee_dob'])
            canvas.drawString(157, 542, employee_data['employee_sex'])
            canvas.drawString(157, 524, employee_data['employee_maritual_status'])
            canvas.drawString(157, 506, employee_data['employee_esic_no'])

            canvas.drawString(225, 489, employee_data['employee_temp_address'])
            canvas.drawString(225, 471, employee_data['employee_const_address'])

        if(page_num == 1):
            canvas.drawString(62, 176, "HALOL")
            employee_full_name = employee_data['employee_name'].split(" ")

            canvas.drawString(452, 140, ":")

            canvas.setFont("Helvetica-Bold", font_size)
            canvas.drawString(431, 229.5, employee_full_name[0] + " " + employee_full_name[1])
            canvas.drawString(34, 211.5, employee_full_name[2])

            canvas.drawString(407, 175, "PARMAR KAILASHBEN KISHORSINH")
            canvas.drawString(456, 140, "PROPRIETOR")

        canvas.showPage()

    canvas.save()
    buffer.seek(0)

    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="result.pdf"'

    buffer.close()

    return response

def form_1_generate_pdf(request, employee_id):
    buffer = BytesIO()

    template = PdfReader(r"Forms/None_Form1_Declaration_ESI.pdf", decompress=False)
    template_page_count = len(template.pages)
    page_width, page_height = A4
    canvas = Canvas(buffer, pagesize=(page_width, page_height))

    font_size = 8
    font_name = "Helvetica"
    img_flage = False
    employee_data = emp_data_by_emp_id(employee_id)

    try:
        pxl_doc = openpyxl.load_workbook(r'Datas/Speed_Ind_Employee_data.xlsx')
        worksheet = pxl_doc['Speed_Industrial Service']

        #calling the image_loader
        image_loader = SheetImageLoader(worksheet)
        row_number = None
        target_employee_code = employee_data['employee_id']
        for row_index, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            employee_code = row[0]
            if employee_code == target_employee_code:
                row_number = row_index
                break

        if row_number is not None:
            img_flage = True
            print(f"Employee with code '{target_employee_code}' found in row {row_number}.")
        else:
            print(f"Employee with code '{target_employee_code}' not found in the Excel file.")

        if img_flage:
            image = image_loader.get(f'B{row_number}')
            temp_image_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
            image.save(temp_image_path)
        else:
            image = None

    except Exception as E:
        print(E)

    for page_num in range(template_page_count):
        template_obj = pagexobj(template.pages[page_num])
        xobj_name = makerl(canvas, template_obj)
        canvas.doForm(xobj_name)
        canvas.setFont(font_name, font_size)

        if(page_num == 0):
            font_size = 8
            font_name = "Helvetica"
            canvas.setFont("Helvetica-Bold", font_size)
            
            employee_full_name = employee_data['employee_name'].split(" ")
            employee_dob = (employee_data['employee_dob'].split(" "))[0].split("-")
            employee_doj = (employee_data['employee_doj'].split(" "))[0].split("-")
            employee_temp_address = employee_data['employee_temp_address']
            employee_const_address = employee_data['employee_const_address']

            canvas.drawString(165, 712, employee_data['employee_esic_no']) # insurance no
            canvas.drawString(165, 693, employee_full_name[0] + " " + employee_full_name[1]) # full name
            canvas.drawString(165, 685, employee_full_name[2])
            canvas.drawString(165, 668, employee_data['employee_father_name']) # father name

            canvas.drawString(163, 621, employee_dob[2]) # date
            canvas.drawString(185, 621, employee_dob[1]) # month
            canvas.drawString(201.5, 621, employee_dob[0]) # month

            try:
                canvas.drawString(75, 570, employee_temp_address[0:20]) # w
                canvas.drawString(75, 561, employee_temp_address[20:40]) # w
                canvas.drawString(75, 551, employee_temp_address[40:]) # w
            except Exception as E:
                print(f"Got the temp address error: {E}")
                pass
            
            try:
                canvas.drawString(195, 570, employee_const_address[0:20]) # w
                canvas.drawString(195, 561, employee_const_address[20:40]) # w
                canvas.drawString(195, 551, employee_const_address[40:]) # w
            except Exception as E:
                print(f"Got the temp address error: {E}")
                pass

            canvas.drawString(394, 712, "38000412640001001") # employee-code
            canvas.drawString(326, 647, "Speed Industrial Service") # employee-code
            canvas.drawString(326, 636, "31/1 Ranchod Nagar Godhra Road Halol") # employee-code
            

            canvas.drawString(434, 668, employee_doj[2]) # date
            canvas.drawString(472, 668, employee_doj[1]) # month
            canvas.drawString(506, 668, employee_doj[0]) # year

            font_size = 20
            font_name = "Helvetica"
            canvas.setFont(font_name, font_size)

            canvas.drawString(270, 610, "") # m
            canvas.drawString(275, 610, "-") # u
            canvas.drawString(283.5, 610, "-") # w

            canvas.drawString(283, 596, "") # m
            canvas.drawString(291, 596, "-") # u

            if (img_flage == True):
                canvas.drawImage(temp_image_path, 402, 59, width=100, height=80)

        canvas.showPage()

    canvas.save()
    buffer.seek(0)

    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="result.pdf"'

    buffer.close()

    return response