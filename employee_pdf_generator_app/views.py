from django.shortcuts import render
from django.http import HttpResponse, JsonResponse, FileResponse
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import tempfile
import textwrap
import zipfile

from io import BytesIO
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageTemplate, Frame, BaseDocTemplate, PageBreak
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet, TA_CENTER
from reportlab.lib import colors

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl

filename = r'Datas/Employee_data.xlsx'

def home_page(request):
    current_url = request.build_absolute_uri()
    context = {
        'current_url': current_url,
    }
    return render(request, 'employee_pdf_generator_app/home_page.html', context)

# for the dropdown employee id
def get_employee_id(request):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    data = df['E.CODE'].to_json(orient='values')
    employee_data = {'data': data}

    return JsonResponse(employee_data)

# after selecting dropdown option this function will return basic information
def get_emp_data_by_emp_id(request, employee_id):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    
    filtered_row = df[df['E.CODE'] == str(employee_id)]
    employee_data_as_dict = filtered_row.to_dict(orient='records')[0]

    new_refine_dict = {
        'employee_name' : str(employee_data_as_dict['Name of Employe']),
        'employee_dob' : str(employee_data_as_dict['DOB']).split(" ")[0],
        'employee_doj' : str(employee_data_as_dict['DOJ']).split(" ")[0],
        'employee_adhar_no' : str(employee_data_as_dict['AADHARCARD']),
        'employee_id' : str(employee_data_as_dict['E.CODE']),
        'employee_department' : str(employee_data_as_dict['DEPARTMENT']),
        'employee_mnumber' : str(employee_data_as_dict['MOBILE']),
        'employee_address' : str(employee_data_as_dict['PERMENANT ADDRESS']).strip('AT:'),
    }
    # print(new_refine_dict)
    return JsonResponse(new_refine_dict)

# this function is dedicated to fetch single employee information
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
    # print(new_refine_dict)
    
    return new_refine_dict

# this function is dedicated to fetch all employee data
def all_emp_data():
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    data_list = []

    for index, row in df.iterrows():
        row_data = {}
        for column_name, cell_value in row.items():
            row_data[column_name] = str(cell_value)
        data_list.append(row_data)

    # print(data_list)
    
    return data_list

pdfmetrics.registerFont(TTFont("Helvetica-Bold", r"Fonts/Helvetica-Bold-Font.ttf"))

def form_2_generate_pdf_by_id(employee_id):
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

    return buffer

def form_2_generate_pdf(request, employee_id):

    buffer = form_2_generate_pdf_by_id(employee_id)
    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="result.pdf"'

    buffer.close()
    return response

# New view function to generate PDFs of form2 for multiple employees
def form_2_generate_multiple_employee_pdfs(request):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    employee_ids = list(df['E.CODE'])
    zip_buffer = BytesIO()

    # Create a zip file and add PDFs for each employee
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for employee_id in employee_ids:

            print(f"This is employee id, generate_multiple_employee_pdfs: {employee_id}")
            pdf_buffer = form_2_generate_pdf_by_id(employee_id)

            # Add the PDF to the zip file with a unique name (e.g., employee_id.pdf)
            zipf.writestr(f'{employee_id}_form-2.pdf', pdf_buffer.getvalue())

    # Close the zip buffer and prepare the response
    zip_buffer.seek(0)
    response = HttpResponse(zip_buffer.read(), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=all_employee_form-2_pdfs.zip'

    return response

def form_1_generate_pdf_by_id(employee_id):
    buffer = BytesIO()

    template = PdfReader(r"Forms/None_Form1_Declaration_ESI.pdf", decompress=False)
    template_page_count = len(template.pages)
    page_width, page_height = A4
    canvas = Canvas(buffer, pagesize=(page_width, page_height))

    font_size = 8
    font_name = "Helvetica"
    img_flage = False
    temp_image_path = None
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
            # print(f"Employee with code '{target_employee_code}' found in row {row_number}.")
        else:
            img_flage = False
            print(f"Employee with code '{target_employee_code}' not found in the Excel file.")

        if img_flage:
            image = image_loader.get(f'B{row_number}')
            temp_image_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
            image.save(temp_image_path)
        else:
            temp_image_path = None
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

            if ((img_flage == True) and (temp_image_path != None)):
                canvas.drawImage(temp_image_path, 402, 59, width=100, height=80)

        canvas.showPage()

    canvas.save()
    buffer.seek(0)

    return buffer

def form_1_generate_pdf(request, employee_id):
    
    buffer = form_1_generate_pdf_by_id(employee_id)
    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="result.pdf"'

    buffer.close()

    return response

# New view function to generate PDFs of form1 for multiple employees
def form_1_generate_multiple_employee_pdfs(request):
    
    df = pd.read_excel(filename, header=[0])
    df.reset_index()
    employee_ids = list(df['E.CODE'])
    zip_buffer = BytesIO()

    # Create a zip file and add PDFs for each employee
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for employee_id in employee_ids:

            print(f"This is employee id, generate_multiple_employee_pdfs: {employee_id}")
            pdf_buffer = form_1_generate_pdf_by_id(employee_id)

            # Add the PDF to the zip file with a unique name (e.g., employee_id.pdf)
            zipf.writestr(f'{employee_id}_form-1.pdf', pdf_buffer.getvalue())

    # Close the zip buffer and prepare the response
    zip_buffer.seek(0)
    response = HttpResponse(zip_buffer.read(), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=all_employee_form-1_pdfs.zip'

    return response

def form_15_generate_pdf(request):
    buffer = BytesIO()
    
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=40, rightMargin=40, topMargin=10, bottomMargin=10)
    elements = []

    all_employee_data = all_emp_data()

    # Header content
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(name='HeaderStyle', parent=styles['Heading1'])
    header_style.alignment = 1
    header_style.alignment = TA_CENTER
    header_style.fontSize = 16

    content_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    content_style.leading = 14
    content_style.fontSize = 10

    tab_line1_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    tab_line1_style.leftIndent = 167

    tab_line2_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    tab_line2_style.leftIndent = 334.3

    tab_line4_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    tab_line4_style.leftIndent = 214

    bold_style = ParagraphStyle(name='BoldStyle', parent=content_style)
    bold_style.fontName = 'Helvetica-Bold'

    header_text = "FORM XV"
    rule_no_text = "(See rule 88)"
    sub_header_text = "Register of Adult worker"
    line1_text = "<b>1.Name and Address of Contract </b>:- Speed Industrial Service"
    line1_cont1_text = "31,Ranchod Nagar,Godhra Road,"
    line1_cont2_text = "Halol.Dsit-Panchmahal-389350"

    line2_text = "<b>2.Name and Address of establishment which contract is carriaed on </b>:- Raychem RPG Private Limited"
    line2_cont1_text = "Near Safari Crossing,Halol GIDC,Kanjari"
    line2_cont2_text = "Dist-Panchmahal-389350"
    
    line3_text = "<b>3.Nature and location of work </b>:- Loading-Unloading,Lifting Shifting etc."
    
    line4_text = "<b>4.Name and address of principal Employer </b>:- Raychem RPG Private Limited"
    line4_cont1_text = "Near Safari Crossing,Halol GIDC,Kanjari"
    line4_cont2_text = "Dist-Panchmahal-389350"

    header = Paragraph(header_text, header_style)
    sub_header = Paragraph(sub_header_text, header_style)
    rule_no = Paragraph(rule_no_text, header_style) 

    line1 = Paragraph(line1_text, content_style)
    line1_cont_1 = Paragraph(line1_cont1_text, tab_line1_style)
    line1_cont_2 = Paragraph(line1_cont2_text, tab_line1_style)
    line2 = Paragraph(line2_text, content_style)
    line2_cont_1 = Paragraph(line2_cont1_text, tab_line2_style)
    line2_cont_2 = Paragraph(line2_cont2_text, tab_line2_style)
    line3 = Paragraph(line3_text, content_style)
    line4 = Paragraph(line4_text, content_style)
    line4_cont_1 = Paragraph(line4_cont1_text, tab_line4_style)
    line4_cont_2 = Paragraph(line4_cont2_text, tab_line4_style)

    for emp_number in range(0, len(all_employee_data)-1, 6):

        elements.append(header)
        elements.append(rule_no)
        elements.append(sub_header)
        elements.append(Spacer(1, 10))
        elements.append(line1)
        elements.append(line1_cont_1)
        elements.append(line1_cont_2)
        elements.append(Spacer(1, 2))
        elements.append(line2)
        elements.append(line2_cont_1)
        elements.append(line2_cont_2)
        elements.append(Spacer(1, 2))
        elements.append(line3)
        elements.append(Spacer(1, 2))
        elements.append(line4)
        elements.append(line4_cont_1)
        elements.append(line4_cont_2)
        elements.append(Spacer(1, 20))

        try:
            data2 = []
            for i in range(6):
                emp_address = ""
                try:
                    full_address = all_employee_data[emp_number+i]['PERMENANT ADDRESS'].strip('AT:')
                    wrapper1 = textwrap.TextWrapper(width=30)
                    word_list1 = wrapper1.wrap(text=full_address)
                    
                    for element in word_list1:
                        emp_address += element + "\n"
                    emp_address = emp_address.rstrip('\n')
                except Exception as E:
                    print(f"error at address convert: {E}")

                emp_name = ""
                try:
                    full_name = all_employee_data[emp_number+i]['Name of Employe']
                    wrapper2 = textwrap.TextWrapper(width=25)
                    word_list2 = wrapper1.wrap(text=full_name)
                    
                    for element in word_list2:
                        emp_name += element + "\n"
                    emp_name = emp_name.rstrip('\n')
                except Exception as E:
                    print(f"error at address convert: {E}")
                    
                data2.append([f"{emp_number+i+1}", emp_name, all_employee_data[emp_number+i]['DOB'].split(" ")[0], "M", emp_address, all_employee_data[emp_number+i]['FATHER NAME'], all_employee_data[emp_number+i]['DOJ'].split(" ")[0], "", "", "", "", "", ""])
        except Exception as E:
            print(E)

        # body content
        data1 = [
            ["Sr.No", "Name", "Date of\nBirth", "Sex", "Residential Address", "Father's/\nHusband\nname", "Date of\nappointment","Group to which worker\nbelongs","", "No. of\nrelay if\nworking in\nshifts", "Adolescent if certified\nas adults","", "Re-\nmarks"],
            ["", "", "", "", "", "", "", "Alphabet\nassigned","Nature of\nwork", "", "No. & date of\ncertificate\nof fitness", "No. under\nsection 68", ""],
        ]

        data = data1 + data2


        table_style = [
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('SPAN',(0,0),(0,1)),
            ('SPAN',(1,0),(1,1)),
            ('SPAN',(2,0),(2,1)),
            ('SPAN',(3,0),(3,1)),
            ('SPAN',(4,0),(4,1)),
            ('SPAN',(5,0),(5,1)),
            ('SPAN',(6,0),(6,1)),
            ('SPAN',(7,0),(8,0)),
            ('SPAN',(9,0),(9,1)),
            ('SPAN',(10,0),(11,0)),
            ('SPAN',(12,0),(12,1)),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('TEXTCOLOR', (0, 0), (-2, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 1), 2),
            ('LEFTPADDING', (0, 0), (-1, 1), 2),
            ('RIGHTPADDING', (0, 0), (-1, 1), 2),

            ('FONTNAME', (0, 2), (-2, -1), 'Helvetica'),
            ('VALIGN', (0, 2), (-2, -1), 'MIDDLE'),  # Center vertically
            ('FONTSIZE', (0, 2), (-2, -1), 8),
            ('LEFTPADDING', (0, 2), (-2, -1), 6),
            ('RIGHTPADDING', (0, 2), (-2, -1), 6),
        ]

        table = Table(data)
        table.setStyle(TableStyle(table_style))
        elements.append(table)
        elements.append(PageBreak())


    doc.build(elements)

    buffer.seek(0)

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="register_of_workmen.pdf"'
    response.write(buffer.read())
    buffer.close()

    return response

def form_13_generate_pdf(request):
    buffer = BytesIO()
    
    doc = SimpleDocTemplate(buffer, pagesize=(A4), leftMargin=40, rightMargin=40, topMargin=10, bottomMargin=10)
    elements = []

    all_employee_data = all_emp_data()

    # Header content
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(name='HeaderStyle', parent=styles['Heading1'])
    header_style.alignment = 1
    header_style.alignment = TA_CENTER
    header_style.fontSize = 16

    content_style = styles["Normal"]
    content_style.leading = 14
    content_style.fontSize = 10

    tab_line1_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    tab_line1_style.leftIndent = 162.5

    tab_line4_style = ParagraphStyle(name='ContentStyle', parent=styles["Normal"])
    tab_line4_style.leftIndent = 201.8

    bold_style = ParagraphStyle(name='BoldStyle', parent=content_style)
    bold_style.fontName = 'Helvetica-Bold'

    header_text = "FORM XIII"
    rule_no_text = "(See rule 75)"
    sub_header_text = "Register of workmen employed by Contractor"
    line1_text = "<b>Name and address of contractor</b>:- Speed Industrial Service"
    line1_cont1_text = "31,Ranchod Nagar,Godhra Road,"
    line1_cont2_text = "Halol.Dsit-Panchmahal-389350"
    line2_text = "<b>Nature and location of work</b>:- Raychem RPG (P) Ltd.:- Loading-unloading, packing, shifting & Lifting Work etc"
    line3_text = "<b>Name and address of establishment in/under which contract is carried on</b>:-"
    line4_text = "<b>Name and address of principal employer</b>:- Raychem RPG (P) Limited"
    line4_cont1_text = "Near Safari Crossing,Halol GIDC,Kanjari"
    line4_cont2_text = "Dist-Panchmahal-389350"

    header = Paragraph(header_text, header_style)
    sub_header = Paragraph(sub_header_text, header_style)
    rule_no = Paragraph(rule_no_text, header_style) 

    line1 = Paragraph(line1_text, content_style)
    line1_cont_1 = Paragraph(line1_cont1_text, tab_line1_style)
    line1_cont_2 = Paragraph(line1_cont2_text, tab_line1_style)
    line2 = Paragraph(line2_text, content_style)
    line3 = Paragraph(line3_text, content_style)
    line4 = Paragraph(line4_text, content_style)
    line4_cont_1 = Paragraph(line4_cont1_text, tab_line4_style)
    line4_cont_2 = Paragraph(line4_cont2_text, tab_line4_style)

    elements.append(header)
    elements.append(rule_no)
    elements.append(sub_header)
    elements.append(Spacer(1, 10))  # Add some space
    elements.append(line1)
    elements.append(line1_cont_1)
    elements.append(line1_cont_2)
    elements.append(Spacer(1, 2))
    elements.append(line2)
    elements.append(line3)
    elements.append(line4)
    elements.append(line4_cont_1)
    elements.append(line4_cont_2)
    elements.append(Spacer(1, 20))

    data2 = []
    for emp_number in range(0, len(all_employee_data)):
        try:
            emp_name = ""
            try:
                full_name = all_employee_data[emp_number]['Name of Employe']
                wrapper = textwrap.TextWrapper(width=25)
                word_list = wrapper.wrap(text=full_name)
                
                for element in word_list:
                    emp_name += element + "\n"
                emp_name = emp_name.rstrip('\n')
            except Exception as E:
                print(f"error at address convert: {E}")
                
            data2.append([f"{emp_number+1}", emp_name, "M", all_employee_data[emp_number]['FATHER NAME'], "Helper", all_employee_data[emp_number]['DOJ'].split(" ")[0], "", "", "", ""])
        except Exception as E:
            print(E)


    # body content
    data1 = [
        ["Sr.No", "Name and surname\nof workman", "Ag.\nand\nsex", "Father's/\nHusband\nname", "Nature of\nemployment/\ndesignation", "Date of\ncommen-\ncement of\nemployment", "Sign-\nature\nof\nwork\nman", "Date\nof\ntermi-\nnation", "Reasons\nfor\ntermi-\nnation", "Re-\nmarks"],
    ]

    data = data1 + data2


    table_style = [
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('TEXTCOLOR', (0, 0), (-2, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
        ('LEFTPADDING', (0, 0), (-1, 0), 2),
        ('RIGHTPADDING', (0, 0), (-1, 0), 2),

        ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),
        ('VALIGN', (0, 2), (-1, -1), 'MIDDLE'),  # Center vertically
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('LEFTPADDING', (0, 2), (-1, -1), 6),
        ('RIGHTPADDING', (0, 2), (-1, -1), 6),
    ]

    table = Table(data)
    table.setStyle(TableStyle(table_style))
    elements.append(table)

    doc.build(elements)

    buffer.seek(0)

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="register_of_workmen.pdf"'
    response.write(buffer.read())
    buffer.close()

    return response

