from django.contrib import admin
from django.urls import path

from employee_pdf_generator_app.views import home_page, get_employee_id, get_emp_data_by_emp_id, form_2_generate_pdf, form_1_generate_pdf, form_15_generate_pdf, form_13_generate_pdf, form_2_generate_multiple_employee_pdfs, form_1_generate_multiple_employee_pdfs

urlpatterns = [
    path('', home_page, name='home_page'),
    path('employee-id', get_employee_id, name='employee-id'),
    path('<str:employee_id>/employee-data/', get_emp_data_by_emp_id, name='employee-data'),
    path('<str:employee_id>/form-2-revised/', form_2_generate_pdf, name='form-2-revised-data'),
    path('all_employee/form-2-revised-all/', form_2_generate_multiple_employee_pdfs, name='form-2-revised-data-all'),
    path('all_employee/form-1-all/', form_1_generate_multiple_employee_pdfs, name='form-1-data-all'),
    path('<str:employee_id>/form-1/', form_1_generate_pdf, name='form-1'),
    path('form-15/', form_15_generate_pdf, name='form-15'),
    path('form-13/', form_13_generate_pdf, name='form-13'),
    
]