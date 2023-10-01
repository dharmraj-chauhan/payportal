from django.contrib import admin
from django.urls import path

from employee_pdf_generator_app.views import home_page, get_employee_id, get_emp_data_by_emp_id, form_2_generate_pdf, form_1_generate_pdf

urlpatterns = [
    path('', home_page, name='home_page'),
    path('employee-id', get_employee_id, name='employee-id'),
    path('<str:employee_id>/employee-data/', get_emp_data_by_emp_id, name='employee-data'),
    path('<str:employee_id>/form-2-revised/', form_2_generate_pdf, name='form-2-revised-data'),
    path('<str:employee_id>/form-1/', form_1_generate_pdf, name='form-1'),
]
