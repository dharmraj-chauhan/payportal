from django.contrib import admin
from django.urls import path

from employee_pdf_generator_app.views import home_page, get_employee_id, get_emp_data_by_emp_id, generate_pdf

urlpatterns = [
    path('', home_page, name='home_page'),
    path('employee-id', get_employee_id, name='employee-id'),
    path('<str:employee_id>/employee-data/', get_emp_data_by_emp_id, name='velocity-data'),
    path('<str:employee_id>/pdf/', generate_pdf, name='velocity-data'),
]
