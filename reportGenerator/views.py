from django.shortcuts import render

# Create your views here.
from django.views.generic import TemplateView

class ExcelView(TemplateView):
    template_name = "reportGenerator/excel_home.html"