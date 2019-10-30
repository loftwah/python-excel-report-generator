from django.urls import path
from django.conf.urls import url

from . import views
from . import api

urlpatterns = [
    # path('', views.index, name='index'),
    path('', views.ExcelView.as_view()),
    url(r'^api/excel_export$', api.ExcelExport.as_view(), name='ExcelExportAPI'),
    url(r'^api/excel_export$', api.ExcelExport.as_view(), name='ExcelExportAPI2'),
]