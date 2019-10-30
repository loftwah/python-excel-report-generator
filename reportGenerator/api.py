"""
Author: Roomey Rahman
mail: roomeyrahman@gmail.com
"""

from .reportPyLibrary import Report
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.http import FileResponse
import pandas as pd
import json

class ExcelExport(APIView):
    def get(self, request, format=None):
        print("+++++")
        test_head_List = [
            {'column': 'A11:A13',
             'title': 'Activities',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'B11:B13',
             'title': 'Target Participants',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'C11:I11',
             'title': 'Participants breakdown',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'C12:C13',
             'title': 'Male',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'D12:D13',
             'title': 'Female',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'E12:E13',
             'title': 'Boy',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'F12:F13',
             'title': 'Girl',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'G12:I12',
             'title': 'Total',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'G13:G13',
             'title': 'Male',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'H13:H13',
             'title': 'Female',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},

            {'column': 'I13:I13',
             'title': 'Total',
             'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}},
        ]

        dfJson = [
            {
                'Activities': 'Project Introduction',
                'Target': 100,
                'Persons Breakdown.Male': 10,
                'Persons Breakdown.Female': 10,
                'Persons Breakdown.Boy': 10,
                'Persons Breakdown.Girl': 10,
                'Persons Breakdown.Total.Male': 20,
                'Persons Breakdown.Total.Female': 20,
                'Persons witBreakdownh.Total.Total': 40,
            },
            {
                'Activities': 'Event Management',
                'Target': 100,
                'Persons Breakdown.Male': 10,
                'Persons Breakdown.Female': 10,
                'Persons Breakdown.Boy': 10,
                'Persons Breakdown.Girl': 10,
                'Persons Breakdown.Total.Male': 20,
                'Persons Breakdown.Total.Female': 20,
                'Persons witBreakdownh.Total.Total': 40,
            },
            {
                'Activities': 'Project Inspection',
                'Target': 100,
                'Persons Breakdown.Male': 10,
                'Persons Breakdown.Female': 10,
                'Persons Breakdown.Boy': 10,
                'Persons Breakdown.Girl': 10,
                'Persons Breakdown.Total.Male': 20,
                'Persons Breakdown.Total.Female': 20,
                'Persons witBreakdownh.Total.Total': 40,
            }
        ]
        JsonDf = json.dumps(dfJson)

        reportObj = Report(jsonObject=JsonDf, header=test_head_List)

        try:
            excelReport = reportObj.exportToExcel()
            response = FileResponse(excelReport, content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=ExcelReport'
            return excelReport
        except:
            return Response({"success": False}, status=status.HTTP_400_BAD_REQUEST)

    def post(self, *args, **kwargs):
        if self.request.method == "POST":
            data = self.request.data
            header = data["header"]
            df = data["df"]


            if type(df) == dict or isinstance(df, pd.DataFrame):
                if type(df) == dict:
                    df = pd.DataFrame(df)
                reportObj = Report(df = df, header = header)
            else:
                reportObj = Report(jsonObject=df, header=header)

            try:
                excelReport = reportObj.exportToExcel()
                response = FileResponse(excelReport, content_type='application/ms-excel')
                response['Content-Disposition'] = 'attachment; filename=ExcelReport'
            except:
                pass

            return response

        else:
            return Response({"success": False}, status=status.HTTP_400_BAD_REQUEST)
