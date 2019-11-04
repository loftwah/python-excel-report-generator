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
from .head_map import ExcelDataProcessing, head

class ExcelExport(APIView):
    def get(self, request, format=None):
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

        ##############Test Object##################
        A = ExcelDataProcessing(head)
        reportObj = Report(jsonObject=JsonDf, header = A.header)
        ###########################################



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

            try:
                tableData = data["tableData"]
                message = """cell data can be send either columnwise or rowwise. If you want to send row wise data then tableData must be a list of dictionary.
                            Otherwise tabelData must be dictionary. Each item of the dictionary must be list of column value."""
                if type(tableData) == dict:
                    for item in tableData:
                        if tableData[item]!=list:
                            raise Exception(message)
                        break
                elif type(tableData)!= list:
                    raise Exception(message)

            except:
                raise Exception("tableData is undefined or not properly set")

            try:
                if "columnHeader" in data:
                    head = data["columnHeader"]
                    headType = 0

                elif "explicitColumnHeader" in data:
                    head = data["explicitColumnHeader"]
                    headType = 1

                if type(head) != list:
                    raise Exception("columnHeader must be a list of dictionary")

            except:
                raise Exception("Column Head is undefined or not properly Set")

            excelMap = ExcelDataProcessing(head, tableData, headType)
            header = excelMap.header
            df = excelMap.dataframe

            style = data.get('style', '')

            if type(style) == str and style =='':
                if isinstance(df, pd.DataFrame):
                    reportObj = Report(df = df, header = header)
                else:
                    reportObj = Report(jsonObject=df, header=header)
            else:
                font = style.get('font', '')
                fill = style.get('fill', '')
                border = style.get('border', '')
                alignment = style.get('alignment', '')
                number_format = style.get('number_formate', '')
                protection = style.get('protection', '')

                if isinstance(df, pd.DataFrame):
                    reportObj = Report(df = df, header = header, font = font, fill = fill, border=border, alignment=alignment, number_format=number_format, protection=protection)
                else:
                    reportObj = Report(jsonObject=df, header=header, font = font, fill = fill, border=border, alignment=alignment, number_format=number_format, protection=protection)


            try:
                excelReport = reportObj.exportToExcel()
                response = FileResponse(excelReport, content_type='application/ms-excel')
                response['Content-Disposition'] = 'attachment; filename=ExcelReport'
                return response

            except:
                raise Exception("excelReport server Error")

        else:
            return Response({"success": False}, status=status.HTTP_400_BAD_REQUEST)
