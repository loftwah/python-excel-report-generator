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
from .head_map import ExcelDataProcessing, head, dataframe
import re
import json

class ExcelExport(APIView):
    def get(self, request, format=None):
        topHeader = [{
                        'column': 'A1:R1',
                        'title': 'Md. Ataur Rahman Bhuiyan',
                        'font': {
                            'font_size': '16',
                            'font_family': 'Calibri',
                            'bold': 'True',
                            'underline': 'none',
                            'color': 'FF000000'
                        },
                        'alignment': {
                            'horizontal': 'center',
                            'vertical': 'center'
                        }
                    }, {
                        'column': 'A2:R2',
                        'title': 'roomeyrahman@gmail.com',
                        'font': {
                            'font_size': '14',
                            'font_family': 'Calibri',
                            'bold': 'True',
                            'italic': 'False',
                            'underline': 'none',
                            'color': 'FF000000'
                        },
                        'alignment': {
                            'horizontal': 'center',
                            'vertical': 'center'
                        }
                    }]

        max_row = -1
        for item in topHeader:
            if type(item) != dict:
                return Response({"success": False, "message": "topHeader's value must be a dictionary"},
                                status=status.HTTP_400_BAD_REQUEST)

            try:
                cell = item["column"]
                cell_splt = cell.split(':')
                cell_l = cell_splt[0]
                cell_r = cell_splt[1]

                max_row = max(max_row, int(re.match(r"([a-z]+)([0-9]+)", cell_l, re.I).groups()[1]))
                max_row = max(max_row, int(re.match(r"([a-z]+)([0-9]+)", cell_r, re.I).groups()[1]))
            except:
                return Response({"success": False, "message": "column is not specify in topHeader's value"},
                                status=status.HTTP_400_BAD_REQUEST)

        ##############Test Object##################
        excelMap = ExcelDataProcessing(head, dataframe, rowStart= max_row+2)

        header = excelMap.header
        df = excelMap.dataframe

        for item in topHeader:
            header.append(item)

        reportObj = Report(df = df, header = header)
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
                # if type(tableData) == dict:
                #     for item in tableData:
                #         print(type(item))
                #         # if type(tableData[item])!=list:
                #             # return Response({"success": False, "message": message}, status=status.HTTP_400_BAD_REQUEST)
                #         break
                # elif type(tableData)!= list:
                #     raise Exception(message)

                if type(tableData)!=list:
                    if type(tableData)!=dict:
                        return Response({"success": False, "message": message}, status=status.HTTP_400_BAD_REQUEST)


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
                # raise Exception("Column Head is undefined or not properly Set")
                return Response({"success": False, "message": "Column Head is undefined or not properly Set"}, status=status.HTTP_400_BAD_REQUEST)

            max_row = -1
            try:
                topHeader = data["topHeader"]
                print(topHeader)
                if type(topHeader) != list:
                    return Response({"success": False, "message": "topHeader must be a list"}, status=status.HTTP_400_BAD_REQUEST)

                for item in topHeader:
                    if type(item) != dict:
                        return Response({"success": False, "message": "topHeader's value must be a dictionary"}, status=status.HTTP_400_BAD_REQUEST)

                    try:
                        cell = item["column"]
                        cell_splt = cell.split(':')
                        cell_l = cell_splt[0]
                        cell_r = cell_splt[1]

                        max_row = max(max_row, int(re.match(r"([a-z]+)([0-9]+)", cell_l, re.I).groups()[1]))
                        max_row = max(max_row, int(re.match(r"([a-z]+)([0-9]+)", cell_r, re.I).groups()[1]))
                    except:
                        return Response({"success": False, "message": "column is not specify in topHeader's value"}, status=status.HTTP_400_BAD_REQUEST)

            except:
                pass


            excelMap = ExcelDataProcessing(head, tableData, headType, rowStart= max_row +2)
            header = excelMap.header
            df = excelMap.dataframe

            if "topHeader" in data:
                for item in data["topHeader"]:
                    header.append(item)

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
                # raise Exception("excelReport server Error")
                return Response({"success": False, "message": "A"}, status=status.HTTP_400_BAD_REQUEST)

        else:
            return Response({"success": False}, status=status.HTTP_400_BAD_REQUEST)
