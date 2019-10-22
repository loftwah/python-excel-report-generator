from .reportPyLibrary import Report
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.http import FileResponse
import pandas as pd

class ExcelExport(APIView):
    def get(self, request, format=None):
        return Response({"success": True}, status=status.HTTP_400_BAD_REQUEST)

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

            excelReport = reportObj.exportToExcel()

            response = FileResponse(excelReport, content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=ExcelReport'

            return response

        else:
            return Response({"success": False}, status=status.HTTP_400_BAD_REQUEST)
