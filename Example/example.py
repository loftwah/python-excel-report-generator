import requests
import json

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
data = {"header": test_head_List, "df": JsonDf}
# if the data is dataframe the it can be send by converting in a dictionary
# data = {"header": test_head_List, "df": reportDf.to_dict()}

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)

with open('report.xlsx', 'wb') as f:
    f.write(excelReport.content)

f.close()
