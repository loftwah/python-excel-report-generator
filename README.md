# python-excel-report-generator
python-excel-report-generator is a microservice written in python to write an excel report in xlsx file. It is born from the lack of an existing library to write natively from Python the Office Open XML format. It can be accessed from any other programming framework using post request. 

### Framework and languages
* Django 2.2
* python 3.6
* pandas
* Django Rest Framework
* openpyxl

### How to use this service
The service accepts a JSON object with two keys from the post request. The first key, "header," will be a list of json object or dictionary. The header must contain the requirements of the excel report such as cell information, alignment, font. A cell information and its requirment can spcify by the following way:
```python
[{
  "column": "A1:A3",
  "title": "Cell information"
}]
```
'column' A1:A3 will merge the column 1 to 3 of cell A in the excel and then the title value 'Cell information' will kept in this cell. Some more paremeter could be passed in the header. The following code will show how to adjust font and alignment of a cell. By default, font size is 11, font family is Calibri, Boldface, Italic and underline is false, and color is black. 

```python
head_list = [{
    "column": "A11:A13",
    "title": "Cell value",
    "font": {
        "font_size": "11",
        'font_family': "Calibri",
        "bold": True,
        "italic": False,
        "underline": "none",
        "color": "FF000000"
    },
    "alignment": {
        "horizontal": "center",
        "vertical": "center"
    }
}]
```

The second key is "df," where the data will send in the form of JSON, dictionary, or data frame object. A json object of dataframe could be created by the following rules:

```python
dfJson = [{
        'Title': 'Project Introduction',
        'Target': 100,
        'Acheive': 90
    },
    {
        'Title': 'Project Organization',
        'Target': 100,
        'Acheive': 90
    },
]
```
If the dataframe will send as json object, than it shoulb be dumps in json from python. It is required to import json library to dumps in json object.
```python
import json
JsonDf = json.dumps(dfJson)
```
Now the data could be prepare for the api request. 

###### url = "http://excel.iofact.com/api/excel_export"
###### request method = post
###### data type = json

```python
excelReport = requests.post("http://excel.iofact.com/api/excel_export", json={"header": head_List, "df": JsonDf})
```


#### Sample Data preparation (API testing)
* simple Api Data
```json
data = { 
   "header":[ 
      { 
         "column":"A1:A3",
         "title":"Budget"
      },
      { 
         "column":"B1:B3",
         "title":"Events"
      }
   ],
   "df":{ 
      "Budget":[ 
         10000,
         15000,
         20000
      ],
      "Events":[ 
         "A",
         "B",
         "C"
      ]
   }
}

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)
```



#### Code Example
```python
import requests
import json

test_head_List = [
        {'column': 'A11:A13',
         'title': 'Activities/Head of Expenditure',
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
  
def test(data):
    excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)

    with open('report.xlsx', 'wb') as f:
        f.write(excelReport.content)

    f.close()

    report = open('report.xlsx', 'rb')
    response = FileResponse(report, content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment; filename=%s' % smart_str('report.xlsx')
    
    return response
    
```
