# python-excel-report-generator
python-excel-report-generator is a microservice written in python to write an excel report in xlsx file. It is born from the lack of an existing library to write natively from Python the Office Open XML format. It can be accessed from any other programming framework using post request. 

### Framework and languages
* Django 2.2
* python 3.6
* pandas
* Django Rest Framework
* openpyxl

### How to use this service
The service accepts a JSON object with two keys from the post request. The first key, "header," will be a list or array. The header must contain the requirements of the excel report such as cell information, alignment, font size, font family etc. The second key is "df," where the data will send in the form of JSON, dictionary, or data frame object. 
