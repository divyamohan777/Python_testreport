from xml.dom import minidom
from json_excel_converter import Converter 
from json_excel_converter.xlsx import Writer
import plotly
import plotly.express as px
import pandas as pd

# parse an xml file by name
mydoc = minidom.parse('results.xml')
items = mydoc.getElementsByTagName('collection')
dict =[]
for elem in items:
    name =elem.attributes['name'].value
    modifiedname = name.split(".")[3]+":"+name.split(".")[4]
    dictlist = {"name" : modifiedname , "passed": elem.attributes['passed'].value,"failed":elem.attributes['failed'].value,"total": elem.attributes['total'].value}
    dict.append(dictlist)
conv = Converter()
conv.convert(dict, Writer(file='TestResults.xlsx'))
xlWorkbook = "TestResults.xlsx"
df = pd.read_excel(xlWorkbook)
fig = px.bar(df, x=df['name'], y=[df['passed'],df['failed']], title="Test Case Summary",labels={'value':'Test Cases Count'} , barmode='stack')
fig.show()
plotly.offline.plot(fig, filename="SnowTestCaseReport.html")

    
    


