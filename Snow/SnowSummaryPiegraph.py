from xml.dom import minidom
from json_excel_converter import Converter 
from json_excel_converter.xlsx import Writer
import plotly
import plotly.graph_objects as go
import pandas as pd
from plotly.subplots import make_subplots

# parse the xml file
mydoc = minidom.parse('results.xml')
items = mydoc.getElementsByTagName('collection')
# make it a python dictionary
dict =[]
for elem in items:
    name =elem.attributes['name'].value
    modifiedname = name.split(".")[3]+":"+name.split(".")[4]
    dictlist = {"name" : modifiedname , "passed": elem.attributes['passed'].value,"failed":elem.attributes['failed'].value,"total": elem.attributes['total'].value}
    dict.append(dictlist)
conv = Converter()
#Convert to excel
conv.convert(dict, Writer(file='TestResults.xlsx'))
xlWorkbook = "TestResults.xlsx"
# create dataframe from excel
df = pd.read_excel(xlWorkbook)
#pivot table for fidning total num of test cases passed/failed
df_pivot=df.pivot_table(index='name',margins=True,margins_name='Total',aggfunc=sum)
df_labels = ['Passed','Failed']
df_values=[df_pivot.at['Total','passed'],df_pivot.at['Total','failed']]
#Subplots in html page in 2D matrix
fig =make_subplots(rows=1,cols=3,specs=[[{'type':'domain','colspan': 2},None,None]],vertical_spacing=0.002,subplot_titles=['Snow -Test Case Summary'])
fig.add_trace(go.Pie(labels=df_labels,values=df_values,marker_colors=['rgb(55, 69, 219)','rgb(32, 40, 133)'],textinfo='label+value'),1,1)
fig.update_traces(hole=.4,hoverinfo="label+percent+name",textposition='inside')
fig['layout'].update(height=800)
#html report generation
plotly.offline.plot(fig, filename="piegrpah.html")


