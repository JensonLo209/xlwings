import xlwings as xw
sheet = xw.Book('4.xlsx').sheets[0]
#新增chart
chart = sheet.charts.add()                        
#数据源：sheet.range('A1:B7')，或者sheet.range('A1').expand()
chart.set_source_data(sheet.range('A1').expand())  
chart.chart_type = 'line' 
#设置图标的类型，此处为线型，具体的类型查看office官网VBA操作的手册
#标题名称
title='python知识学堂粉丝数'                    
chart.api[1].SetElement(2)
#设置标题名称
chart.api[1].ChartTitle.Text =title          
chart.api[1].SetElement(302)                  #横线
#横轴标题名称
chart.api[1].Axes(1).AxisTitle.Text = "日期"  
chart.api[1].SetElement(311)
chart.api[1].Axes(2).AxisTitle.Text = "粉丝数" #纵轴标题名称
