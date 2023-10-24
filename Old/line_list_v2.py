import os, re, openpyxl

# INPUT FILE 
current_path=os.getcwd()
input_file_name='Line_list_base.xlsx'
input_file_path=os.path.join(current_path,input_file_name)
input_wb=openpyxl.load_workbook(input_file_path)
input_ws=input_wb.active    

# TEMPLATE FILE 
template_file_name='line_list_template.xlsx'
template_file_path=os.path.join(current_path,template_file_name)
template_wb=openpyxl.load_workbook(template_file_path)
template_ws=template_wb.active    


# Cargo las lineas filtrando con regex. OJO LA EXPRESION DE REGEX ES MUY VAGA. PODRIA COMPLICARCLA.
lineas=[]
medidas={'OD8mm': '8', 'OD12mm': '12', '1/8"': 6, '1/4"': 8, '3/8"': 10, '1/2"': 15, '3/4"': 20, '1"': 25, '1 1/4"': 32, '1 1/2"': 40, '2"': 50, '2 1/2"': 65, '3"': 80, '4"': 100, '6"': 150, '8"': 200, '10"': 250}
patron=r'-\d{4}-'
# Patron solo filtra cadenas que tengan 4 numeros encerrados por guiones. 
for row in input_ws['C']:
    if re.search(patron,row.value):
        lineas.append(row.value)
#print(lineas) Ahoa linea guarda todas las lineas en forma de lista.

row_index=11
for linea in lineas:
    # try:
    template_ws.cell(row=row_index, column=4, value=linea.split('-')[0]) # dimension
    template_ws.cell(row=row_index, column=5,value=linea.split('-')[1]) # fluido
    template_ws.cell(row=row_index, column=16,value=linea.split('-')[1]) # fluido
    template_ws.cell(row=row_index, column=6,value=linea.split('-')[2]) # tag
    template_ws.cell(row=row_index, column=7,value=linea.split('-')[3]) # clase cañeria
    if len(linea.split("-")) > 4:
        template_ws.cell(row=row_index, column=8, value=linea.split('-')[4]) # aislamiento
    else:
        template_ws.cell(row=row_index, column=8, value="-")
    if "L" in linea.split('-')[1]:
        template_ws.cell(row=row_index, column=17, value='LIQ')

    template_ws.cell(row=row_index,column=35,value='NO')   # Ver si no hay que poner algún condicional para estas líneas. 
    template_ws.cell(row=row_index,column=34,value='1')
    template_ws.cell(row=row_index,column=36,value='6')

    if medidas.get(linea.split('-')[0]) is not None:
        if  float(medidas.get(linea.split('-')[0]))<= 25:
            template_ws.cell(row=row_index, column=38, value='Sound Engineering Practice')
            template_ws.cell(row=row_index, column=37, value='4.3') 

      
# pdis_value = float(cell_Pdis.value)  # Convert cell_Pdis.value to a float
#                 if nps_value <= 100 and nps_value * pdis_value <= 1000:
#                     wsheet_ll.cell(row=row_index, column=37, value='1')
#                     wsheet_ll.cell(row=row_index, column=38, value='A')
#                 elif 25 < nps_value <= 350 and nps_value * pdis_value <= 3500:
#                     wsheet_ll.cell(row=row_index, column=37, value='2')
#                     wsheet_ll.cell(row=row_index, column=38, value='A2,D1,E1')
#                 elif (nps_value > 350 and pdis_value <= 10) or (pdis_value < 10 and nps_value * pdis_value > 3500):
#                     wsheet_ll.cell(row=row_index, column=37, value='3')








    
    row_index+=1
template_wb.save('line_list_MOD.xlsx')


#     FALTA  HACER LA PARTE DE LOS CONDICIONALES Y COMPROBAR QUE FUNCIONE TODO JUNTO. 