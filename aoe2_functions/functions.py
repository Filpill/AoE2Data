from json import loads
import os
import sys
import requests
import xlsxwriter

def api2df(requestAddr):

    #Making the API Request
    print(f'Making API Call to: {requestAddr}')
    response = requests.get(requestAddr)
    if response.status_code == 200:
        print('API Request Successful')
    elif response.status_code == 404:
        print('API Request Failed')
        KeyboardInterrupt

    #Converting JSON to Dict and Parsing API Data into a Dataframe
    dict = loads(response.text)

    return dict



#Function to Detect Operating System and Adjust Pathing to Respective Filesystem
def pathing(folder_path,filename):

    #Windows Operating System
    if 'win' in sys.platform:
        if folder_path == 'root':
            filepath = f'{sys.path[0]}\\{filename}'
        else:
            filepath = f'{sys.path[0]}\\{folder_path}\\{filename}'
    #Linux/Mac Operating Sytem
    else:
        if folder_path == 'root':
            filepath = f'{sys.path[0]}/{filename}'
        else:
            filepath = f'{sys.path[0]}/{folder_path}/{filename}'

    return filepath



# Folder Clean-up Process - Removing all the saved png's
def png_cleaner():
    for picture in os.listdir(sys.path[0]):
        if picture.endswith('.png'):
            os.remove(picture)



#Function For Annotating Charts
def dataLabel_stacked(table,col_val_list,text_offset,text_format,fontsize,ax):

    for i in range(table.shape[0]):
        #Determining Text Position for Bottom Stack
        value = table.iat[i,table.columns.get_loc(col_val_list[0])]
        if value != 0:
            if text_format == 'integer':
                value = int('%1g' % value)
            else:
                value = '{:.0%}'.format(value)
            pos = table.iat[i,table.columns.get_loc(col_val_list[0])]/2 - text_offset
            ax.text(i,pos,value,fontsize=fontsize,color='white',weight='bold',ha="center")

        #Determining Text Position for Top Stack
        value = table.iat[i,table.columns.get_loc(col_val_list[1])]
        if value != 0:
            if text_format == 'integer':
                value = int('%1g' % value)
            else:
                value = '{:.0%}'.format(value)

            pos = table.iat[i,table.columns.get_loc(col_val_list[1])]/2 + table.iat[i,table.columns.get_loc(col_val_list[0])] - text_offset
            ax.text(i,pos,value,fontsize=fontsize,color='white',weight='bold',ha="center")



#Function for Writing .xlsx Data Tables which will provide a standard output for table formatting 
def xlsx_table_writer(data_table,sheet_name,col_width_list,title_str,startrow,startcol,worksheet,workbook,writer):

    #Layout/Formatting
    t_vert_spacing = 5 #Vertical Spacing between tables
    t_horz_spacing = 1 #Horizontal Spacing between tables
    title_format = workbook.add_format({'bold': True, 'font_size' : 20, 'fg_color' : '#76933C', 'font_color' : 'white' }) #Standardised Title Format for all tables
    header_format = workbook.add_format({'bold' : True, 'font_size' : 12, 'text_wrap' : True}) #text-wrapping for table headers
    
    # Adding and changing active sheet
    try:
        worksheet=workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet 
        worksheet.set_zoom(70)
        startrow = 3
        startcol = 1
    except:
        pass

    for i in range(len(col_width_list)):
        worksheet.set_column(i+startcol, i+startcol, col_width_list[i])

    #Writing in Cell Data and Merging Cells for Table Titles
    data_table.to_excel(writer,sheet_name=sheet_name,startrow=startrow , startcol=startcol, index=False, header=False)
    worksheet.merge_range(startrow-2, startcol,startrow-2,data_table.shape[1] + startcol -1, title_str,title_format) #writing in title formatting above table

    #Column settings to use in add table function
    column_settings = [{'header' : column} for column in data_table.columns]

    #Populating Excel with Table Format - Adding table to xls for each df
    worksheet.add_table(startrow-1, startcol, startrow + data_table.shape[0], data_table.shape[1] + startcol - 1, {'columns' : column_settings, 'style': 'Table Style Medium 4', 'autofilter' : False})   

    #Applying a text wrap to the Column Header
    for col_num, value in enumerate(data_table.columns.values):
        worksheet.write(startrow-1, col_num + startcol, value, header_format)
    
    #Setting Positions of Following Tables Insertions
    startrow = startrow + data_table.shape[0] + t_vert_spacing #Setting to start row for next table
    # startrow = 3  
    # startcol = startcol + data_table.shape[1] + t_horz_spacing  #Disabling horizontally displaced tables in favour of vertically displacements
    startcol = 1

    #Setting the column width at the end of the table to keep to spacing minimal between tables
    worksheet.set_column(startcol-1,startcol-1,1)
    
    #Return the start row in order to index for future function calls
    return [startrow,startcol,worksheet,workbook,writer]
    


#Plotting the Red Bull Reporting into series of Charts
def xlsx_chart(v_idx,v_space,chart_path,sheet_name,worksheet,workbook,writer):

    chart_bg_format = workbook.add_format({'fg_color' : 'white', 'font_color' : 'black' })

    # Adding and changing active sheet
    try:
        worksheet=workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet 
        worksheet.set_zoom(70)

        #Conditional Formatting For the Charts Spreadsheet
        worksheet.conditional_format(0,0,1000,1000,{'type':'cell',
                                                    'criteria': '=',
                                                    'value': r'""',
                                                    'format': chart_bg_format})
    except:
        pass

    #Chart Data Added to Workbook
    worksheet.insert_image(v_idx,1,chart_path)

    #Setting Chart Spacing For the Next Chart Insertion
    v_idx = v_idx + v_space

    #Return v_idx and graph path ahead of next chart production
    return v_idx,sheet_name,worksheet