import xlwt
from datetime import datetime, date
import calendar
import os
from REnergy.settings import BASE_DIR
from django.http import HttpResponse, JsonResponse
from costestimator.models import Cost_CNG, Cost_LPG, Cost_DG, Cost_PNG, Cost_Petrol, cost_water, Cost_Wind
from meter_data.models import Masterdatatable
from django.db.models.aggregates import Sum
from source_management.models import AddSource
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import json
from django.views.decorators.csrf import csrf_exempt
from encrypt.renderers import CustomAesRenderer
from REnergy.settings import BASE_DIR,EMAIL_HOST_USER,EMAIL_HOST,EMAIL_PORT,EMAIL_HOST_PASSWORD  

@csrf_exempt
def carbonexcel(request):
    
    plantname = request.GET['plantname']
    
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet1 = workbook.add_sheet('Scope 1', cell_overwrite_ok = True)
    worksheet2 = workbook.add_sheet('Scope 2', cell_overwrite_ok = True)
    worksheet3 = workbook.add_sheet('Scope 3')
    worksheet4 = workbook.add_sheet('Water details', cell_overwrite_ok = True)
    worksheet5 = workbook.add_sheet('Electricity & Cost Details')
    worksheet6 = workbook.add_sheet('I-IV Consumption & Waste Dis.')
    worksheet7 = workbook.add_sheet('VIII - Goods transport')
    worksheet8 = workbook.add_sheet('V - VII Employee Commuting')
    
    worksheet1.show_grid = False
    worksheet2.show_grid = False
    worksheet3.show_grid = False
    worksheet4.show_grid = False
    worksheet5.show_grid = False
    worksheet6.show_grid = False
    worksheet7.show_grid = False
    worksheet8.show_grid = False
    
    style1 = xlwt.XFStyle()
    style2 = xlwt.XFStyle()
    style3 = xlwt.XFStyle()
    style4 = xlwt.XFStyle()
    style5 = xlwt.XFStyle()
    style6 = xlwt.XFStyle()
    style7 = xlwt.XFStyle()
    style8 = xlwt.XFStyle()
    style9 = xlwt.XFStyle()
    style10 = xlwt.XFStyle()
    style11 = xlwt.XFStyle()
    
    alignment1 = xlwt.Alignment()
    alignment1.vert = xlwt.Alignment.VERT_CENTER
    alignment1.horz = xlwt.Alignment.HORZ_CENTER
    alignment1.wrap = 1
    style1.alignment = alignment1
    style2.alignment = alignment1
    style3.alignment = alignment1
    style4.alignment = alignment1
    style5.alignment = alignment1
    style6.alignment = alignment1
    style7.alignment = alignment1
    style8.alignment = alignment1
    style11.alignment = alignment1
    
    alignment2 = xlwt.Alignment()
    alignment2.vert = xlwt.Alignment.VERT_CENTER
    alignment2.wrap = 1
    style9.alignment = alignment2
    style10.alignment = alignment2
    
    border1 = xlwt.Borders()
    border1.left = border1.THIN
    border1.right = border1.THIN
    border1.top = border1.THIN
    border1.bottom = border1.THIN
    style1.borders = border1
    style2.borders = border1
    style3.borders = border1
    style4.borders = border1
    style5.borders = border1
    style6.borders = border1
    style7.borders = border1
    style8.borders = border1
    style9.borders = border1
    style10.borders = border1
    style11.borders = border1
    
    font1 = xlwt.Font()
    font1.name = 'Atlanta'
    font1.colour_index = xlwt.Style.colour_map['red']
    font1.bold = True
    style1.font = font1
    
    font2 = xlwt.Font()
    font2.name = 'Atlanta'
    font2.colour_index = xlwt.Style.colour_map['white']
    font2.bold = True
    style2.font = font2
    
    font3 = xlwt.Font()
    font3.name = 'Atlanta'
    font3.bold = True
    style3.font = font3
    style4.font = font3
    style6.font = font3
    style8.font = font3
    style10.font = font3
    style11.font = font3
    
    font4 = xlwt.Font()
    font4.name = 'Atlanta'
    style5.font = font4
    style7.font = font4
    style9.font = font4
    
    pattern1 = xlwt.Pattern()
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['red']
    style2.pattern = pattern1
    
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = xlwt.Style.colour_map['tan']
    style4.pattern = pattern2
    
    pattern3 = xlwt.Pattern()
    pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern3.pattern_fore_colour = xlwt.Style.colour_map['lime']
    style6.pattern = pattern3
    
    pattern4 = xlwt.Pattern()
    pattern4.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern4.pattern_fore_colour = xlwt.Style.colour_map['ice_blue']
    style7.pattern = pattern4
    style8.pattern = pattern4
    style10.pattern = pattern4
    
    pattern5 = xlwt.Pattern()
    pattern5.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern5.pattern_fore_colour = xlwt.Style.colour_map['blue_gray']
    style11.pattern = pattern5
    
    current_date = datetime.now()
    current_month = current_date.month
    current_year = current_date.year
    start_month = 4
    total_month = 12
    
    ##################################################################################### SHEET 1 ##############################################################################################
    
    worksheet1.write_merge(3, 3, 0, 42, 'Scope 1 Emissions : Fuel consumption, Process emissions, Water Conservation', style1)
    worksheet1.write_merge(0, 0, 0, 5, 'Environment Sustainability Data Responsibility - Scope 1', style2)
    worksheet1.write(1, 0, 'Name of Champion', style3)
    worksheet1.write(2, 0, 'E-mail address', style3)
    worksheet1.write_merge(1, 1, 1, 5, None, style5)
    worksheet1.write_merge(2, 2, 1, 5, None, style5)
    worksheet1.write(4, 0, 'SCOPE 1 Emission', style4)
    worksheet1.write(5, 0, 'Source', style4)
    worksheet1.write_merge(4, 5, 1, 1, 'Fuel type', style4)
    worksheet1.write_merge(4, 5, 2, 2, 'Purpose', style4)
    worksheet1.write_merge(4, 5, 3, 3, 'Units', style4)
    worksheet1.write_merge(4, 4, 4, 16, 'Consumption, Used, Disposed, etc', style4)
    worksheet1.write_merge(4, 4, 17, 21, 'Consumption Past Years', style4)
    worksheet1.write_merge(4, 5, 22, 22, 'Name of GHG', style4)
    worksheet1.write_merge(4, 5, 23, 23, 'GHG Co2 equv', style4)
    worksheet1.write_merge(4, 5, 24, 24, 'Units', style4)
    worksheet1.write_merge(4, 4, 25, 37, 'Co2 equv calculation', style4)
    worksheet1.write_merge(4, 4, 38, 42, 'Co2 equv calculation Past Years', style4)
    worksheet1.write_merge(6, 10, 0, 0, 'Fuel combustion', style5)
    worksheet1.write_merge(11, 13, 0, 0, 'Process emission Released', style5)
    worksheet1.write(14, 0, 'Ground water consumption (In-House)', style5)
    worksheet1.write_merge(15, 16, 0, 0, 'Carbon sink', style5)
    worksheet1.set_panes_frozen(True)
    worksheet1.set_vert_split_pos(4)
    worksheet1.set_horz_split_pos(6)
    
    worksheet1.row(0).height_mismatch = True
    worksheet1.row(0).height = 500
    worksheet1.row(1).height_mismatch = True
    worksheet1.row(1).height = 500
    worksheet1.row(2).height_mismatch = True
    worksheet1.row(2).height = 500
    worksheet1.row(3).height_mismatch = True
    worksheet1.row(3).height = 530
    worksheet1.row(4).height_mismatch = True
    worksheet1.row(4).height = 600
    worksheet1.row(5).height_mismatch = True
    worksheet1.row(5).height = 600
    
    worksheet1.col(0).width = 6000
    worksheet1.col(1).width = 5000
    worksheet1.col(2).width = 4000
    worksheet1.col(3).width = 4000
    for col in range(4, 43):
        worksheet1.col(col).width = 3000
    
    row = 6
    array = ['Diesel', 'Petrol', 'LPG', 'CNG', 'PNG', 'Air conditioning', 'Process Chillers / HVAC', 'Fire extinguisher Co2', 'Fresh Water', 'Trees planted', 'Water Charging (Rain Water Harvesting)']
    for lists in array:
        worksheet1.write(row, 1, lists, style5)
        worksheet1.row(row).height_mismatch = True
        worksheet1.row(row).height = 640
        row += 1
        
    row = 6
    array = ['Fuel burnt', 'Fuel burnt', 'Fuel burnt', 'Fuel burnt', 'Fuel burnt', 'HCFC-22 (R-22)', 'HCFC-134a (R134a)', 'Co2 make up', 'Process & Domestic', 'General', 'General']
    for lists in array:
        worksheet1.write(row, 2, lists, style5)
        row += 1
        
    row = 6
    array = ['Litres', 'Litres', 'Litres', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'kL', 'Nos', 'kL']
    for lists in array:
        worksheet1.write(row, 3, lists, style5)
        row += 1
        
    row = 6
    array = ['Co2', 'Co2', 'Co2', 'Co2', 'Co2', 'HFC', 'HFC', 'Co2', 'Co2', 'Co2', 'Co2']
    for lists in array:
        worksheet1.write(row , 22, lists, style5)
        row += 1
        
    row = 6
    array = [2.6, 2.32, 2.19, 0.614, 0.056, 1810, 1300, 1, 0.59, -2, -0.59]
    for lists in array:
        worksheet1.write(row, 23, lists, style5)
        row += 1
            
    row = 6
    array = ['Kg Co2 equv / L', 'Kg Co2 equv / L', 'Kg Co2 equv / L', 'Kg Co2 equv / kg', 'Kg Co2 equv / kg', 'Kg Co2 equv / kG', 'Kg Co2 equv / kG', 'Kg Co2 equv / kG', 'Kg Co2 equv / kL', 'Kg Co2 equv / Tree', 'Kg Co2 equv / kL']
    for lists in array:
        worksheet1.write(row, 24, lists, style5)
        row += 1
    
    #This_year 
    col1 = 4
    col2 = 25
    if current_month >= start_month:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet1.write(5, col1, 'Cons. ' + month_names + ' ' + str(current_year)[2:], style4)
            worksheet1.write(5, col2, 'Emission - ' + month_names + ' ' + str(current_year)[2:], style4)
            
            worksheet2.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year)[2:], style4)
            worksheet2.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year)[2:], style4)
            
            worksheet3.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year)[2:], style4)
            worksheet3.write(5, col2 + 1, 'Emission - ' + month_names + " ' " + str(current_year)[2:], style4)
            
            worksheet4.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            
            worksheet5.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            col1 += 1
            col2 += 1
            start_month += 1
        x = current_year
        col1 = 13
        col2 = 34
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet1.write(5, col1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet1.write(5, col2, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet2.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet2.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet3.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet3.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet4.write(2, col1 - 2,  month_names + " ' " + str(current_year + 1)[2:], style4)
            
            worksheet5.write(2, col1 - 2,  month_names + " ' " + str(current_year + 1)[2:], style4)
            col1 += 1
            col2 += 1
            start_month += 1
        y = current_year + 1
        worksheet1.write(5, 16, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet1.write(5, 37, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet2.write(5, 17, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet2.write(5, 38, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet3.write(5, 17, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet3.write(5, 38, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet4.write(2, 14, 'YTM ' + str(x) + ' - ' +str(y), style4)
        
        worksheet5.write(2, 14, 'YTM ' + str(x) + ' - ' +str(y), style4)
    else:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet1.write(5, col1, 'Cons. ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            worksheet1.write(5, col2, 'Emission - ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            
            worksheet2.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            worksheet2.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            
            worksheet3.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            worksheet3.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year - 1)[2:], style4)
            
            worksheet4.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            
            worksheet5.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            col1 += 1
            col2 += 1
            start_month += 1
        x = current_year - 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet1.write(5, col1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet1.write(5, col2, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet2.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet2.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet3.write(5, col1 + 1, 'Cons. ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            worksheet3.write(5, col2 + 1, 'Emission - ' + month_names + ' ' + str(current_year + 1)[2:], style4)
            
            worksheet4.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            
            worksheet5.write(2, col1 - 2, month_names + " ' " + str(current_year)[2:], style4)
            col1 += 1
            col2 += 1
            start_month += 1
        y = current_year
        worksheet1.write(5, 16, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet1.write(5, 37, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet2.write(5, 17, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet2.write(5, 38, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet3.write(5, 17, 'Total ' + 'Cons. ' + str(x) + ' - ' +str(y), style4)
        worksheet3.write(5, 38, 'Total ' + 'Emission ' + str(x) + ' - ' +str(y), style4)
        
        worksheet4.write(2, 14, 'YTM ' + str(x) + ' - ' +str(y), style4)
        
        worksheet5.write(2, 14, 'YTM ' + str(x) + ' - ' +str(y), style4)
            
    totaldgcons = []; totaldgemis = []
    totalptcons = []; totalptemis = []
    totallpgcons = []; totallpgemis = []
    totalcngcons = []; totalcngemis = []
    totalpngcons = []; totalpngemis = []
    if current_month >= start_month:
        while start_month <= total_month:
            
            dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
            if dg_month_cons is None:
                dg_month_cons = 0
            worksheet1.write(6, start_month, dg_month_cons, style5)
                        
            totaldgcons.append(dg_month_cons)
            try:
                dg_month_emis = round(float(dg_month_cons) * 2.6)
            except:
                dg_month_emis = 0
            worksheet1.write(6, start_month + 21, dg_month_emis, style5)
            totaldgemis.append(dg_month_emis)

            pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
            if pt_month_cons is None:
                pt_month_cons = 0
            worksheet1.write(7, start_month, pt_month_cons, style5)
            totalptcons.append(pt_month_cons)
            try:
                pt_month_emis = round(float(pt_month_cons) * 2.32)
            except:
                pt_month_emis = 0
            worksheet1.write(7, start_month + 21, pt_month_emis, style5)
            totalptemis.append(pt_month_emis)
            
            lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
            if lpg_month_cons is None:
                lpg_month_cons = 0
            worksheet1.write(8, start_month, lpg_month_cons, style5)
            totallpgcons.append(lpg_month_cons)
            try:
                lpg_month_emis = round(float(lpg_month_cons) * 2.19)
            except:
                lpg_month_emis = 0
            worksheet1.write(8, start_month + 21, lpg_month_emis, style5)
            totallpgemis.append(lpg_month_emis)
            
            cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
            if cng_month_cons is None:
                cng_month_cons = 0
            worksheet1.write(9, start_month, cng_month_cons, style5)
            
            totalcngcons.append(cng_month_cons)
            try:
                cng_month_emis = round(float(cng_month_cons) * 0.614)
            except:
                cng_month_emis = 0
            worksheet1.write(9, start_month + 21, cng_month_emis, style5)
            
            totalcngemis.append(cng_month_emis)
            
            png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
            if png_month_cons is None:
                png_month_cons = 0
            worksheet1.write(10, start_month, png_month_cons, style5)
            totalpngcons.append(png_month_cons)
            try:
                png_month_emis = round(float(png_month_cons) * 0.056)
            except:
                png_month_emis = 0
            worksheet1.write(10, start_month + 21, png_month_emis, style5)
            totalpngemis.append(png_month_emis)
            
            start_month += 1
        start_month = 1
        while start_month < 4:
            
            dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year + 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
            if dg_month_cons is None:
                dg_month_cons = 0
            worksheet1.write(6, start_month + 12, dg_month_cons, style5)
            
            totaldgcons.append(dg_month_cons)
            try:
                dg_month_emis = round(float(dg_month_cons) * 2.6)
            except:
                dg_month_emis = 0
            worksheet1.write(6, start_month + 33, dg_month_emis, style5)
            totaldgemis.append(dg_month_emis)
            
            pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year + 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
            if pt_month_cons is None:
                pt_month_cons = 0
            worksheet1.write(7, start_month + 12, pt_month_cons, style5)
            totalptcons.append(pt_month_cons)
            try:
                pt_month_emis = round(float(pt_month_cons) * 2.32)
            except:
                pt_month_emis = 0
            worksheet1.write(7, start_month + 33, pt_month_emis, style5)
            totalptemis.append(pt_month_emis)
            
            lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year + 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
            if lpg_month_cons is None:
                lpg_month_cons = 0
            worksheet1.write(8, start_month + 12, lpg_month_cons, style5)
            totallpgcons.append(lpg_month_cons)
            try:
                lpg_month_emis = round(float(lpg_month_cons) * 2.19)
            except:
                lpg_month_emis = 0
            worksheet1.write(8, start_month + 33, lpg_month_emis, style5)
            totallpgemis.append(lpg_month_emis)
            
            cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year + 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
            if cng_month_cons is None:
                cng_month_cons = 0
            worksheet1.write(9, start_month + 12, cng_month_cons, style5)
            
            totalcngcons.append(cng_month_cons)
            try:
                cng_month_emis = round(float(cng_month_cons) * 0.614)
            except:
                cng_month_emis = 0
            worksheet1.write(9, start_month + 33, cng_month_emis, style5)
            
            totalcngemis.append(cng_month_emis)
            
            png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year + 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
            if png_month_cons is None:
                png_month_cons = 0
            worksheet1.write(10, start_month + 12, png_month_cons, style5)
            totalpngcons.append(png_month_cons)
            try:
                png_month_emis = round(float(png_month_cons) * 0.056)
            except:
                png_month_emis = 0
            worksheet1.write(10, start_month + 33, png_month_emis, style5)
            totalpngemis.append(png_month_emis)
            
            start_month += 1
    else:
        while start_month <= total_month:
            
            dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year - 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
            if dg_month_cons is None:
                dg_month_cons = 0
            worksheet1.write(6, start_month, dg_month_cons, style5)
            
            totaldgcons.append(dg_month_cons)
            try:
                dg_month_emis = round(float(dg_month_cons) * 2.6)
            except:
                dg_month_emis = 0
            worksheet1.write(6, start_month + 21, dg_month_emis, style5)
            totaldgemis.append(dg_month_emis)
            
            pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year - 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
            if pt_month_cons is None:
                pt_month_cons = 0
            worksheet1.write(7, start_month, pt_month_cons, style5)
            totalptcons.append(pt_month_cons)
            try:
                pt_month_emis = round(float(pt_month_cons) * 2.32)
            except:
                pt_month_emis = 0
            worksheet1.write(7, start_month + 21, pt_month_emis, style5)
            totalptemis.append(pt_month_emis)
            
            lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year - 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
            if lpg_month_cons is None:
                lpg_month_cons= 0
            worksheet1.write(8, start_month, lpg_month_cons, style5)
            totallpgcons.append(lpg_month_cons)
            try:
                lpg_month_emis = round(float(lpg_month_cons) * 2.19)
            except:
                lpg_month_emis = 0
            worksheet1.write(8, start_month + 21, lpg_month_emis, style5)
            totallpgemis.append(lpg_month_emis)
            
            cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year - 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
            if cng_month_cons is None:
                cng_month_cons = 0
            worksheet1.write(9, start_month, cng_month_cons, style5)
            
            totalcngcons.append(cng_month_cons)
            try:
                cng_month_emis = round(float(cng_month_cons) * 0.614)
            except:
                cng_month_emis = 0
            worksheet1.write(9, start_month + 21, cng_month_emis, style5)
            
            totalcngemis.append(cng_month_emis)
            
            png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year - 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
            if png_month_cons is None:
                png_month_cons = 0
            worksheet1.write(10, start_month, png_month_cons, style5)
            totalpngcons.append(png_month_cons)
            try:
                png_month_emis = round(float(png_month_cons) * 0.056)
            except:
                png_month_emis = 0
            worksheet1.write(10, start_month + 21, png_month_emis, style5)
            totalpngemis.append(png_month_emis)
            
            start_month += 1
        start_month = 1
        while start_month < 4:
            
            dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
            if dg_month_cons is None:
                dg_month_cons = 0
            worksheet1.write(6, start_month + 12, dg_month_cons, style5)
            
            totaldgcons.append(dg_month_cons)
            try:
                dg_month_emis = round(float(dg_month_cons) * 2.6)
            except:
                dg_month_emis = 0
            worksheet1.write(6, start_month + 33, dg_month_emis, style5)
            totaldgemis.append(dg_month_emis)
            
            pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
            if pt_month_cons is None:
                pt_month_cons = 0
            worksheet1.write(7, start_month + 12, pt_month_cons, style5)
            totalptcons.append(pt_month_cons)
            try:
                pt_month_emis = round(float(pt_month_cons) * 2.32)
            except:
                pt_month_emis = 0
            worksheet1.write(7, start_month + 33, pt_month_emis, style5)
            totalptemis.append(pt_month_emis)
            
            lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
            if lpg_month_cons is None:
                lpg_month_cons= 0
            worksheet1.write(8, start_month + 12, lpg_month_cons, style5)
            totallpgcons.append(lpg_month_cons)
            try:
                lpg_month_emis = round(float(lpg_month_cons) * 2.19)
            except:
                lpg_month_emis = 0
            worksheet1.write(8, start_month + 33, lpg_month_emis, style5)
            totallpgemis.append(lpg_month_emis)
            
            cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
            if cng_month_cons is None:
                cng_month_cons = 0
            worksheet1.write(9, start_month + 12, cng_month_cons, style5)
            
            totalcngcons.append(cng_month_cons)
            try:
                cng_month_emis = round(float(cng_month_cons) * 0.614)
            except:
                cng_month_emis = 0
            worksheet1.write(9, start_month + 33, cng_month_emis, style5)
            totalcngemis.append(cng_month_emis)
            
            png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
            if png_month_cons is None:
                png_month_cons = 0
            worksheet1.write(10, start_month + 12, png_month_cons, style5)
            totalpngcons.append(png_month_cons)
            try:
                png_month_emis = round(float(png_month_cons) * 0.056)
            except:
                png_month_emis = 0
            worksheet1.write(10, start_month + 33, png_month_emis, style5)
            
            totalpngemis.append(png_month_emis)
            
            start_month += 1

    worksheet1.write(6, 16, round(sum(totaldgcons)), style5)
    worksheet1.write(6, 37, round(sum(totaldgemis)), style5)
    worksheet1.write(7, 16, round(sum(totalptcons)), style5)
    worksheet1.write(7, 37, round(sum(totalptemis)), style5)
    worksheet1.write(8, 16, round(sum(totallpgcons)), style5)
    worksheet1.write(8, 37, round(sum(totallpgemis)), style5)
    worksheet1.write(9, 16, round(sum(totalcngcons)), style5)
    worksheet1.write(9, 37, round(sum(totalcngemis)), style5)
    worksheet1.write(10, 16, round(sum(totalpngcons)), style5)
    worksheet1.write(10, 37, round(sum(totalpngemis)), style5)
    
    #Past_five_years
    
    col1 = 17
    col2 = 38
    for year in range(1, 6):
        if current_month >= start_month:
            while start_month <= total_month:
                start_month += 1
            x = current_year - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = (current_year + 1) - year
            worksheet1.write(5, col1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet1.write(5, col2, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet2.write(5, col1 + 1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet2.write(5, col2 + 1, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet3.write(5, col1 + 1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet3.write(5, col2 + 1, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet4.write(2, col1 - 2, str(x) + ' - ' + str(y), style4)
            
            worksheet5.write(2, col1 - 2, str(x) + ' - ' + str(y), style4)
            col1 += 1
            col2 += 1
        else:
            while start_month <= total_month:
                start_month += 1
            x = (current_year - 1) - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = current_year - year
            worksheet1.write(5, col1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet1.write(5, col2, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet2.write(5, col1 + 1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet2.write(5, col2 + 1, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet3.write(5, col1 + 1, 'Total ' + 'Cons. ' + str(x) + ' - ' + str(y), style4)
            worksheet3.write(5, col2 + 1, 'Total ' + 'Emission ' + str(x) + ' - ' + str(y), style4)
            
            worksheet4.write(2, col1 - 2, str(x) + ' - ' + str(y), style4)
            
            worksheet5.write(2, col1 - 2, str(x) + ' - ' + str(y), style4)
            col1 += 1
            col2 += 1
    
    col1 = 17
    col2 = 38
    for year in range(1, 6):
        totaldgcons = []; totaldgemis = []
        totalptcons = []; totalptemis = []
        totallpgcons = []; totallpgemis = []
        totalcngcons = []; totalcngemis = []
        totalpngcons = []; totalpngemis = []
        if current_month >= start_month:
            while start_month <= total_month:
                
                dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                if dg_month_cons is None:
                    dg_month_cons = 0
                totaldgcons.append(dg_month_cons)
                try:
                    dg_month_emis = round(float(dg_month_cons) * 2.6)
                except:
                    dg_month_emis = 0
                totaldgemis.append(dg_month_emis)

                pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                if pt_month_cons is None:
                    pt_month_cons = 0
                totalptcons.append(pt_month_cons)
                try:
                    pt_month_emis = round(float(pt_month_cons) * 2.32)
                except:
                    pt_month_emis = 0
                totalptemis.append(pt_month_emis)
                
                lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                if lpg_month_cons is None:
                    lpg_month_cons = 0
                totallpgcons.append(lpg_month_cons)
                try:
                    lpg_month_emis = round(float(lpg_month_cons) * 2.19)
                except:
                    lpg_month_emis = 0
                totallpgemis.append(lpg_month_emis)
                
                cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                if cng_month_cons is None:
                    cng_month_cons = 0
                totalcngcons.append(cng_month_cons)
                try:
                    cng_month_emis = round(float(cng_month_cons) * 0.614)
                except:
                    cng_month_emis = 0
                totalcngemis.append(cng_month_emis)
                
                png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                if png_month_cons is None:
                    png_month_cons = 0
                totalpngcons.append(png_month_cons)
                try:
                    png_month_emis = round(float(png_month_cons) * 0.056)
                except:
                    png_month_emis = 0
                totalpngemis.append(png_month_emis)
                
                start_month += 1
            start_month = 1
            while start_month < 4:
                
                dg_month_cons = Cost_DG.objects.filter(dg_date__year = (current_year + 1) - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                if dg_month_cons is None:
                    dg_month_cons = 0
                totaldgcons.append(dg_month_cons)
                try:
                    dg_month_emis = round(float(dg_month_cons) * 2.6)
                except:
                    dg_month_emis = 0
                totaldgemis.append(dg_month_emis)
                
                pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = (current_year + 1) - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                if pt_month_cons is None:
                    pt_month_cons = 0
                totalptcons.append(pt_month_cons)
                try:
                    pt_month_emis = round(float(pt_month_cons) * 2.32)
                except:
                    pt_month_emis = 0
                totalptemis.append(pt_month_emis)
                
                lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = (current_year + 1) - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                if lpg_month_cons is None:
                    lpg_month_cons = 0
                totallpgcons.append(lpg_month_cons)
                try:
                    lpg_month_emis = round(float(lpg_month_cons) * 2.19)
                except:
                    lpg_month_emis = 0
                totallpgemis.append(lpg_month_emis)
                
                cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = (current_year + 1) - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                if cng_month_cons is None:
                    cng_month_cons = 0
                totalcngcons.append(cng_month_cons)
                try:
                    cng_month_emis = round(float(cng_month_cons) * 0.614)
                except:
                    cng_month_emis = 0
                totalcngemis.append(cng_month_emis)
                
                png_month_cons = Cost_PNG.objects.filter(PNG_date__year = (current_year + 1) - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                if png_month_cons is None:
                    png_month_cons = 0
                totalpngcons.append(png_month_cons)
                try:
                    png_month_emis = round(float(png_month_cons) * 0.056)
                except:
                    png_month_emis = 0
                totalpngemis.append(png_month_emis)
                
                start_month += 1
        else:
            while start_month <= total_month:
                
                dg_month_cons = Cost_DG.objects.filter(dg_date__year = (current_year - 1) - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                if dg_month_cons is None:
                    dg_month_cons = 0
                totaldgcons.append(dg_month_cons)
                try:
                    dg_month_emis = round(float(dg_month_cons) * 2.6)
                except:
                    dg_month_emis = 0
                totaldgemis.append(dg_month_emis)
                
                pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = (current_year - 1) - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                if pt_month_cons is None:
                    pt_month_cons = 0
                totalptcons.append(pt_month_cons)
                try:
                    pt_month_emis = round(float(pt_month_cons) * 2.32)
                except:
                    pt_month_emis = 0
                totalptemis.append(pt_month_emis)
                
                lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = (current_year - 1) - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                if lpg_month_cons is None:
                    lpg_month_cons= 0
                totallpgcons.append(lpg_month_cons)
                try:
                    lpg_month_emis = round(float(lpg_month_cons) * 2.19)
                except:
                    lpg_month_emis = 0
                totallpgemis.append(lpg_month_emis)
                
                cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = (current_year - 1) - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                if cng_month_cons is None:
                    cng_month_cons = 0
                totalcngcons.append(cng_month_cons)
                try:
                    cng_month_emis = round(float(cng_month_cons) * 0.614)
                except:
                    cng_month_emis = 0
                totalcngemis.append(cng_month_emis)
                
                png_month_cons = Cost_PNG.objects.filter(PNG_date__year = (current_year - 1) - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                if png_month_cons is None:
                    png_month_cons = 0
                totalpngcons.append(png_month_cons)
                try:
                    png_month_emis = round(float(png_month_cons) * 0.056)
                except:
                    png_month_emis = 0
                totalpngemis.append(png_month_emis)
                
                start_month += 1
            start_month = 1
            while start_month < 4:
                
                dg_month_cons = Cost_DG.objects.filter(dg_date__year = current_year - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                if dg_month_cons is None:
                    dg_month_cons = 0
                totaldgcons.append(dg_month_cons)
                try:
                    dg_month_emis = round(float(dg_month_cons) * 2.6)
                except:
                    dg_month_emis = 0
                totaldgemis.append(dg_month_emis)
                
                pt_month_cons = Cost_Petrol.objects.filter(pt_date__year = current_year - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                if pt_month_cons is None:
                    pt_month_cons = 0
                totalptcons.append(pt_month_cons)
                try:
                    pt_month_emis = round(float(pt_month_cons) * 2.32)
                except:
                    pt_month_emis = 0
                totalptemis.append(pt_month_emis)
                
                lpg_month_cons = Cost_LPG.objects.filter(LPG_date__year = current_year - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                if lpg_month_cons is None:
                    lpg_month_cons= 0
                totallpgcons.append(lpg_month_cons)
                try:
                    lpg_month_emis = round(float(lpg_month_cons) * 2.19)
                except:
                    lpg_month_emis = 0
                totallpgemis.append(lpg_month_emis)
                
                cng_month_cons = Cost_CNG.objects.filter(CNG_date__year = current_year - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                if cng_month_cons is None:
                    cng_month_cons = 0
                totalcngcons.append(cng_month_cons)
                try:
                    cng_month_emis = round(float(cng_month_cons) * 0.614)
                except:
                    cng_month_emis = 0
                totalcngemis.append(cng_month_emis)
                
                png_month_cons = Cost_PNG.objects.filter(PNG_date__year = current_year - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                if png_month_cons is None:
                    png_month_cons = 0
                totalpngcons.append(png_month_cons)
                try:
                    png_month_emis = round(float(png_month_cons) * 0.056)
                except:
                    png_month_emis = 0
                totalpngemis.append(png_month_emis)
                
                start_month += 1

        worksheet1.write(6, col1, round(sum(totaldgcons)), style5)
        worksheet1.write(6, col2, round(sum(totaldgemis)), style5)
        worksheet1.write(7, col1, round(sum(totalptcons)), style5)
        worksheet1.write(7, col2, round(sum(totalptemis)), style5)
        worksheet1.write(8, col1, round(sum(totallpgcons)), style5)
        worksheet1.write(8, col2, round(sum(totallpgemis)), style5)
        worksheet1.write(9, col1, round(sum(totalcngcons)), style5)
        worksheet1.write(9, col2, round(sum(totalcngemis)), style5)
        worksheet1.write(10, col1, round(sum(totalpngcons)), style5)
        worksheet1.write(10, col2, round(sum(totalpngemis)), style5)
        
        col1 += 1
        col2 += 1
    
    for row in range(11, 14):
        for col in range(4, 22):
            worksheet1.write(row, col, None, style5)
            
    for row in range(11, 14):
        for col in range(25, 43):
            worksheet1.write(row, col, None, style5)
    
    for row in range(14, 17):
        for col in range(4, 22):
            worksheet1.write(row, col, None, style5)
            
    for row in range(14, 17):
        for col in range(25, 43):
            worksheet1.write(row, col, None, style5)
    
    ##################################################################################### SHEET 2 ###############################################################################################
    
    worksheet2.write_merge(3, 3, 0, 43, 'Scope 2 Emissions : Purchased electricity, stream, heat & cooling', style1)
    worksheet2.write_merge(0, 0, 0, 5, 'Environment Sustainability Data Responsibility - Scope 2', style2)
    worksheet2.write(1, 0, 'Name of Champion', style3)
    worksheet2.write(2, 0, 'Department & Position', style3)
    worksheet2.write_merge(1, 1, 1, 5, None, style5)
    worksheet2.write_merge(2, 2, 1, 5, None, style5)
    worksheet2.write_merge(4, 4, 0, 4, 'SCOPE 2 Emission', style4)
    worksheet2.write_merge(4, 4, 5, 17, 'Consumption, Used, Disposed, etc', style4)
    worksheet2.write_merge(4, 4, 18, 22, 'Consumption Past Years', style4)
    worksheet2.write_merge(4, 4, 26, 38, 'Co2 equv calculation', style4)
    worksheet2.write_merge(4, 4, 39, 43, 'Co2 equv calculation Past Years', style4)
    worksheet2.write_merge(6, 11, 0, 0, 'I', style5)
    worksheet2.write(12, 0,  'II', style5)
    worksheet2.write_merge(4, 5, 23, 23, 'Name of GHG', style4)
    worksheet2.write_merge(4, 5, 24, 24, 'GHG Co2 equv', style4)
    worksheet2.write_merge(4, 5, 25, 25, 'Units', style4)
    worksheet2.write(5, 0, 'Chapter', style4)
    worksheet2.write(5, 1, 'Source', style4)
    worksheet2.write(5, 2, 'Type', style4)
    worksheet2.write(5, 3, 'Purpose', style4)
    worksheet2.write(5, 4, 'Units', style4)
    worksheet2.set_panes_frozen(True)
    worksheet2.set_vert_split_pos(5)
    worksheet2.set_horz_split_pos(6)
    
    worksheet2.row(0).height_mismatch = True
    worksheet2.row(0).height = 500
    worksheet2.row(1).height_mismatch = True
    worksheet2.row(1).height = 500
    worksheet2.row(2).height_mismatch = True
    worksheet2.row(2).height = 500
    worksheet2.row(3).height_mismatch = True
    worksheet2.row(3).height = 530
    worksheet2.row(4).height_mismatch = True
    worksheet2.row(4).height = 600
    worksheet2.row(5).height_mismatch = True
    worksheet2.row(5).height = 650
    worksheet2.row(13).height_mismatch = True
    worksheet2.row(13).height = 500
    
    worksheet2.col(0).width = 6000
    worksheet2.col(1).width = 8500
    worksheet2.col(2).width = 4000
    worksheet2.col(3).width = 4000
    worksheet2.col(4).width = 2500
    for col in range(5, 44):
        worksheet2.col(col).width = 3000
    
    array = ['Purchased electricity from Electricity Authorities (non renewable)', 'Purchased electricity from Third Party (Nuclear - Non-Renewable)', 'Purchased electricity from Third Party (CNG - Non-Renewable)', 'Purchased electricity from Third Party (Biogas - Renewable)', 'Purchased electricity from Second Party (Solar - Renewable)', 'Purchased electricity from Third Party (Wind - Renewable)', 'Water Consumption from Out-source']
    row = 6
    for lists in array:
        worksheet2.write(row, 1, lists, style5)
        worksheet2.row(row).height_mismatch = True
        worksheet2.row(row).height = 900
        row += 1
        
    row = 6
    array = ['Electricity', 'Electricity', 'Electricity', 'Electricity', 'Electricity', 'Electricity', 'Water']
    for lists in array:
        worksheet2.write(row, 2, lists, style5)
        row += 1
        
    row = 6
    array = ['Production', 'Production', 'Production', 'Production', 'Production', 'Production', 'Process / Domestic']
    for lists in array:
        worksheet2.write(row, 3, lists, style5)
        row += 1
        
    row = 6
    array = ['kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kL']
    for lists in array:
        worksheet2.write(row, 4, lists, style5)
        row += 1
        
    row = 6
    array = ['Co2', 'Co2', 'Co2', 'Methane', 'Co2', 'Co2', 'Co2']
    for lists in array:
        worksheet2.write(row, 23, lists, style5)
        row += 1
        
    row = 6
    array = [0.7132, 0.7132, 0.214, 1.13, 0.041, 0.018, 0.59]
    for lists in array:
        worksheet2.write(row, 24, lists, style5)
        row += 1
        
    row = 6
    array = ['Kg Co2 equv / kWh', 'Kg Co2 equv / kWh', 'Kg Co2 equv / kWh', 'Kg Co2 equv / kWh', 'Kg Co2 equv / kWh', 'Kg Co2 equv / kWh', 'Kg Co2 equv / kL']
    for lists in array:
        worksheet2.write(row, 25, lists, style5)
        row += 1
        
    sourcename = list(AddSource.objects.filter(asplantname = plantname).values('assourcename'))
    transcons = []; transemis = []
    windcons = []; windemis = []
    solarcons = []; solaremis = []
    for source in sourcename:
        if source['assourcename'] == 'Transformer1':
            if current_month >= start_month:
                while start_month <= total_month:
                    transformer = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if transformer is None:
                        transformer = 0
                    wind = Cost_Wind.objects.filter(Wind_year = current_year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                    if wind is None:
                        wind = 0
                    actual_wind = transformer * wind / 100
                    actual_trans = transformer - actual_wind
                    actualtrans_emis = float(actual_trans) * 0.7132
                    actualwind_emis = float(actual_wind) * 0.018
                    transcons.append(round(actual_trans))
                    transemis.append(round(actualtrans_emis))
                    windcons.append(round(actual_wind))
                    windemis.append(round(actualwind_emis))
                    worksheet2.write(6, start_month + 1, round(actual_trans), style5)
                    worksheet2.write(11, start_month + 1, round(actual_wind), style5)
                    worksheet2.write(6, start_month + 22, round(actualtrans_emis), style5)
                    worksheet2.write(11, start_month + 22, round(actualwind_emis), style5)
                    
                    worksheet5.write(10, start_month - 2, round(transformer), style11)
                    worksheet5.write(3, start_month - 2, round(actual_trans), style5)
                    worksheet5.write(9, start_month - 2, round(actual_wind), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    transformer = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if transformer is None:
                        transformer = 0
                    wind = Cost_Wind.objects.filter(Wind_year = current_year + 1, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                    if wind is None:
                        wind = 0
                    actual_wind = transformer * wind / 100
                    actual_trans = transformer - actual_wind
                    actualtrans_emis = float(actual_trans) * 0.7132
                    actualwind_emis = float(actual_wind) * 0.018
                    transcons.append(round(actual_trans))
                    transemis.append(round(actualtrans_emis))
                    windcons.append(round(actual_wind))
                    windemis.append(round(actualwind_emis))
                    worksheet2.write(6, start_month + 13, round(actual_trans), style5)
                    worksheet2.write(11, start_month + 13, round(actual_wind), style5)
                    worksheet2.write(6, start_month + 34, round(actualtrans_emis), style5)
                    worksheet2.write(11, start_month + 34, round(actualwind_emis),style5)
                    
                    worksheet5.write(10, start_month + 10, round(transformer), style11)
                    worksheet5.write(3, start_month + 10, round(actual_trans), style5)
                    worksheet5.write(9, start_month + 10, round(actual_wind), style5)
                    
                    start_month += 1
            else:
                while start_month <= total_month:
                    transformer = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if transformer is None:
                        transformer = 0
                    wind = Cost_Wind.objects.filter(Wind_year = current_year - 1, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                    if wind is None:
                        wind = 0
                    actual_wind = transformer * wind / 100
                    actual_trans = transformer - actual_wind
                    actualtrans_emis = float(actual_trans) * 0.7132
                    actualwind_emis = float(actual_wind) * 0.018
                    transcons.append(round(actual_trans))
                    transemis.append(round(actualtrans_emis))
                    windcons.append(round(actual_wind))
                    windemis.append(round(actualwind_emis))
                    worksheet2.write(6, start_month + 1, round(actual_trans), style5)
                    worksheet2.write(11, start_month + 1, round(actual_wind), style5)
                    worksheet2.write(6, start_month + 22, round(actualtrans_emis), style5)
                    worksheet2.write(11, start_month + 22, round(actualwind_emis), style5)
                    
                    worksheet5.write(10, start_month - 2, round(transformer), style11)
                    worksheet5.write(3, start_month - 2, round(actual_trans), style5)
                    worksheet5.write(9, start_month - 2, round(actual_wind), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    transformer = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if transformer is None:
                        transformer = 0
                    wind = Cost_Wind.objects.filter(Wind_year = current_year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                    if wind is None:
                        wind = 0
                    actual_wind = transformer * wind / 100
                    actual_trans = transformer - actual_wind
                    actualtrans_emis = float(actual_trans) * 0.7132
                    actualwind_emis = float(actual_wind) * 0.018
                    transcons.append(round(actual_trans))
                    transemis.append(round(actualtrans_emis))
                    windcons.append(round(actual_wind))
                    windemis.append(round(actualwind_emis))
                    worksheet2.write(6, start_month + 13, round(actual_trans), style5)
                    worksheet2.write(11, start_month + 13, round(actual_wind), style5)
                    worksheet2.write(6, start_month + 34, round(actualtrans_emis), style5)
                    worksheet2.write(11, start_month + 34, round(actualwind_emis),style5)
                    
                    worksheet5.write(10, start_month + 10, round(transformer), style11)
                    worksheet5.write(3, start_month + 10, round(actual_trans), style5)
                    worksheet5.write(9, start_month + 10, round(actual_wind), style5)
                    
                    start_month += 1
                    
        if source['assourcename'] == 'Solar Energy':
            if current_month >= start_month:
                while start_month <= total_month:
                    solar = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if solar is None:
                        solar = 0
                    solarcons.append(round(solar))
                    solaremis.append(round(float(solar) * 0.041))
                    worksheet2.write(10, start_month + 1, round(solar), style5)
                    worksheet2.write(10, start_month + 22, round(float(solar) * 0.041), style5)
                    
                    worksheet5.write(8, start_month - 2, round(solar), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    solar = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if solar is None:
                        solar = 0
                    solarcons.append(round(solar))
                    solaremis.append(round(float(solar) * 0.041))
                    worksheet2.write(10, start_month + 13, round(solar), style5)
                    worksheet2.write(10, start_month + 34, round(float(solar) * 0.041), style5)
                    
                    worksheet5.write(8, start_month + 10, round(solar), style5)
                    
                    start_month += 1
            else:
                while start_month <= total_month:
                    solar = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if solar is None:
                        solar = 0
                    solarcons.append(round(solar))
                    solaremis.append(round(float(solar) * 0.041))
                    worksheet2.write(10, start_month + 1, round(solar), style5)
                    worksheet2.write(10, start_month + 22, round(float(solar) * 0.041), style5)
                    
                    worksheet5.write(8, start_month - 2, round(solar), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    solar = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if solar is None:
                        solar = 0
                    solarcons.append(round(solar))  
                    solaremis.append(round(float(solar) * 0.041))
                    worksheet2.write(10, start_month + 13, round(solar), style5)
                    worksheet2.write(10, start_month + 34, round(float(solar) * 0.041), style5)
                    
                    worksheet5.write(8, start_month + 10, round(solar), style5)
                    
                    start_month += 1
                    
    worksheet2.write(6, 17, sum(transcons), style5)
    worksheet2.write(6, 38, sum(transemis), style5)
    worksheet2.write(11, 17, sum(windcons), style5)
    worksheet2.write(11, 38, sum(windemis), style5)
    worksheet2.write(10, 17, sum(solarcons), style5)
    worksheet2.write(10, 38, sum(solaremis), style5)
    
    worksheet5.write(3, 14, sum(transcons), style5)
    worksheet5.write(8, 14, sum(solarcons), style5)
    worksheet5.write(9, 14, sum(windcons), style5)
    worksheet5.write(10, 14, round(sum(transcons + windcons)), style11)
    
    col1 = 18
    col2 = 39
    for year in range(1, 6):
        sourcename = list(AddSource.objects.filter(asplantname = plantname).values('assourcename'))
        transcons = []; transemis = []
        windcons = []; windemis = []
        solarcons = []; solaremis = []
        for source in sourcename:
            if source['assourcename'] == 'Transformer1':
                if current_month >= start_month:
                    while start_month <= total_month:
                        transformer = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if transformer is None:
                            transformer = 0
                        wind = Cost_Wind.objects.filter(Wind_year = current_year - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                        if wind is None:
                            wind = 0
                        actual_wind = transformer * wind / 100
                        actual_trans = transformer - actual_wind
                        actualtrans_emis = float(actual_trans) * 0.7132
                        actualwind_emis = float(actual_wind) * 0.018
                        transcons.append(round(actual_trans))
                        transemis.append(round(actualtrans_emis))
                        windcons.append(round(actual_wind))
                        windemis.append(round(actualwind_emis))
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        transformer = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if transformer is None:
                            transformer = 0
                        wind = Cost_Wind.objects.filter(Wind_year = (current_year + 1) - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                        if wind is None:
                            wind = 0
                        actual_wind = transformer * wind / 100
                        actual_trans = transformer - actual_wind
                        actualtrans_emis = float(actual_trans) * 0.7132
                        actualwind_emis = float(actual_wind) * 0.018
                        transcons.append(round(actual_trans))
                        transemis.append(round(actualtrans_emis))
                        windcons.append(round(actual_wind))
                        windemis.append(round(actualwind_emis))
                        start_month += 1
                else:
                    while start_month <= total_month:
                        transformer = Masterdatatable.objects.filter(mtdate__year = (current_year - 1) - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if transformer is None:
                            transformer = 0
                        wind = Cost_Wind.objects.filter(Wind_year = (current_year - 1) - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                        if wind is None:
                            wind = 0
                        actual_wind = transformer * wind / 100
                        actual_trans = transformer - actual_wind
                        actualtrans_emis = float(actual_trans) * 0.7132
                        actualwind_emis = float(actual_wind) * 0.018
                        transcons.append(round(actual_trans))
                        transemis.append(round(actualtrans_emis))
                        windcons.append(round(actual_wind))
                        windemis.append(round(actualwind_emis))
                        
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        transformer = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if transformer is None:
                            transformer = 0
                        wind = Cost_Wind.objects.filter(Wind_year = current_year - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                        if wind is None:
                            wind = 0
                        actual_wind = transformer * wind / 100
                        actual_trans = transformer - actual_wind
                        actualtrans_emis = float(actual_trans) * 0.7132
                        actualwind_emis = float(actual_wind) * 0.018
                        transcons.append(round(actual_trans))
                        transemis.append(round(actualtrans_emis))
                        windcons.append(round(actual_wind))
                        windemis.append(round(actualwind_emis))
                        start_month += 1
                        
            if source['assourcename'] == 'Solar Energy':
                if current_month >= start_month:
                    while start_month <= total_month:
                        solar = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if solar is None:
                            solar = 0
                        solarcons.append(round(solar))
                        solaremis.append(round(float(solar) * 0.041))
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        solar = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if solar is None:
                            solar = 0
                        solarcons.append(round(solar))
                        solaremis.append(round(float(solar) * 0.041))
                        start_month += 1
                else:
                    while start_month <= total_month:
                        solar = Masterdatatable.objects.filter(mtdate__year = (current_year - 1) - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if solar is None:
                            solar = 0
                        solarcons.append(round(solar))
                        solaremis.append(round(float(solar) * 0.041))
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        solar = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                        if solar is None:
                            solar = 0
                        solarcons.append(round(solar))
                        solaremis.append(round(float(solar) * 0.041))
                        start_month += 1
                        
        worksheet2.write(6, col1, sum(transcons), style5)
        worksheet2.write(6, col2, sum(transemis), style5)
        worksheet2.write(11, col1, sum(windcons), style5)
        worksheet2.write(11, col2, sum(windemis), style5)
        worksheet2.write(10, col1, sum(solarcons), style5)
        worksheet2.write(10, col2, sum(solaremis), style5)
        
        worksheet5.write(3, col1 - 3, sum(transcons), style5)
        worksheet5.write(8, col1 - 3, sum(solarcons), style5)
        worksheet5.write(9, col1 - 3, sum(windcons), style5)
        worksheet5.write(10, col1 - 3, round(sum(transcons + windcons)), style11)
        col1 += 1
        col2 += 1
        
    for row in range(7, 10):
        for col in range(5, 23):
            worksheet2.write(row, col, None, style5)
        
    for row in range(7, 10):
        for col in range(26, 44):
            worksheet2.write(row, col, None, style5)
    
    for col in range(5, 23):
        worksheet2.write(12, col, None, style5)
        
    for col in range(26, 44):
        worksheet2.write(12, col, None, style5)
            
    ##################################################################################### SHEET 3 ###############################################################################################
    
    worksheet3.write_merge(3, 3, 0, 43, 'Scope 3 Emissions : Transportation and distribution (up and downstream), Purchased goods and services, business travel, employee commuting, investments, leased assets and Vendors', style1)
    worksheet3.write_merge(0, 0, 0, 5, 'Environment Sustainability Data Responsibility - Scope 3', style2)
    worksheet3.write(1, 0, 'Name of Champion', style3)
    worksheet3.write(2, 0, 'Department & Position', style3)
    worksheet3.write_merge(1, 1, 1, 5, None, style5)
    worksheet3.write_merge(2, 2, 1, 5, None, style5)
    worksheet3.write_merge(4, 4, 0, 4, 'SCOPE 3 Emission', style4)
    worksheet3.write_merge(4, 4, 5, 17, 'Consumption, Used, Disposed, etc', style4)
    worksheet3.write_merge(4, 4, 18, 22, 'Consumption Past Years', style4)
    worksheet3.write_merge(4, 4, 26, 38, 'Co2 equv calculation', style4)
    worksheet3.write_merge(4, 4, 39, 43, 'Co2 equv calculation Past Years', style4)
    worksheet3.write_merge(6, 9, 0, 0, 'I', style5)
    worksheet3.write_merge(6, 9, 1, 1, 'Raw material Consumption', style5)
    worksheet3.write_merge(10, 17, 0, 0, 'II', style5)
    worksheet3.write_merge(10, 17, 1, 1, 'Consumables', style5)
    worksheet3.write_merge(18, 26, 0, 0, 'III', style5)
    worksheet3.write_merge(18, 26, 1, 1, 'Non-Hazardous waste disposal', style5)
    worksheet3.write_merge(27, 34, 0, 0, 'IV', style5)
    worksheet3.write_merge(27, 34, 1, 1, 'Hazardous waste disposal', style5)
    worksheet3.write_merge(35, 37, 0, 0, 'V', style5)
    worksheet3.write_merge(35, 37, 1, 1, 'Employee Commuting (common transport)', style5)
    worksheet3.write_merge(38, 40, 0, 0, 'VI', style5)
    worksheet3.write_merge(38, 40, 1, 1, 'Employee Commuting (company provided vehicles for employees)', style5)
    worksheet3.write_merge(41, 44, 0, 0, 'VII', style5)
    worksheet3.write_merge(41, 44, 1, 1, 'Employee Commuting (business Travel)', style5)
    worksheet3.write_merge(45, 47, 0, 0, 'VIII', style5)
    worksheet3.write_merge(45, 47, 1, 1, 'Finished goods (dispatch to customer end)', style5)
    worksheet3.set_panes_frozen(True)
    worksheet3.set_vert_split_pos(5)
    worksheet3.set_horz_split_pos(6)
    
    worksheet3.row(0).height_mismatch = True
    worksheet3.row(0).height = 500
    worksheet3.row(1).height_mismatch = True
    worksheet3.row(1).height = 500
    worksheet3.row(2).height_mismatch = True
    worksheet3.row(2).height = 500
    worksheet3.row(3).height_mismatch = True
    worksheet3.row(3).height = 530
    worksheet3.row(4).height_mismatch = True
    worksheet3.row(4).height = 600
    worksheet3.row(5).height_mismatch = True
    worksheet3.row(5).height = 600
    
    worksheet3.col(0).width = 6000
    worksheet3.col(1).width = 7000
    worksheet3.col(2).width = 7500
    worksheet3.col(3).width = 6000
    worksheet3.col(25).width = 4000
    for col in range(6, 48):
        worksheet3.row(col).height_mismatch = True
        worksheet3.row(col).height = 1000
    for col in range(4, 44):
        worksheet3.col(col).width = 3400
    
    col = 0
    array = ['Chapter', 'Source', 'Fuel Type', 'Description', 'Units']
    for lists in array:
        worksheet3.write(5, col, lists, style4)
        col += 1
        
    col = 23
    array = ['Name of GHG', 'GHG Co2 equv', 'Units']
    for lists in array:
        worksheet3.write_merge(4, 5, col, col, lists, style4)
        col += 1
        
    row = 6
    array = ['Total Plastic polymer resin (IIM)', 'Ink & Paint, etc. (Paint shop)', 'Thinner & Other solvents (Paint shop)', 'Hydraulic Oil Cons', 'Plastic & Poly covers', 'Metals', 'Metals', 'Metals', 'Paper', 'wood', 'Carton', 'Cloth', 'Rejected Poly Covers & Plastic Packaging', 'Aluminium waste', 'Copper waste', 'Iron & Steel waste', 'Wood', 'Carton', 'Cloth', 'Rejected Plastic / Lumps to scrap vendor', 'Food Waste', 'Paint Sludge', 'Used Hydraulic oil', 'Waste Thinner', 'Waste Paint', 'Oil Soaked Cotton', 'Empty Oil / Paint Containers', 'ETP Sludge', 'Electronic Waste (E -Waste)', 'Employee Transport Vehicles Diesel consumption', 'Employee Transport Vehicles Petrol consumption', 'Employee Transport Vehicles CNG consumption', 'Employee Vehicles Diesel consumption', 'Employee Vehicles Petrol consumption', 'Employee Vehicles CNG consumption', 'Business travel by train ', 'Business travel by Air', 'Travel by Road - Diesel', 'Travel by Road - Petrol', 'Goods Transport Vehicles Diesel consumption (MATE to customer)', 'Goods Vehicles Petrol consumption (MATE to customer)', 'Goods Vehicles CNG consumption (MATE to customer)',]
    for lists in array:
        worksheet3.write(row, 2, lists, style5)
        row += 1
        
    row = 6
    array = ['Injection moulding', 'painting', 'painting', 'Hydraulic Oil', 'Raw material & Product packing', 'Aluminium', 'Copper', 'Iron & Steels', 'Raw material & Product packing', 'Raw material & Product packing', 'Raw material & Product packing', 'Raw material & Product packing', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Waste for disposal', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Total Distance', 'Total Air Miles', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport', 'Fuel used for transport']
    for lists in array:
        worksheet3.write(row, 3, lists, style5)
        row += 1
        
    row = 6
    array = ['Kg', 'Litre', 'Litre', 'Litre', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Kg', 'Litre', 'Litre', 'Litre', 'Kg', 'Kg', 'Kg', 'Kg', 'Litre', 'Litre', 'Kg', 'Litre', 'Litre', 'Kg', 'KM', 'Air Mile', 'Litre', 'Litre', 'Litre', 'Litre', 'Kg']
    for lists in array:
        worksheet3.write(row, 4, lists, style5)
        row += 1
        
    row = 6
    for lists in range(1, 43):
        if lists == 21:
            worksheet3.write(26, 23, 'COH4', style5)
        else:
            worksheet3.write(row, 23, 'Co2', style5)
        row += 1
    
    row = 6
    array = [2.2, 3, 3, 3, 2.7, 0.42, 0.181, 1.4, 2.42, 2.61, 2.42, 3.07, 1.91, 0.42, 0.181, 1.41, 2.61, 2.42, 3.07, '', 9.2, 3, 1.07, 3, '', 1.07, 1.77, 3, 1.44025, 2.6, 2.32, 1.51, 2.6, 2.32, 1.51, 0.027, 0.289, 2.6, 2.32, 2.6, 2.32, 1.51]
    for lists in array:
        worksheet3.write(row, 24, lists, style5)
        row += 1
    
    row = 6
    array = ['Kg Co2 equv / kG', 'Kg Co2 equv / L', 'Kg Co2 equv / L', 'Kg Co2 equv / L', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg CO2 equ/kG', 'Kg Co2 equv / Kg', 'Kg Co2 equv / L', 'Kg Co2 equv / L', '', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / KMs', 'Kg Co2 equv / Air Mi', 'Kg Co2 equv / lit', 'Kg Co2 equv / lit', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg', 'Kg Co2 equv / Kg']
    for lists in array:
        worksheet3.write(row, 25, lists, style5)
        row += 1
        
    for row in range(6, 48):
        for col in range(5, 23):
            worksheet3.write(row, col, None, style5)
            
    for row in range(6, 48):
        for col in range(26, 44):
            worksheet3.write(row, col, None, style5)

    ##################################################################################### SHEET 4 ###############################################################################################
    
    worksheet4.write_merge(0, 1, 0, 0, 'Water Data', style6)
    worksheet4.write_merge(3, 3, 0, 19, 'Total Water Consumption', style1)
    worksheet4.write_merge(9, 9, 0, 19, 'Total Treated Water', style1)
    worksheet4.write_merge(16, 16, 0, 19, 'Rain Water Harvesting', style1)
    worksheet4.write(7, 0, 'Total Fresh Water Consumption (A)', style8)
    worksheet4.write(7, 1, 'KL', style3)
    worksheet4.write(14, 0, 'Total Water Reused (B)', style8)
    worksheet4.write(14, 1, 'KL', style3)
    worksheet4.write(24, 0, 'Total Rain Water Harvesting Done (C)', style8)
    worksheet4.write(24, 1, 'KL', style3)
    worksheet4.write(2, 0, None, style4)
    worksheet4.write(2, 1, 'Unit', style4)
    worksheet4.set_panes_frozen(True)
    worksheet4.set_vert_split_pos(2)
    worksheet4.set_horz_split_pos(3)
    
    worksheet4.row(0).height_mismatch = True
    worksheet4.row(0).height = 500
    worksheet4.row(2).height_mismatch = True
    worksheet4.row(2).height = 550
    for row in range(3, 26):
        worksheet4.row(row).height_mismatch = True
        worksheet4.row(row).height = 550
    
    worksheet4.col(0).width = 8700
    for col in range(1, 22):
        worksheet4.col(col).width = 3500
    
    for row_ind in range(4,8):
        for col_ind in range(2,20):
            worksheet4.write(row_ind, col_ind, '', style5)
            
    for row_ind in range(10,15):
        for col_ind in range(2,20):
            worksheet4.write(row_ind, col_ind, '', style5)
            
    for row_ind in range(17,25):
        for col_ind in range(2,20):
            worksheet4.write(row_ind, col_ind, '', style5)
    
    freshwater_array = ['Fresh Water - Drinking', 'Fresh Water - Outsource', 'Fresh Water - Borewell']
    row = 4
    for data in freshwater_array:
        worksheet4.write(row, 0, data, style7)
        worksheet4.write(row, 1, 'KL', style3)
        row += 1
        
    treatedwater_array = ['Treated Water - Disposed', 'Treated Water - ETP', 'Treated Water - STP', 'Total Treated Water']
    row = 10
    for data in treatedwater_array:
        worksheet4.write(row, 0, data, style7)
        worksheet4.write(row, 1, 'KL', style3)
        row += 1
        
    rwh_array = ['Rain Water - Harvesting Pits', 'Rain Water - Pit Capacity', 'Installed Rain Water Harvesting Capacity', 'Rain Water - Average Rainfall', 'Rain Water - Roof Area', 'Rain Water - Non-Roof Area', 'Total Rain Water Harvesting Potential']
    rwh_unit = ['Nos', 'KL', 'KL', 'in mm', 'Sq. mtr.', 'Sq. mtr', 'KL']
    row = 17
    for data in rwh_array:
        worksheet4.write(row, 0, data, style7)
        row += 1
    row = 17
    for data in rwh_unit:
        worksheet4.write(row, 1, data, style3)
        row += 1
        
    drinking_array = []; outsource_array = []; borewell_array = []; total_fwc_array = []; disposed_array = []; ept_array = []; stp_array = []; totaltreated_array = []; waterreused_array = []; pits_array = []; capacity_array = []; installed_capacity_array = []; average_array = []; roofarea_array = []; nonroofarea_array = []; potential_array = []; total_rwh_array = []; emis_rwh_array = []; emis_fwater = []
    freshwater = cost_water.objects.filter(wt_plantname = plantname).values('wt_source').distinct().order_by('wt_source')
    for data in freshwater:
        if data['wt_source'] == 'Fresh_Water':
            if current_month >= start_month:
                while start_month <= total_month:
                    
                    drinking = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if drinking is None:
                        drinking = 0
                    drinking_array.append(drinking)
                    worksheet4.write(4, start_month - 2, round(drinking), style5)
                        
                    outsource = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if outsource is None:
                        outsource = 0
                    outsource_array.append(outsource)
                    worksheet4.write(5, start_month - 2, round(outsource), style5)
                    worksheet2.write(12, start_month + 1, round(outsource), style5)
                    worksheet2.write(12, start_month + 22, round(float(outsource) * 0.59), style5)
                        
                    borewell = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if borewell is None:
                        borewell = 0
                    borewell_array.append(borewell)
                    worksheet4.write(6, start_month - 2, round(borewell), style5)
                    total_fwc_array.append(float(borewell) + float(outsource))
                    worksheet4.write(7, start_month - 2, round(float(borewell) + float(outsource)), style5)
                    worksheet1.write(14, start_month, round(float(borewell) + float(outsource)), style5)
                    emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                    worksheet1.write(14, start_month + 21, round((float(borewell) + float(outsource)) * 0.59), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    
                    drinking = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if drinking is None:
                        drinking = 0
                    drinking_array.append(drinking)
                    worksheet4.write(4, start_month + 10, round(drinking), style5)
                        
                    outsource = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if outsource is None:
                        outsource = 0
                    outsource_array.append(outsource)
                    worksheet4.write(5, start_month + 10, round(outsource), style5)
                    worksheet2.write(12, start_month + 13, round(outsource), style5)
                    worksheet2.write(12, start_month + 34, round(float(outsource) * 0.59), style5)
                        
                    borewell = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if borewell is None:
                        borewell = 0
                    borewell_array.append(borewell)
                    worksheet4.write(6, start_month + 10, round(outsource), style5)
                    total_fwc_array.append(float(borewell) + float(outsource))
                    worksheet4.write(7, start_month + 10, round(float(borewell) + float(outsource)), style5)
                    worksheet1.write(14, start_month + 12, float(outsource) + float(borewell), style5)
                    emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                    worksheet1.write(14, start_month + 33, round((float(outsource) + float(borewell)) * 0.59), style5)
                    
                    start_month += 1
            else:
                while start_month <= total_month:
                    drinking = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if drinking is None:
                        drinking = 0
                    drinking_array.append(drinking)
                    worksheet4.write(4, start_month - 2, round(drinking), style5)
                        
                    outsource = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if outsource is None:
                        outsource = 0
                    outsource_array.append(outsource)
                    worksheet4.write(5, start_month - 2, round(outsource), style5)
                    worksheet2.write(12, start_month + 1, round(outsource), style5)
                    worksheet2.write(12, start_month + 22, round(float(outsource) * 0.59), style5)
                        
                    borewell = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if borewell is None:
                        borewell = 0
                    borewell_array.append(borewell)
                    worksheet4.write(6, start_month - 2, round(borewell), style5)
                    total_fwc_array.append(float(borewell) + float(outsource))
                    worksheet4.write(7, start_month - 2, round(float(borewell) + float(outsource)), style5)
                    worksheet1.write(14, start_month, round(float(borewell) + float(outsource)), style5)
                    emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                    worksheet1.write(14, start_month + 21, round((float(borewell) + float(outsource)) * 0.59), style5)
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    drinking = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if drinking is None:
                        drinking = 0
                    drinking_array.append(drinking)
                    worksheet4.write(4, start_month + 10, round(drinking), style5)
                        
                    outsource = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if outsource is None:
                        outsource = 0
                    outsource_array.append(outsource)
                    worksheet4.write(5, start_month + 10, round(outsource), style5)
                    worksheet2.write(12, start_month + 13, round(outsource), style5)
                    worksheet2.write(12, start_month + 34, round(float(outsource) * 0.59), style5)
                        
                    borewell = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if borewell is None:
                        borewell = 0
                    borewell_array.append(borewell)
                    worksheet4.write(6, start_month + 10, round(outsource), style5)
                    total_fwc_array.append(float(borewell) + float(outsource))
                    worksheet4.write(7, start_month + 10, round(float(borewell) + float(outsource)), style5)
                    worksheet1.write(14, start_month + 12, float(outsource) + float(borewell), style5)
                    emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                    worksheet1.write(14, start_month + 33, round((float(outsource) + float(borewell)) * 0.59), style5)
                    start_month += 1
                    
        if data['wt_source'] == 'Treated Water':
            if current_month >= start_month:
                while start_month <= total_month:
                    
                    disposed = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if disposed is None:
                        disposed = 0
                    disposed_array.append(disposed)
                    worksheet4.write(10, start_month - 2, round(disposed), style5)
                    
                    etp = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if etp is None:
                        etp = 0
                    ept_array.append(etp)
                    worksheet4.write(11, start_month - 2, round(etp), style5)
                    
                    stp = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if stp is None:
                        stp = 0
                    stp_array.append(stp)
                    worksheet4.write(12, start_month - 2, round(stp), style5)
                    totaltreated_array.append(float(stp) + float(etp))
                    worksheet4.write(13, start_month - 2, round(float(stp) + float(etp)), style5)
                    waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                    worksheet4.write(14, start_month - 2, round((float(stp) + float(etp)) - float(disposed)), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    
                    disposed = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if disposed is None:
                        disposed = 0
                    disposed_array.append(disposed)
                    worksheet4.write(10, start_month + 10, round(disposed), style5)
                    
                    etp = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if etp is None:
                        etp = 0
                    ept_array.append(etp)
                    worksheet4.write(11, start_month + 10, round(etp), style5)
                    
                    stp = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if stp is None:
                        stp = 0
                    stp_array.append(stp)
                    worksheet4.write(12, start_month + 10, round(stp), style5)
                    totaltreated_array.append(float(stp) + float(etp))
                    worksheet4.write(13, start_month + 10, round(float(stp) + float(etp)), style5)
                    waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                    worksheet4.write(14, start_month + 10, round((float(stp) + float(etp)) - float(disposed)), style5)
                    
                    start_month += 1
            else:
                while start_month <= total_month:
                    disposed = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if disposed is None:
                        disposed = 0
                    disposed_array.append(disposed)
                    worksheet4.write(10, start_month - 2, round(disposed), style5)
                    
                    etp = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if etp is None:
                        etp = 0
                    ept_array.append(etp)
                    worksheet4.write(11, start_month - 2, round(etp), style5)
                    
                    stp = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if stp is None:
                        stp = 0
                    stp_array.append(stp)
                    worksheet4.write(12, start_month - 2, round(stp), style5)
                    totaltreated_array.append(float(stp) + float(etp))
                    worksheet4.write(13, start_month - 2, round(float(stp) + float(etp)), style5)
                    waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                    worksheet4.write(14, start_month - 2, round((float(stp) + float(etp)) - float(disposed)), style5)
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    disposed = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if disposed is None:
                        disposed = 0
                    disposed_array.append(disposed)
                    worksheet4.write(10, start_month + 10, round(disposed), style5)
                    
                    etp = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if etp is None:
                        etp = 0
                    ept_array.append(etp)
                    worksheet4.write(11, start_month + 10, round(etp), style5)
                    
                    stp = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if stp is None:
                        stp = 0
                    stp_array.append(stp)
                    worksheet4.write(12, start_month + 10, round(stp), style5)
                    totaltreated_array.append(float(stp) + float(etp))
                    worksheet4.write(13, start_month + 10, round(float(stp) + float(etp)), style5)
                    waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                    worksheet4.write(14, start_month + 10, round((float(stp) + float(etp)) - float(disposed)), style5)
                    start_month += 1
                    
        if data['wt_source'] == 'Rain Water':
            if current_month >= start_month:
                while start_month <= total_month:
                    
                    harvestingpits = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if harvestingpits is None:
                        harvestingpits = 0
                    pits_array.append(harvestingpits)
                    worksheet4.write(17, start_month - 2, harvestingpits, style5)
                    
                    pitcapacity = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if pitcapacity is None:
                        pitcapacity = 0
                    capacity_array.append(pitcapacity)
                    worksheet4.write(18, start_month - 2, pitcapacity, style5)
                    installed_capacity = float(harvestingpits) * float(pitcapacity)
                    installed_capacity_array.append(installed_capacity)
                    worksheet4.write(19, start_month - 2, round(installed_capacity), style5)
                    
                    rainfall = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if rainfall is None:
                        rainfall = 0
                    average_array.append(rainfall)
                    worksheet4.write(20, start_month - 2, rainfall, style5)
                    
                    roofarea = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if roofarea is None:
                        roofarea = 0
                    roofarea_array.append(roofarea)
                    worksheet4.write(21, start_month - 2, roofarea, style5)
                    
                    nonroofarea = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if nonroofarea is None:
                        nonroofarea = 0
                    nonroofarea_array.append(nonroofarea)
                    worksheet4.write(22, start_month - 2, nonroofarea, style5)
                    try:
                        potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                    except:
                        potential = 0
                    potential_array.append(potential)
                    worksheet4.write(23, start_month - 2, round(potential), style5)
                    
                    if potential > installed_capacity:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    elif installed_capacity > potential:
                        total_rwh = potential
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    else:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    
                    harvestingpits = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if harvestingpits is None:
                        harvestingpits = 0
                    pits_array.append(harvestingpits)
                    worksheet4.write(17, start_month + 10, harvestingpits, style5)
                    
                    pitcapacity = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if pitcapacity is None:
                        pitcapacity = 0
                    capacity_array.append(pitcapacity)
                    worksheet4.write(18, start_month + 10, pitcapacity, style5)
                    installed_capacity = float(harvestingpits) * float(pitcapacity)
                    installed_capacity_array.append(installed_capacity)
                    worksheet4.write(19, start_month + 10, round(installed_capacity), style5)
                    
                    rainfall = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if rainfall is None:
                        rainfall = 0
                    average_array.append(rainfall)
                    worksheet4.write(20, start_month + 10, rainfall, style5)
                    
                    roofarea = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if roofarea is None:
                        roofarea = 0
                    roofarea_array.append(roofarea)
                    worksheet4.write(21, start_month + 10, roofarea, style5)
                    
                    nonroofarea = cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if nonroofarea is None:
                        nonroofarea = 0
                    nonroofarea_array.append(nonroofarea)
                    worksheet4.write(22, start_month + 10, nonroofarea, style5)
                    try:
                        potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                    except:
                        potential = 0
                    potential_array.append(potential)
                    worksheet4.write(23, start_month + 10, round(potential), style5)
                    
                    if potential > installed_capacity:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    elif installed_capacity > potential:
                        total_rwh = potential
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    else:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    
                    start_month += 1
            else:
                while start_month <= total_month:
                    harvestingpits = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if harvestingpits is None:
                        harvestingpits = 0
                    pits_array.append(harvestingpits)
                    worksheet4.write(17, start_month - 2, harvestingpits, style5)
                    
                    pitcapacity = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if pitcapacity is None:
                        pitcapacity = 0
                    capacity_array.append(pitcapacity)
                    worksheet4.write(18, start_month - 2, pitcapacity, style5)
                    installed_capacity = float(harvestingpits) * float(pitcapacity)
                    installed_capacity_array.append(installed_capacity)
                    worksheet4.write(19, start_month - 2, round(installed_capacity), style5)
                    
                    rainfall = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if rainfall is None:
                        rainfall = 0
                    average_array.append(rainfall)
                    worksheet4.write(20, start_month - 2, rainfall, style5)
                    
                    roofarea = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if roofarea is None:
                        roofarea = 0
                    roofarea_array.append(roofarea)
                    worksheet4.write(21, start_month - 2, roofarea, style5)
                    
                    nonroofarea = cost_water.objects.filter(wt_date__year = current_year - 1, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if nonroofarea is None:
                        nonroofarea = 0
                    nonroofarea_array.append(nonroofarea)
                    worksheet4.write(22, start_month - 2, nonroofarea, style5)
                    try:
                        potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                    except:
                        potential = 0
                    potential_array.append(potential)
                    worksheet4.write(23, start_month - 2, round(potential), style5)
                    
                    if potential > installed_capacity:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    elif installed_capacity > potential:
                        total_rwh = potential
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    else:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month - 2, total_rwh, style5)
                        worksheet1.write(16, start_month, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 21, round(total_rwh * -0.59), style5)
                    start_month += 1
                start_month = 1
                while start_month < 4:
                    harvestingpits = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if harvestingpits is None:
                        harvestingpits = 0
                    pits_array.append(harvestingpits)
                    worksheet4.write(17, start_month + 10, harvestingpits, style5)
                    
                    pitcapacity = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if pitcapacity is None:
                        pitcapacity = 0
                    capacity_array.append(pitcapacity)
                    worksheet4.write(18, start_month + 10, pitcapacity, style5)
                    installed_capacity = float(harvestingpits) * float(pitcapacity)
                    installed_capacity_array.append(installed_capacity)
                    worksheet4.write(19, start_month + 10, round(installed_capacity), style5)
                    
                    rainfall = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if rainfall is None:
                        rainfall = 0
                    average_array.append(rainfall)
                    worksheet4.write(20, start_month + 10, rainfall, style5)
                    
                    roofarea = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if roofarea is None:
                        roofarea = 0
                    roofarea_array.append(roofarea)
                    worksheet4.write(21, start_month + 10, roofarea, style5)
                    
                    nonroofarea = cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if nonroofarea is None:
                        nonroofarea = 0
                    nonroofarea_array.append(nonroofarea)
                    worksheet4.write(22, start_month + 10, nonroofarea, style5)
                    try:
                        potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                    except:
                        potential = 0
                    potential_array.append(potential)
                    worksheet4.write(23, start_month + 10, round(potential), style5)
                    
                    if potential > installed_capacity:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    elif installed_capacity > potential:
                        total_rwh = potential
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    else:
                        total_rwh = installed_capacity
                        total_rwh_array.append(total_rwh)
                        worksheet4.write(24, start_month + 10, total_rwh, style5)
                        worksheet1.write(16, start_month + 12, total_rwh, style5)
                        emis_rwh_array.append(total_rwh * -0.59)
                        worksheet1.write(16, start_month + 33, round(total_rwh * -0.59), style5)
                    start_month += 1
                    
    worksheet4.write(4, 14, round(sum(drinking_array), 2), style5)
    worksheet4.write(5, 14, round(sum(outsource_array), 2), style5)
    worksheet4.write(6, 14, round(sum(borewell_array), 2), style5)
    worksheet4.write(7, 14, round(sum(total_fwc_array), 2), style5)
    worksheet4.write(17, 14, round(sum(pits_array) / current_month, 2), style5)
    worksheet4.write(18, 14, round(sum(capacity_array) / current_month, 2), style5)
    worksheet4.write(20, 14, round(sum(average_array) / current_month, 2), style5)
    worksheet4.write(21, 14, round(sum(roofarea_array) / current_month, 2), style5)
    worksheet4.write(22, 14, round(sum(nonroofarea_array) / current_month, 2), style5)
    worksheet4.write(23, 14, round(sum(potential_array), 2), style5)
    worksheet4.write(24, 14, round(sum(total_rwh_array), 2), style5)
    
    worksheet1.write(14, 16, round(sum(total_fwc_array)), style5)
    worksheet1.write(16, 16, round(sum(total_rwh_array)), style5)
    worksheet1.write(14, 38, round(sum(emis_fwater)), style5)
    worksheet1.write(16, 38, round(sum(emis_rwh_array)), style5)
    
    col = 15; col2 = 17; col3 = 38
    for year in range(1, 6):
        drinking_array = []; outsource_array = []; borewell_array = []; total_fwc_array = []; disposed_array = []; ept_array = []; stp_array = []; totaltreated_array = []; waterreused_array = []; pits_array = []; capacity_array = []; installed_capacity_array = []; average_array = []; roofarea_array = []; nonroofarea_array = []; potential_array = []; total_rwh_array = []; emis_rwh_array = []; emis_fwater = []
        freshwater = cost_water.objects.filter(wt_plantname = plantname).values('wt_source').distinct().order_by('wt_source')
        for data in freshwater:
            if data['wt_source'] == 'Fresh_Water':
                if current_month >= start_month:
                    while start_month <= total_month:
                        
                        drinking = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if drinking is None:
                            drinking = 0
                        drinking_array.append(drinking)
                            
                        outsource = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if outsource is None:
                            outsource = 0
                        outsource_array.append(outsource)
                            
                        borewell = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if borewell is None:
                            borewell = 0
                        borewell_array.append(borewell)
                        worksheet4.write(6, start_month - 2, round(borewell), style5)
                        total_fwc_array.append(float(borewell) + float(outsource))
                        emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                        
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        
                        drinking = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if drinking is None:
                            drinking = 0
                        drinking_array.append(drinking)
                            
                        outsource = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if outsource is None:
                            outsource = 0
                        outsource_array.append(outsource)
                            
                        borewell = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if borewell is None:
                            borewell = 0
                        borewell_array.append(borewell)
                        total_fwc_array.append(float(borewell) + float(outsource))
                        emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                        
                        start_month += 1
                else:
                    while start_month <= total_month:
                        drinking = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if drinking is None:
                            drinking = 0
                        drinking_array.append(drinking)
                            
                        outsource = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if outsource is None:
                            outsource = 0
                        outsource_array.append(outsource)
                            
                        borewell = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if borewell is None:
                            borewell = 0
                        borewell_array.append(borewell)
                        worksheet4.write(6, start_month - 2, round(borewell), style5)
                        total_fwc_array.append(float(borewell) + float(outsource))
                        emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        drinking = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Drinking', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if drinking is None:
                            drinking = 0
                        drinking_array.append(drinking)
                            
                        outsource = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Outsource', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if outsource is None:
                            outsource = 0
                        outsource_array.append(outsource)
                            
                        borewell = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_type = 'Borewell', wt_plantname = plantname).values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if borewell is None:
                            borewell = 0
                        borewell_array.append(borewell)
                        worksheet4.write(6, start_month + 10, round(outsource), style5)
                        total_fwc_array.append(float(borewell) + float(outsource))
                        emis_fwater.append((float(borewell) + float(outsource)) * 0.59)
                        start_month += 1
                        
            if data['wt_source'] == 'Treated Water':
                if current_month >= start_month:
                    while start_month <= total_month:
                        
                        disposed = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if disposed is None:
                            disposed = 0
                        disposed_array.append(disposed)
                        
                        etp = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if etp is None:
                            etp = 0
                        ept_array.append(etp)
                        
                        stp = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if stp is None:
                            stp = 0
                        stp_array.append(stp)
                        totaltreated_array.append(float(stp) + float(etp))
                        waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                        
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        
                        disposed = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if disposed is None:
                            disposed = 0
                        disposed_array.append(disposed)
                        
                        etp = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if etp is None:
                            etp = 0
                        ept_array.append(etp)
                        
                        stp = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if stp is None:
                            stp = 0
                        stp_array.append(stp)
                        totaltreated_array.append(float(stp) + float(etp))
                        waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                        
                        start_month += 1
                else:
                    while start_month <= total_month:
                        disposed = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if disposed is None:
                            disposed = 0
                        disposed_array.append(disposed)
                        
                        etp = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if etp is None:
                            etp = 0
                        ept_array.append(etp)
                        
                        stp = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if stp is None:
                            stp = 0
                        stp_array.append(stp)
                        totaltreated_array.append(float(stp) + float(etp))
                        waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        disposed = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Disposed').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if disposed is None:
                            disposed = 0
                        disposed_array.append(disposed)
                        
                        etp = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'ETP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if etp is None:
                            etp = 0
                        ept_array.append(etp)
                        
                        stp = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'STP').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if stp is None:
                            stp = 0
                        stp_array.append(stp)
                        totaltreated_array.append(float(stp) + float(etp))
                        waterreused_array.append((float(stp) + float(etp)) - float(disposed))
                        start_month += 1
                        
            if data['wt_source'] == 'Rain Water':
                if current_month >= start_month:
                    while start_month <= total_month:
                        
                        harvestingpits = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if harvestingpits is None:
                            harvestingpits = 0
                        pits_array.append(harvestingpits)
                        
                        pitcapacity = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if pitcapacity is None:
                            pitcapacity = 0
                        capacity_array.append(pitcapacity)
                        installed_capacity = float(harvestingpits) * float(pitcapacity)
                        installed_capacity_array.append(installed_capacity)
                        
                        rainfall = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if rainfall is None:
                            rainfall = 0
                        average_array.append(rainfall)
                        
                        roofarea = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if roofarea is None:
                            roofarea = 0
                        roofarea_array.append(roofarea)
                        
                        nonroofarea = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if nonroofarea is None:
                            nonroofarea = 0
                        nonroofarea_array.append(nonroofarea)
                        try:
                            potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                        except:
                            potential = 0
                        potential_array.append(potential)
                        
                        if potential > installed_capacity:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        elif installed_capacity > potential:
                            total_rwh = potential
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        else:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        
                        harvestingpits = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if harvestingpits is None:
                            harvestingpits = 0
                        pits_array.append(harvestingpits)
                        
                        pitcapacity = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if pitcapacity is None:
                            pitcapacity = 0
                        capacity_array.append(pitcapacity)
                        installed_capacity = float(harvestingpits) * float(pitcapacity)
                        installed_capacity_array.append(installed_capacity)
                        
                        rainfall = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if rainfall is None:
                            rainfall = 0
                        average_array.append(rainfall)
                        
                        roofarea = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if roofarea is None:
                            roofarea = 0
                        roofarea_array.append(roofarea)
                        
                        nonroofarea = cost_water.objects.filter(wt_date__year = (current_year + 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if nonroofarea is None:
                            nonroofarea = 0
                        nonroofarea_array.append(nonroofarea)
                        try:
                            potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                        except:
                            potential = 0
                        potential_array.append(potential)
                        
                        if potential > installed_capacity:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        elif installed_capacity > potential:
                            total_rwh = potential
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        else:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        
                        start_month += 1
                else:
                    while start_month <= total_month:
                        harvestingpits = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if harvestingpits is None:
                            harvestingpits = 0
                        pits_array.append(harvestingpits)
                        
                        pitcapacity = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if pitcapacity is None:
                            pitcapacity = 0
                        capacity_array.append(pitcapacity)
                        installed_capacity = float(harvestingpits) * float(pitcapacity)
                        installed_capacity_array.append(installed_capacity)
                        
                        rainfall = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if rainfall is None:
                            rainfall = 0
                        average_array.append(rainfall)
                        
                        roofarea = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if roofarea is None:
                            roofarea = 0
                        roofarea_array.append(roofarea)
                        
                        nonroofarea = cost_water.objects.filter(wt_date__year = (current_year - 1) - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if nonroofarea is None:
                            nonroofarea = 0
                        nonroofarea_array.append(nonroofarea)
                        try:
                            potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                        except:
                            potential = 0
                        potential_array.append(potential)
                        
                        if potential > installed_capacity:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        elif installed_capacity > potential:
                            total_rwh = potential
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        else:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                            emis_rwh_array.append(total_rwh * -0.59)
                        start_month += 1
                    start_month = 1
                    while start_month < 4:
                        harvestingpits = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Harvesting Pits').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if harvestingpits is None:
                            harvestingpits = 0
                        pits_array.append(harvestingpits)
                        
                        pitcapacity = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Pit Capacity').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if pitcapacity is None:
                            pitcapacity = 0
                        capacity_array.append(pitcapacity)
                        installed_capacity = float(harvestingpits) * float(pitcapacity)
                        installed_capacity_array.append(installed_capacity)
                        
                        rainfall = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Average Rainfall').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if rainfall is None:
                            rainfall = 0
                        average_array.append(rainfall)
                        
                        roofarea = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if roofarea is None:
                            roofarea = 0
                        roofarea_array.append(roofarea)
                        
                        nonroofarea = cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_source = data['wt_source'], wt_plantname = plantname, wt_type = 'Non-Roof Area').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                        if nonroofarea is None:
                            nonroofarea = 0
                        nonroofarea_array.append(nonroofarea)
                        try:
                            potential = round(((float(roofarea) + float(nonroofarea)) * float(rainfall) * 0.6) / 1000)
                        except:
                            potential = 0
                        potential_array.append(potential)
                        
                        if potential > installed_capacity:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                        elif installed_capacity > potential:
                            total_rwh = potential
                            total_rwh_array.append(total_rwh)
                        else:
                            total_rwh = installed_capacity
                            total_rwh_array.append(total_rwh)
                        start_month += 1
                        
        worksheet4.write(4, col, round(sum(drinking_array)), style5)
        worksheet4.write(5, col, round(sum(outsource_array)), style5)
        worksheet4.write(6, col, round(sum(borewell_array)), style5)
        worksheet4.write(7, col, round(sum(total_fwc_array)), style5)               
        worksheet4.write(23, col, round(sum(potential_array)), style5)
        worksheet4.write(24, col, round(sum(total_rwh_array)), style5)
        
        worksheet1.write(14, col2, round(sum(total_fwc_array)), style5)
        worksheet1.write(16, col2, round(sum(total_rwh_array)), style5)
        worksheet1.write(14, col3, round(sum(emis_fwater)), style5)
        worksheet1.write(16, col3, round(sum(emis_rwh_array)), style5)
        
        col += 1; col2 += 1; col3 += 1

    ##################################################################################### SHEET 5 ###############################################################################################
    
    worksheet5.write_merge(0, 1, 0, 0, 'Energy Data', style6)
    worksheet5.write(2, 0, None,style4)
    worksheet5.write(2, 1, 'Unit', style4)
    worksheet5.write(10, 0, 'Total Power units used', style10)
    worksheet5.write(11, 0, 'Green power to total Power consumed', style10)
    worksheet5.write(12, 0, '% of Green power to total Power consumed', style10)
    worksheet5.write(15, 0, 'Power cost to sales', style10)
    worksheet5.write(13, 0, 'Manufacturing sale', style9)
    worksheet5.write(14, 0, 'Total power cost', style9)
    worksheet5.set_panes_frozen(True)
    worksheet5.set_vert_split_pos(2)
    worksheet5.set_horz_split_pos(3)
    
    worksheet5.row(0).height_mismatch = True
    worksheet5.row(0).height = 500
    for row in range(2, 22):
        worksheet5.row(row).height_mismatch = True
        worksheet5.row(row).height = 550
        
    worksheet5.col(0).width = 12000
    for col in range(1, 20):
        worksheet5.col(col).width = 3000
    
    row = 3
    array = ['Purchased electricity from Electricity  Authorities (non renewable)', 'Purchased electricity from Third Party (Nuclear - Non-Renewable)', 'Purchased electricity from Third Party (CNG - Non-Renewable)', 'Purchased electricity from Third Party  (Biogas-Renewable)', 'Units  generated in-house by burning fuel (Diesel Generator - Non-Renewable)', 'Purchased electricity from  Second Party (solar - Renewable)', 'Purchased electricity from  Third Party (wind-Renewable)']
    for lists in array:
        worksheet5.write(row, 0, lists, style9)
        row += 1
        
    row = 16
    array = ['Cost Spend for Diesel', 'Cost Spend for Petrol', 'Cost Spend for CNG', 'Cost Spend for LPG', 'Cost Spend for PNG', 'Cost Spend for Water from Out Souced']
    for lists in array:
        worksheet5.write(row, 0, lists, style9)
        row += 1
        
    row = 3
    array28 = ['kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kWh', 'kWh', '%','EUR', 'EUR', '%','EUR', 'EUR', 'EUR', 'EUR', 'EUR', 'EUR',]
    for lists in array28:
        worksheet5.write(row, 1, lists, style5)
        row += 1
        
    percentage_greenpower = []; green_power_consumed = []; diesel_units = []
    cost_diesel = []; cost_petrol = []; cost_lpg = []; cost_cng = []; cost_png = []; cost_waters = []
    if current_month >= start_month:
        while start_month <= total_month:
            
            solar = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if solar is None:
                solar = 0
            
            transformer = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if transformer is None:
                transformer = 0
            wind = Cost_Wind.objects.filter(Wind_year = current_year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
            if wind is None:
                wind = 0
            actual_wind = (transformer * wind) / 100
            greenpower = round(solar + actual_wind)
            worksheet5.write(11, start_month - 2, greenpower, style11)
            green_power_consumed.append(greenpower)
            try:
                worksheet5.write(12, start_month - 2, round(greenpower/transformer, 1), style11)
                percentage_greenpower.append(greenpower/transformer)
            except:
                worksheet5.write(12, start_month - 2, 0, style11)
                percentage_greenpower.append(0)
            
            diesel = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if diesel is None:
                diesel = 0
            worksheet5.write(7, start_month - 2, round(diesel), style5)
            diesel_units.append(diesel)
            
            costdg = list(Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
            total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
            cost_diesel.append(total_diesel_cost)
            worksheet5.write(16, start_month - 2, round(total_diesel_cost), style5)
            
            costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
            total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
            worksheet5.write(17, start_month - 2, round(total_petrol_cost), style5)
            cost_petrol.append(total_petrol_cost)
            
            costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
            total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
            cost_cng.append(total_cng_cost)
            worksheet5.write(18, start_month - 2, round(total_cng_cost), style5)
            
            costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
            total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
            cost_lpg.append(total_lpg_cost)
            worksheet5.write(19, start_month - 2, round(total_lpg_cost), style5)
            
            costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
            total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
            cost_png.append(total_png_cost)
            worksheet5.write(20, start_month - 2, round(total_png_cost), style5)
            
            costwater = list(cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
            water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
            cost_waters.append(water_cost)
            worksheet5.write(21, start_month - 2, round(water_cost), style5)
            
            start_month += 1
        start_month = 1
        while start_month < 4:
            
            solar = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if solar is None:
                solar = 0
            
            transformer = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if transformer is None:
                transformer = 0
            wind = Cost_Wind.objects.filter(Wind_year = current_year + 1, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
            if wind is None:
                wind = 0
            actual_wind = (transformer * wind) / 100
            greenpower = round(solar + actual_wind)
            worksheet5.write(11, start_month + 10, greenpower, style11)
            green_power_consumed.append(greenpower)
            try:
                worksheet5.write(12, start_month + 10, round(greenpower/transformer, 1), style11)
                percentage_greenpower.append(greenpower/transformer)
            except:
                worksheet5.write(12, start_month + 10, 0, style11)
                percentage_greenpower.append(0)
            
            diesel = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if diesel is None:
                diesel = 0
            worksheet5.write(7, start_month + 10, round(diesel), style5)
            diesel_units.append(diesel)
            
            costdg = list(Cost_DG.objects.filter(dg_date__year = current_year + 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
            total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
            cost_diesel.append(total_diesel_cost)
            worksheet5.write(16, start_month + 10, round(total_diesel_cost), style5)
            
            costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year + 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
            total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
            cost_petrol.append(total_petrol_cost)
            worksheet5.write(17, start_month + 10, round(total_petrol_cost), style5)
            
            costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year + 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
            total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
            cost_cng.append(total_cng_cost)
            worksheet5.write(18, start_month + 10, round(total_cng_cost), style5)
            
            costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year + 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
            total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
            cost_lpg.append(total_lpg_cost)
            worksheet5.write(19, start_month + 10, round(total_lpg_cost), style5)
            
            costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year + 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
            total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
            cost_png.append(total_png_cost)
            worksheet5.write(20, start_month + 10, round(total_png_cost), style5)
            
            costwater = list(cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
            water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
            cost_waters.append(water_cost)
            worksheet5.write(21, start_month + 10, round(water_cost), style5)
            
            start_month += 1
    else:
        while start_month <= total_month:
            
            solar = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if solar is None:
                solar = 0
            
            transformer = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if transformer is None:
                transformer = 0
            wind = Cost_Wind.objects.filter(Wind_year = current_year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
            if wind is None:
                wind = 0
            actual_wind = (transformer * wind) / 100
            greenpower = round(solar + actual_wind)
            worksheet5.write(11, start_month - 2, greenpower, style11)
            green_power_consumed.append(greenpower)
            try:
                worksheet5.write(12, start_month - 2, round(greenpower/transformer, 1), style11)
                percentage_greenpower.append(greenpower/transformer)
            except:
                worksheet5.write(12, start_month - 2, 0, style11)
                percentage_greenpower.append(0)
            
            diesel = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if diesel is None:
                diesel = 0
            worksheet5.write(7, start_month - 2, round(diesel), style5)
            diesel_units.append(diesel)
            
            costdg = list(Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
            total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
            cost_diesel.append(total_diesel_cost)
            worksheet5.write(16, start_month - 2, round(total_diesel_cost), style5)
            
            costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
            total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
            worksheet5.write(17, start_month - 2, round(total_petrol_cost), style5)
            cost_petrol.append(total_petrol_cost)
            
            costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
            total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
            cost_cng.append(total_cng_cost)
            worksheet5.write(18, start_month - 2, round(total_cng_cost), style5)
            
            costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
            total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
            cost_lpg.append(total_lpg_cost)
            worksheet5.write(19, start_month - 2, round(total_lpg_cost), style5)
            
            costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
            total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
            cost_png.append(total_png_cost)
            worksheet5.write(20, start_month - 2, round(total_png_cost), style5)
            
            costwater = list(cost_water.objects.filter(wt_date__year = current_year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
            water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
            cost_waters.append(water_cost)
            worksheet5.write(21, start_month - 2, round(water_cost), style5)
            
            start_month += 1
        start_month = 1
        while start_month < 4:
            
            solar = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if solar is None:
                solar = 0
            transformer = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if transformer is None:
                transformer = 0
            
            wind = Cost_Wind.objects.filter(Wind_year = current_year + 1, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
            if wind is None:
                wind = 0
            actual_wind = (transformer * wind) / 100
            greenpower = round(solar + actual_wind)
            worksheet5.write(11, start_month + 10, greenpower, style11)
            green_power_consumed.append(greenpower)
            try:
                worksheet5.write(12, start_month + 10, round(greenpower/transformer, 1), style11)
                percentage_greenpower.append(greenpower/transformer)
            except:
                worksheet5.write(12, start_month + 10, 0, style11)
                percentage_greenpower.append(0)
            
            diesel = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
            if diesel is None:
                diesel = 0
            worksheet5.write(7, start_month + 10, round(diesel), style5)
            diesel_units.append(diesel)
            
            costdg = list(Cost_DG.objects.filter(dg_date__year = current_year + 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
            total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
            cost_diesel.append(total_diesel_cost)
            worksheet5.write(16, start_month + 10, round(total_diesel_cost), style5)
            
            costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year + 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
            total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
            cost_petrol.append(total_petrol_cost)
            worksheet5.write(17, start_month + 10, round(total_petrol_cost), style5)
            
            costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year + 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
            total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
            cost_cng.append(total_cng_cost)
            worksheet5.write(18, start_month + 10, round(total_cng_cost), style5)
            
            costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year + 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
            total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
            cost_lpg.append(total_lpg_cost)
            worksheet5.write(19, start_month + 10, round(total_lpg_cost), style5)
            
            costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year + 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
            total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
            cost_png.append(total_png_cost)
            worksheet5.write(20, start_month + 10, round(total_png_cost), style5)
            
            costwater = list(cost_water.objects.filter(wt_date__year = current_year + 1, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
            water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
            cost_waters.append(water_cost)
            worksheet5.write(21, start_month + 10, round(water_cost), style5)
            
            start_month += 1
    worksheet5.write(7, 14, round(sum(diesel_units)), style5)
    worksheet5.write(11, 14, sum(green_power_consumed), style11)
    worksheet5.write(12, 14, round(sum(percentage_greenpower) / 12, 1), style11)
    worksheet5.write(16, 14, round(sum(cost_diesel)), style5)
    worksheet5.write(17, 14, round(sum(cost_petrol)), style5)
    worksheet5.write(18, 14, round(sum(cost_cng)), style5)
    worksheet5.write(19, 14, round(sum(cost_lpg)), style5)
    worksheet5.write(20, 14, round(sum(cost_png)), style5)
    worksheet5.write(21, 14, round(sum(cost_waters)), style5)
    
    col = 15
    for year in range(1, 6):
        percentage_greenpower = []; green_power_consumed = []; diesel_units = []
        cost_diesel = []; cost_petrol = []; cost_lpg = []; cost_cng = []; cost_png = []; cost_waters = []
        if current_month >= start_month:
            while start_month <= total_month:
                solar = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if solar is None:
                    solar = 0
                
                transformer = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if transformer is None:
                    transformer = 0
                wind = Cost_Wind.objects.filter(Wind_year = current_year - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                if wind is None:
                    wind = 0
                actual_wind = (transformer * wind) / 100
                greenpower = round(float(solar) + float(actual_wind))
                green_power_consumed.append(greenpower)
                try:
                    percentage_greenpower.append(greenpower/transformer)
                except:
                    percentage_greenpower.append(0)
                
                diesel = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if diesel is None:
                    diesel = 0
                diesel_units.append(diesel)
                
                costdg = list(Cost_DG.objects.filter(dg_date__year = current_year - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
                total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
                cost_diesel.append(total_diesel_cost)
                
                costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
                total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
                cost_petrol.append(total_petrol_cost)
                
                costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
                total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
                cost_cng.append(total_cng_cost)
                
                costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
                total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
                cost_lpg.append(total_lpg_cost)
                
                costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
                total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
                cost_png.append(total_png_cost)
                
                costwater = list(cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
                water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
                cost_waters.append(water_cost)
                start_month += 1
            start_month = 1
            while start_month < 4:
                solar = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if solar is None:
                    solar = 0
                
                transformer = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if transformer is None:
                    transformer = 0
                wind = Cost_Wind.objects.filter(Wind_year = (current_year + 1) + year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                if wind is None:
                    wind = 0
                actual_wind = (transformer * wind) / 100
                greenpower = round(float(solar) + float(actual_wind))
                green_power_consumed.append(greenpower)
                try:
                    percentage_greenpower.append(greenpower/transformer)
                except:
                    percentage_greenpower.append(0)
                
                diesel = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if diesel is None:
                    diesel = 0
                diesel_units.append(diesel)
                
                costdg = list(Cost_DG.objects.filter(dg_date__year = (current_year + 1) + year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
                total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
                cost_diesel.append(total_diesel_cost)
                
                costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = (current_year + 1) + year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
                total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
                cost_petrol.append(total_petrol_cost)
                
                costcng = list(Cost_CNG.objects.filter(CNG_date__year = (current_year + 1) + year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
                total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
                cost_cng.append(total_cng_cost)
                
                costlpg = list(Cost_LPG.objects.filter(LPG_date__year = (current_year + 1) + year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
                total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
                cost_lpg.append(total_lpg_cost)
                
                costpng = list(Cost_PNG.objects.filter(PNG_date__year = (current_year + 1) + year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
                total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
                cost_png.append(total_png_cost)
                
                costwater = list(cost_water.objects.filter(wt_date__year = (current_year + 1) + year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
                water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
                cost_waters.append(water_cost)
                start_month += 1
        else:
            while start_month <= total_month:
                solar = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if solar is None:
                    solar = 0
                
                transformer = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if transformer is None:
                    transformer = 0
                wind = Cost_Wind.objects.filter(Wind_year = current_year - year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                if wind is None:
                    wind = 0
                actual_wind = (transformer * wind) / 100
                greenpower = round(float(solar) + float(actual_wind))
                green_power_consumed.append(greenpower)
                try:
                    percentage_greenpower.append(greenpower/transformer)
                except:
                    percentage_greenpower.append(0)
                
                diesel = Masterdatatable.objects.filter(mtdate__year = current_year - year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if diesel is None:
                    diesel = 0
                diesel_units.append(diesel)
                
                costdg = list(Cost_DG.objects.filter(dg_date__year = current_year - year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
                total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
                cost_diesel.append(total_diesel_cost)
                
                costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = current_year - year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
                total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
                cost_petrol.append(total_petrol_cost)
                
                costcng = list(Cost_CNG.objects.filter(CNG_date__year = current_year - year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
                total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
                cost_cng.append(total_cng_cost)
                
                costlpg = list(Cost_LPG.objects.filter(LPG_date__year = current_year - year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
                total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
                cost_lpg.append(total_lpg_cost)
                
                costpng = list(Cost_PNG.objects.filter(PNG_date__year = current_year - year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
                total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
                cost_png.append(total_png_cost)
                
                costwater = list(cost_water.objects.filter(wt_date__year = current_year - year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
                water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
                cost_waters.append(water_cost)
                start_month += 1
            start_month = 1
            while start_month < 4:
                solar = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if solar is None:
                    solar = 0
                
                transformer = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if transformer is None:
                    transformer = 0
                wind = Cost_Wind.objects.filter(Wind_year = (current_year + 1) + year, Wind_month = calendar.month_name[start_month], Wind_plantname = plantname).values('Wind_percentage').aggregate(Sum('Wind_percentage'))['Wind_percentage__sum']
                if wind is None:
                    wind = 0
                actual_wind = (transformer * wind) / 100
                greenpower = round(float(solar) + float(actual_wind))
                green_power_consumed.append(greenpower)
                try:
                    percentage_greenpower.append(greenpower/transformer)
                except:
                    percentage_greenpower.append(0)
                
                diesel = Masterdatatable.objects.filter(mtdate__year = (current_year + 1) + year, mtdate__month = start_month, mtsrcname = 'DG', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                if diesel is None:
                    diesel = 0
                diesel_units.append(diesel)
                
                costdg = list(Cost_DG.objects.filter(dg_date__year = (current_year + 1) + year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons', 'dg_lit_cpl'))
                total_diesel_cost = sum(list(map(lambda x: x['dg_lit_cons']*x['dg_lit_cpl'], costdg)))
                cost_diesel.append(total_diesel_cost)
                
                costpetrol = list(Cost_Petrol.objects.filter(pt_date__year = (current_year + 1) + year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons', 'pt_cpl'))
                total_petrol_cost = sum(list(map(lambda x: x['pt_lit_cons']*x['pt_cpl'], costpetrol)))
                cost_petrol.append(total_petrol_cost)
                
                costcng = list(Cost_CNG.objects.filter(CNG_date__year = (current_year + 1) + year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons', 'CNG_cpkg'))
                total_cng_cost = sum(list(map(lambda x: x['CNG_kg_cons']*x['CNG_cpkg'], costcng)))
                cost_cng.append(total_cng_cost)
                
                costlpg = list(Cost_LPG.objects.filter(LPG_date__year = (current_year + 1) + year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons', 'LPG_cpkg'))
                total_lpg_cost = sum(list(map(lambda x: x['LPG_kg_cons']*x['LPG_cpkg'], costlpg)))
                cost_lpg.append(total_lpg_cost)
                
                costpng = list(Cost_PNG.objects.filter(PNG_date__year = (current_year + 1) + year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons', 'PNG_cpkg'))
                total_png_cost = sum(list(map(lambda x: x['PNG_kg_cons']*x['PNG_cpkg'], costpng)))
                cost_png.append(total_png_cost)
                
                costwater = list(cost_water.objects.filter(wt_date__year = (current_year + 1) + year, wt_date__month = start_month, wt_plantname = plantname).values('wt_consume', 'wt_cost'))
                water_cost = sum(list(map(lambda x: x['wt_consume']*x['wt_cost'], costwater)))
                cost_waters.append(water_cost)
                start_month += 1
                
        worksheet5.write(7, col, round(sum(diesel_units)), style5)
        worksheet5.write(11, col, sum(green_power_consumed), style11)
        worksheet5.write(12, col, round(sum(percentage_greenpower) / 12, 1), style11)
        worksheet5.write(16, col, round(sum(cost_diesel)), style5)
        worksheet5.write(17, col, round(sum(cost_petrol)), style5)
        worksheet5.write(18, col, round(sum(cost_cng)), style5)
        worksheet5.write(19, col, round(sum(cost_lpg)), style5)
        worksheet5.write(20, col, round(sum(cost_png)), style5)
        worksheet5.write(21, col, round(sum(cost_waters)), style5)
        col += 1
        
    for row in range(4, 7):
        for col in range(2, 20):
            worksheet5.write(row, col, None, style5)
            
    for row in range(13, 16):
        for col in range(2, 20):
            worksheet5.write(row, col, None, style5)
    
                    
    ##################################################################################### SHEET 6 ###############################################################################################
     
    worksheet6.write_merge(0, 1, 0, 0, 'Name of the material', style4)
    worksheet6.write_merge(0, 1, 1, 1, 'Details', style4)
    worksheet6.write_merge(0, 1, 2, 2, 'Unit', style4)
    worksheet6.write(8, 0, 'Name of the material', style4)
    worksheet6.write(8, 1, 'Details', style4)
    worksheet6.write(8, 2, 'Unit', style4)
    worksheet6.write_merge(6, 6, 0, 2, 'Total', style3)
    worksheet6.write_merge(17, 17, 0, 2, 'Total', style3)
    worksheet6.write_merge(28, 28, 0, 2, 'Total', style3)
    worksheet6.write_merge(38, 38, 0, 2, 'Total', style3)
    worksheet6.write_merge(9, 16, 0, 0, 'Consumption of Packaging and other consumables', style10)
    worksheet6.write_merge(19, 27, 0, 0, 'Non Hazardous Waste Generation', style10)
    worksheet6.write_merge(30, 37, 0, 0, 'Hazardous waste Generation', style10)
    worksheet6.set_panes_frozen(True)
    worksheet6.set_vert_split_pos(3)
    
    worksheet6.row(0).height_mismatch = True
    worksheet6.row(0).height = 300
    worksheet6.row(6).height_mismatch = True
    worksheet6.row(6).height = 600
    worksheet6.row(17).height_mismatch = True
    worksheet6.row(17).height = 600
    worksheet6.row(28).height_mismatch = True
    worksheet6.row(28).height = 600
    worksheet6.row(38).height_mismatch = True
    worksheet6.row(38).height = 600
    
    worksheet6.col(0).width = 9000
    worksheet6.col(1).width = 6000
    
    row = 2
    array = ['Consumption of Raw material', 'Consumption of Paint & Ink', 'Consumption of Thinner and other solvents', 'Consumption of Hydraulic Oil']
    for lists in array:
        worksheet6.write(row, 0, lists, style10)
        worksheet6.row(row).height_mismatch = True
        worksheet6.row(row).height = 600
        row += 1
        
    row = 2
    array = ['Engg. Plastic/metal etc.', 'paint', 'Thinner', 'Hyd Oil (Top Up)']
    for lists in array:
        worksheet6.write(row, 1, lists, style5)
        row += 1
        
    row = 9
    array = ['Plastic & Poly covers', 'Aluminium', 'Copper', 'Iron & Steel', 'Paper', 'Wood', 'Carton', 'Cloth']
    for lists in array:
        worksheet6.write(row, 1, lists, style5)
        worksheet6.row(row).height_mismatch = True
        worksheet6.row(row).height = 600
        row += 1
        
    row = 2
    array = ['Kg', 'Litre', 'Litre', 'Litre']
    for lists in array:
        worksheet6.write(row, 2, lists, style5)
        row += 1
        
    for lists in range(0, 8):
        worksheet6.write(9 + lists, 2, 'Kg', style5)
        
    row = 19
    array = ['Rejected Poly Covers & Plastic Packaging', 'Aluminium waste', 'Copper waste', 'Iron & Steel waste', 'Wood waste', 'Carton waste', 'Cloth waste', 'Rejected Plastic / Lumps to scrap vendor', 'Food waste']
    for lists in array:
        worksheet6.write(row, 1, lists, style5)
        worksheet6.row(row).height_mismatch = True
        worksheet6.row(row).height = 600
        row += 1
        
    for lists in range(0, 9):
        worksheet6.write(19 + lists, 2, 'Kg', style5)
        
    row = 30
    array = ['Paint Sludge', 'Used Hydraulic oil', 'Waste Thinner', 'Waste Paint', 'Oil Soaked Cotton', 'Empty Oil / Paint Containers', 'ETP Sludge', 'Electronic Waste (E -Waste)']
    for lists in array:
        worksheet6.write(row, 1, lists, style5)
        worksheet6.row(row).height_mismatch = True
        worksheet6.row(row).height = 600
        row += 1
        
    row = 30
    array = ['Kg', 'Litre', 'Litre', 'Litre', 'Kg', 'Kg', 'Kg', 'Kg']
    for lists in array:
        worksheet6.write(row, 2, lists, style5)
        row += 1
        
    col1 = 3; col2 = 5
    if current_month >= start_month:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet6.write_merge(0, 0, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 3
            col2 += 3
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet6.write_merge(0, 0, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 3
            col2 += 3
            start_month += 1
    else:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet6.write_merge(0, 0, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 3
            col2 += 3
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet6.write_merge(0, 0, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 3
            col2 += 3
            start_month += 1
            
    col1 = 39; col2 = 41
    for year in range(0, 6):
        if current_month >= start_month:
            while start_month <= total_month:
                start_month += 1
            x = current_year - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = (current_year + 1) - year
            worksheet6.write_merge(0, 0, col1, col2, str(x) + " - " + str(y), style4)
            col1 +=3
            col2 += 3
        else:
            while start_month <= total_month:
                start_month += 1
            x = (current_year - 1) - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = current_year - year
            worksheet6.write_merge(0, 0, col1, col2, str(x) + " - " + str(y), style4)
            col1 +=3
            col2 += 3
            
    col = 3
    for lists in range(1, 19):
        worksheet6.write(1, col, 'Total Cons', style4)
        worksheet6.col(col).width = 3000
        worksheet6.write(1, col + 1, 'Reused Material Cons', style4)
        worksheet6.col(col + 1).width = 3000
        worksheet6.write(1, col + 2, 'Virgin Material Cons', style4)
        worksheet6.col(col + 2).width = 3000
        worksheet6.write(8, col, 'Initial Waste Generated', style4)
        worksheet6.write(8, col + 1, 'Waste Reuse-Recycled', style4)
        worksheet6.write(8, col + 2, 'Waste Disposed', style4)
        col +=3
        
    for row in range(2, 7):
        for col in range(3, 57):
            worksheet6.write(row, col, None, style5)
            
    for row in range(9, 18):
        for col in range(3, 57):
            worksheet6.write(row, col, None, style5)
            
    for row in range(19, 29):
        for col in range(3, 57):
            worksheet6.write(row, col, None, style5)
    
    for row in range(30, 39):
        for col in range(3, 57):
            worksheet6.write(row, col, None, style5)
            
    # ##################################################################################### SHEET 7 ###############################################################################################
    
    worksheet7.write_merge(0, 0, 0, 3, 'Finished Goods Transport', style6)
    worksheet7.write_merge(2, 3, 0, 0, 'Chapter', style4)
    worksheet7.write_merge(2, 3, 1, 1, 'Details', style4)
    worksheet7.write_merge(2, 3, 2, 2, 'Fuel Type', style4)
    worksheet7.write_merge(2, 3, 3, 3, 'Details', style4)
    worksheet7.write_merge(4, 10, 0, 0, 'VIII', style5)
    worksheet7.write_merge(4, 10, 1, 1, 'Finished Goods Transport', style5)
    worksheet7.set_panes_frozen(True)
    worksheet7.set_vert_split_pos(4)
    
    worksheet7.row(0).height_mismatch = True
    worksheet7.row(0).height = 600
    worksheet7.row(1).height_mismatch = True
    worksheet7.row(1).height = 300
    
    worksheet7.col(1).width = 5500
    worksheet7.col(2).width = 6000
    worksheet7.col(3).width = 5000
    for col in range(4, 66):
        worksheet7.col(col).width = 3000
    
    col1 = 4; col2 = 8
    if current_month >= start_month:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet7.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 5
            col2 += 5
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet7.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 5
            col2 += 5
            start_month += 1
    else:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet7.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 5
            col2 += 5
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet7.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            col1 += 5
            col2 += 5
            start_month += 1
            
    col = 64
    for year in range(0, 2):
        if current_month >= start_month:
            while start_month <= total_month:
                start_month += 1
            x = current_year - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = (current_year + 1) - year
            # worksheet6.write_merge(0, 0, col1, col2, str(x) + " - " + str(y), style4)
            worksheet7.write_merge(2, 3, col, col, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            col += 1
        else:
            while start_month <= total_month:
                start_month += 1
            x = (current_year - 1) - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = current_year - year
            # worksheet6.write_merge(0, 0, col1, col2, str(x) + " - " + str(y), style4)
            worksheet7.write_merge(2, 3, col, col, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            col += 1
            
    col = 0
    for lists in range(0, 12):
        worksheet7.write(3, col + 4, 'No.of trips', style4)
        worksheet7.write(3, col + 5, 'KM per trip', style4)
        worksheet7.write(3, col + 6, 'Total KM', style4)
        worksheet7.write(3, col + 7, 'Mileage of the vehicle (Kms / Lit)', style4)
        worksheet7.write(3, col + 8, 'Total fuel consumption', style4)
        col += 5
        
    for lists in range(4, 9):
        worksheet7.write(lists, 2, 'Finished Goods Transport Vehicles Diesel consumption', style5)
        
    row = 4
    array = ['MATE 1 to MOBIS', 'MATE 1 to Ashok leyland', 'MATE 1 to Bangalore', 'MATE 1 to Unit 3', 'MATE 1 to Ford PDC']
    for lists in array:
        worksheet7.write(row, 3, lists, style5)
        row += 1
        
    for row in range(4, 11):
        for col in range(4, 66):
            worksheet7.write(row, col, None, style5)
            
    for row in range(9, 11):
        for col in range(2, 4):
            worksheet7.write(row, col, None, style5)
            
    # ##################################################################################### SHEET 8 ###############################################################################################
    
    col1 = 4; col2 = 9
    col3 = 2; col4 = 6
    col5 = 2; col6 = 7
    if current_month >= start_month:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet8.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(11, 11, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(14, 14, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(17, 17, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(20, 20, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(23, 23, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(26, 26, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(29, 29, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(32, 32, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(35, 35, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(38, 38, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(42, 42, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(55, 55, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(63, 63, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(76, 76, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(89, 89, col3, col4, month_names + " ' " + str(current_year), style4)
            col5 += 6; col6 += 6
            col1 += 6; col2 += 6
            col3 += 5; col4 += 5
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet8.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(11, 11, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(14, 14, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(17, 17, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(20, 20, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(23, 23, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(26, 26, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(29, 29, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(32, 32, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(35, 35, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(38, 38, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(42, 42, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(55, 55, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(63, 63, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(76, 76, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(89, 89, col3, col4, month_names + " ' " + str(current_year), style4)
            col5 += 6; col6 += 6
            col1 += 6; col2 += 6
            col3 += 5; col4 += 5
            start_month += 1
    else:
        while start_month <= total_month:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet8.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(11, 11, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(14, 14, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(17, 17, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(20, 20, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(23, 23, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(26, 26, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(29, 29, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(32, 32, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(35, 35, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(38, 38, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(42, 42, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(55, 55, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(63, 63, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(76, 76, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(89, 89, col3, col4, month_names + " ' " + str(current_year), style4)
            col5 += 6; col6 += 6
            col1 += 6; col2 += 6
            col3 += 5; col4 += 5
            start_month += 1
        start_month = 1
        while start_month < 4:
            month_names = (calendar.month_name[start_month])[:3]
            worksheet8.write_merge(2, 2, col1, col2, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(11, 11, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(14, 14, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(17, 17, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(20, 20, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(23, 23, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(26, 26, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(29, 29, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(32, 32, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(35, 35, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(38, 38, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(42, 42, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(55, 55, col5, col6, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(63, 63, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(76, 76, col3, col4, month_names + " ' " + str(current_year), style4)
            worksheet8.write_merge(89, 89, col3, col4, month_names + " ' " + str(current_year), style4)
            col5 += 6; col6 += 6
            col1 += 6; col2 += 6
            col3 += 5; col4 += 5
            start_month += 1
            
    col = 76
    col2 = 62
    for year in range(0, 2):
        if current_month >= start_month:
            while start_month <= total_month:
                start_month += 1
            x = current_year - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = (current_year + 1) - year
            worksheet8.write_merge(2, 3, col, col, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(11, 12, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(14, 15, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(17, 18, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(20, 21, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(23, 24, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(26, 27, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(29, 30, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(32, 33, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(35, 36, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(38, 39, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            col += 1; col2 += 1
        else:
            while start_month <= total_month:
                start_month += 1
            x = (current_year - 1) - year
            start_month = 1
            while start_month < 4:
                start_month += 1
            y = current_year - year
            worksheet8.write_merge(2, 3, col, col, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(11, 12, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(14, 15, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(17, 18, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(20, 21, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(23, 24, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(26, 27, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(29, 30, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(32, 33, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(35, 36, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            worksheet8.write_merge(38, 39, col2, col2, 'Total fuel consumption ' + str(x) + " - " + str(y), style4)
            col += 1; col2 += 1
    
    worksheet8.write_merge(0, 0, 0, 3, 'Employee commuting', style6)
    worksheet8.write_merge(2, 2, 0, 3, 'V Employee Travels', style4)
    worksheet8.write_merge(4, 8, 0, 0, 'V', style5)
    worksheet8.write_merge(4, 8, 1, 1, 'Employee Comuting (Common Transport)', style5)
    
    worksheet8.row(0).height_mismatch = True
    worksheet8.row(0).height = 600
    
    for col in range(1, 80):
        worksheet8.col(col).width = 3000
        
    col = 0
    array = ['Chapter', 'Details', 'Fuel Type', 'Details']
    for lists in array:
        worksheet8.write(3, col, lists, style4)
        col += 1
        
    for lists in range(1, 6):
        worksheet8.write(3 + lists, 2, 'Employee Transport Vehicles Diesel consumption', style5)
        
    row = 4
    array = ['MATE 1 to Tambaram', 'MATE 1 to Maduranthagam 1', 'MATE 1 to Maduranthagam 2', 'Fuel used for transport', 'Fuel used for transport']
    for lists in array:
        worksheet8.write(row, 3, lists, style5)
        row += 1
        
    col = 0
    for lists in range(1, 13):
        worksheet8.write(3, col + 4, 'No.of trips', style4)
        worksheet8.write(3, col + 5, 'No.of working days', style4)
        worksheet8.write(3, col + 6, 'KM per trip', style4)
        worksheet8.write(3, col + 7, 'Total KM', style4)
        worksheet8.write(3, col + 8, 'Mileage of the vehicle (Kms / Lit)', style4)
        worksheet8.write(3, col + 9, 'Total Fuel consumption', style4)
        col += 6
        
    for row in range(4, 9):
        for col in range(4, 78):
            worksheet8.write(row, col, None, style5)
    
    worksheet8.write_merge(11, 11, 0, 1, 'VI Employee', style4)
    worksheet8.write(12, 0, 'Chapter', style4)
    worksheet8.write(12, 1, 'Details', style4)
    worksheet8.write(13, 0, 'VI(a)', style5)
    worksheet8.write(13, 1, 'Employee Commuting (Company provided Vehicles for employees)(Two Wheeler)', style5)
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(12, col, 'No. of employees having company provided two wheeler', style4)
        worksheet8.write(12, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(12, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(12, col + 3, 'Fuel price per Lit', style4)
        worksheet8.write(12, col + 4, 'Total fuel consumed', style4)
        col += 5
        
    for col in range(2, 64):
        worksheet8.write(13, col, None, style5)
        worksheet8.write(16, col, None, style5)
        worksheet8.write(19, col, None, style5)
        worksheet8.write(22, col, None, style5)
        worksheet8.write(25, col, None, style5)
        worksheet8.write(28, col, None, style5)
        worksheet8.write(31, col, None, style5)
        worksheet8.write(34, col, None, style5)
        worksheet8.write(37, col, None, style5)
        worksheet8.write(40, col, None, style5)
    
    worksheet8.write_merge(14, 14, 0, 1, 'VI Employee', style4)
    worksheet8.write_merge(15, 22, 0, 0, 'VI(b)', style5)
    worksheet8.write_merge(15, 22, 1, 1, 'Employee Commuting (Company provided Vehicles for employees)(Petrol)', style5)
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(15, col, 'No. of employees having company provided  Car ( CC under 999)', style4)
        worksheet8.write(15, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(15, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(15, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(15, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(18, col, 'No. of employees having company provided car ( CC With in 1000 to 1999)', style4)
        worksheet8.write(18, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(18, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(18, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(18, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(21, col, 'No. of employees having company provided car ( CC 1200 and above)', style4)
        worksheet8.write(21, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(21, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(21, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(21, col + 4, 'Total fuel consumed', style4)
        col += 5
    
    worksheet8.write_merge(23, 23, 0, 1, 'VI Employee', style4)
    worksheet8.write_merge(24, 31, 0, 0, 'VI(c)', style5)
    worksheet8.write_merge(24, 31, 1, 1, 'Employee Commuting (Company provided Vehicles for employees)(Diesel)', style5)
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(24, col, 'No. of employees having company provided  Car ( CC under 999)', style4)
        worksheet8.write(24, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(24, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(24, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(24, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(27, col, 'No. of employees having company provided car ( CC With in 1000 to 1999)', style4)
        worksheet8.write(27, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(27, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(27, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(27, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(30, col, 'No. of employees having company provided car ( CC 1200 and above)', style4)
        worksheet8.write(30, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(30, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(30, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(30, col + 4, 'Total fuel consumed', style4)
        col += 5
    
    worksheet8.write_merge(32, 32, 0, 1, 'VI Employee', style4)
    worksheet8.write_merge(33, 40, 0, 0, 'VI(d)', style5)
    worksheet8.write_merge(33, 40, 1, 1, 'Employee Commuting (Company provided Vehicles for employees)(CNG)', style5)
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(33, col, 'No. of employees having company provided Car ( CC under 999)', style4)
        worksheet8.write(33, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(33, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(33, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(33, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(36, col, 'No. of employees having company provided car ( CC With in 1000 to 1999)', style4)
        worksheet8.write(36, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(36, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(36, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(36, col + 4, 'Total fuel consumed', style4)
        worksheet8.write(39, col, 'No. of employees having company provided car ( CC 1200 and above)', style4)
        worksheet8.write(39, col + 1, 'Fuel allowance provided to per employee', style4)
        worksheet8.write(39, col + 2, 'Total allowance for the month', style4)
        worksheet8.write(39, col + 3, 'Fuel price per lit', style4)
        worksheet8.write(39, col + 4, 'Total fuel consumed', style4)
        col += 5
        
    worksheet8.write_merge(42, 42, 0, 1, 'VII Employee Business Travels', style4)
    worksheet8.write_merge(43, 54, 0, 0, 'VII(a)', style5)
    worksheet8.write_merge(43, 54, 1, 1, 'Business Travels-Train', style5)
    worksheet8.write_merge(55, 62, 0, 0, 'VII(b)', style5)
    worksheet8.write_merge(55, 62, 1, 1, 'Business Travels-Air', style5)
    worksheet8.write_merge(63, 75, 0, 0, 'VII(c)', style5)
    worksheet8.write_merge(63, 75, 1, 1, 'Business Travels-Road(Petrol)', style5)
    worksheet8.write_merge(76, 88, 0, 0, 'VII(d)', style5)
    worksheet8.write_merge(76, 88, 1, 1, 'Business Travels-Road(Diesel)', style5)
    worksheet8.write_merge(89, 102, 0, 0, 'VII(e)', style5)
    worksheet8.write_merge(89, 102, 1, 1, 'Business Travels-Road-Two wheeler(Petrol)', style5)
    
    if current_month >= start_month:
        while start_month <= total_month:
            start_month += 1
        x = current_year
        start_month = 1
        while start_month < 4:
            start_month += 1
        y = current_year + 1
    else:
        while start_month <= total_month:
            start_month += 1
        x = current_year - 1
        start_month = 1
        while start_month < 4:
            start_month += 1
        y = current_year
    worksheet8.write_merge(42, 43, 74, 74, 'Total Emission '+ str(x) + ' ' + str(y), style4)
    worksheet8.write_merge(55, 56, 74, 74, 'Total Emission '+ str(x) + ' ' + str(y), style4)
    worksheet8.write_merge(63, 64, 62, 62, 'Total Emission '+ str(x) + ' ' + str(y), style4)
    worksheet8.write_merge(76, 77, 62, 62, 'Total Emission '+ str(x) + ' ' + str(y), style4)
    worksheet8.write_merge(89, 90, 62, 62, 'Total Emission '+ str(x) + ' ' + str(y), style4)
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(43, col, 'Business travel by train', style4)
        worksheet8.write(43, col + 1, 'Name of employees', style4)
        worksheet8.write(43, col + 2, 'Traveling destination', style4)
        worksheet8.write(43, col + 3, 'Total Kms', style4)
        worksheet8.write(43, col + 4, 'Emission per Kms', style4)
        worksheet8.write(43, col + 5, 'Total Emission', style4)
        worksheet8.write(56, col, 'Business travel by train', style4)
        worksheet8.write(56, col + 1, 'Name of employees', style4)
        worksheet8.write(56, col + 2, 'Traveling destination', style4)
        worksheet8.write(56, col + 3, 'Total Kms', style4)
        worksheet8.write(56, col + 4, 'Emission per Kms', style4)
        worksheet8.write(56, col + 5, 'Total Emission', style4)
        col += 6
    
    col = 2
    for lists in range(1, 13):
        worksheet8.write(64, col, 'Business travel by Road', style4)
        worksheet8.write(64, col + 1, 'Name of employees', style4)
        worksheet8.write(64, col + 2, 'Traveling destination', style4)
        worksheet8.write(64, col + 3, 'Total Kms', style4)
        worksheet8.write(64, col + 4, 'Fuel consumed', style4)
        worksheet8.write(77, col, 'Business travel by Road', style4)
        worksheet8.write(77, col + 1, 'Name of employees', style4)
        worksheet8.write(77, col + 2, 'Traveling destination', style4)
        worksheet8.write(77, col + 3, 'Total Kms', style4)
        worksheet8.write(77, col + 4, 'Fuel consumed', style4)
        worksheet8.write(90, col, 'Business travel by Road', style4)
        worksheet8.write(90, col + 1, 'Name of employees', style4)
        worksheet8.write(90, col + 2, 'Traveling destination', style4)
        worksheet8.write(90, col + 3, 'Total Kms', style4)
        worksheet8.write(90, col + 4, 'Fuel consumed', style4)
        col += 5
        
    for row in range(44, 55):
        for col in range(2, 75):
            worksheet8.write(row, col, None, style5)
            
    for row in range(57, 63):
        for col in range(2, 75):
            worksheet8.write(row, col, None, style5)
            
    for row in range(65, 76):
        for col in range(2, 63):
            worksheet8.write(row, col, None, style5)
            
    for row in range(78, 89):
        for col in range(2, 63):
            worksheet8.write(row, col, None, style5)
        
    for row in range(91, 103):
        for col in range(2, 63):
            worksheet8.write(row, col, None, style5)
    
    savepath = os.path.join(BASE_DIR, 'Report', 'carbonexcel.xls')
    workbook.save(savepath)
    return savepath

@csrf_exempt
def sendmail(request):
    if request.method == 'POST':
        reportpath = carbonexcel(request)
        request_data = json.loads(request.body)
        tomail = request_data['email']
        sub = 'Carbon Emission Report'
        messg = '''
        Dear Team, 
            Please find the attached Emission report.
        Regards, 
        EMs Team - ROBIS
        '''
        #Adding the email details to the message header of the MIMEMultipart
        msg = MIMEMultipart('alternative')
        msg['Subject']=sub
        msg['From']="R-Energy <no-reply@renergy.com>"
        msg['To']=tomail
        text = str(messg)
        part1=MIMEText(text,'plain')
        msg.attach(part1)

        #Reading the file and encoding it back to the xls format
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(reportpath, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="R-CarbonReport.xls"')
        msg.attach(part)
        
        #Establishing connection with the smtp server
        smtpObj = smtplib.SMTP(EMAIL_HOST,EMAIL_PORT)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
        smtpObj.sendmail('R-Energy <no-reply@renergy.com>',tomail.split(';'),msg.as_string()) 
        smtpObj.quit()
    return JsonResponse("Mail Sent", safe=False)