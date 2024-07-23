import json
from meter_data.models import Masterdatatable
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from encrypt.renderers import CustomAesRenderer
from django.db.models.aggregates import Sum
from source_management.models import AddSource
from datetime import datetime, timedelta
from costestimator.models import Cost_DG, Cost_Petrol, Cost_LPG, Cost_CNG, Cost_PNG
from rest_framework.decorators import authentication_classes, permission_classes, api_view,renderer_classes
from encrypt.renderers import CustomAesRenderer
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import IsAuthenticated

# @renderer_classes([CustomAesRenderer])
@api_view(['GET', 'POST', 'PUT', 'DELETE'])
@authentication_classes([TokenAuthentication])
@permission_classes([IsAuthenticated])
# @csrf_exempt
def Analytics(request):
    if request.method == 'POST':
        
        plantname = request.GET['plantname']
        request_data = json.loads(request.body)
        filetype = request_data['type']
        
        current_date = datetime.now()
        current_month = current_date.month
        current_year = current_date.year
        previous_year = current_year - 1
        
        sourcename = AddSource.objects.filter(asplantname = plantname).values('assourcename')
        data_names = []
        thismonth = []
        lastmonth = []
        last6months = []
        thisyear = []
        lastyear = []
        for data in sourcename:
            
            # Transformer
            if data['assourcename'] == 'Transformer1':
                data_names.append(data['assourcename'])
                try:
                    transformer_thismonth = (Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    thismonth.append(round(float(transformer_thismonth) * 0.7132))
                except:
                    thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        transformer_lastmonth = (Masterdatatable.objects.filter(mtdate__year = previous_year, mtdate__month = 12, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    else:
                        transformer_lastmonth = (Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month - 1, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    lastmonth.append(round(float(transformer_lastmonth) * 0.7132))
                except:
                    lastmonth.append(0)
                
                
                transformer_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    transformer_6months = Masterdatatable.objects.filter(mtdate__year = target_year, mtdate__month = target_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        transformer_6months_array.append(round(float(transformer_6months['mtenergycons__sum']) * 0.7132))
                    except:
                        transformer_6months_array.append(0)
                last6months.append(sum(transformer_6months_array))
                
                transformer_thisyear_array = []
                for month in range(1, 4):
                    transformer_thisyear = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        transformer_thisyear_array.append(round(float(transformer_thisyear['mtenergycons__sum']) * 0.7132))
                    except:
                        transformer_thisyear_array.append(0)
                for month in range(4, 13):
                    transformer_thisyear = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        transformer_thisyear_array.append(round(float(transformer_thisyear['mtenergycons__sum']) * 0.7132))
                    except:
                        transformer_thisyear_array.append(0)
                thisyear.append(sum(transformer_thisyear_array))
                
                transformer_lastyear_array = []
                for month in range(1, 4):
                    transformer_lastyear = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        transformer_lastyear_array.append(round(float(transformer_lastyear['mtenergycons__sum']) * 0.7132))
                    except:
                        transformer_lastyear_array.append(0)
                for month in range(4, 13):
                    transformer_lastyear = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        transformer_lastyear_array.append(round(float(transformer_lastyear['mtenergycons__sum']) * 0.7132))
                    except:
                        transformer_lastyear_array.append(0)
                lastyear.append(sum(transformer_lastyear_array))
                    
            # Solar Energy
            if data['assourcename'] == 'Solar Energy':
                data_names.append(data['assourcename'])
                try:
                    solar_thismonth = (Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    thismonth.append(round(float(solar_thismonth) * 0.041))
                except:
                    thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        solar_lastmonth = (Masterdatatable.objects.filter(mtdate__year = previous_year, mtdate__month = 12, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    else:
                        solar_lastmonth = (Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month - 1, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
                    lastmonth.append(round(float(solar_lastmonth) * 0.041))
                except:
                    lastmonth.append(0)
                
                
                solar_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    solar_6months = Masterdatatable.objects.filter(mtdate__year = target_year, mtdate__month = target_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        solar_6months_array.append(round(float(solar_6months['mtenergycons__sum']) * 0.041))
                    except:
                        solar_6months_array.append(0)
                last6months.append(sum(solar_6months_array))
                
                solar_thisyear_array = []
                for month in range(1, 4):
                    solar_thisyear = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        solar_thisyear_array.append(round(float(solar_thisyear['mtenergycons__sum']) * 0.041))
                    except:
                        solar_thisyear_array.append(0)
                for month in range(4, 13):
                    solar_thisyear = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        solar_thisyear_array.append(round(float(solar_thisyear['mtenergycons__sum']) * 0.041))
                    except:
                        solar_thisyear_array.append(0)
                thisyear.append(sum(solar_thisyear_array))
                
                solar_lastyear_array = []
                for month in range(1, 4):
                    solar_lastyear = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        solar_lastyear_array.append(round(float(solar_lastyear['mtenergycons__sum']) * 0.041))
                    except:
                        solar_lastyear_array.append(0)
                for month in range(4, 13):
                    solar_lastyear = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    try:
                        solar_lastyear_array.append(round(float(solar_lastyear['mtenergycons__sum']) * 0.041))
                    except:
                        solar_lastyear_array.append(0)
                lastyear.append(sum(solar_lastyear_array))
                
            # Diesel
            if data['assourcename'] == 'DG':
                data_names.append(data['assourcename'])
                try:
                    diesel_thismonth = (Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = current_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons')))['dg_lit_cons__sum']
                    thismonth.append(round(float(diesel_thismonth)* 2.6))
                except:
                   thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        diesel_lastmonth = (Cost_DG.objects.filter(dg_date__year = previous_year, dg_date__month = 12, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons')))['dg_lit_cons__sum']
                    else:
                        diesel_lastmonth = (Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = current_month - 1, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons')))['dg_lit_cons__sum']
                    lastmonth.append(round(float(diesel_lastmonth) * 2.6))
                except:
                    lastmonth.append(0)
                
                
                diesel_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    diesel_6months = Cost_DG.objects.filter(dg_date__year = target_year, dg_date__month = target_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    try:
                        diesel_6months_array.append(round(float(diesel_6months['dg_lit_cons__sum']) * 2.6))
                    except:
                        diesel_6months_array.append(0)
                last6months.append(sum(diesel_6months_array))
                        
                diesel_thisyear_array = []
                for month in range(1, 4):
                    diesel_thisyear = Cost_DG.objects.filter(dg_date__year = current_year + 1, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    try:
                        diesel_thisyear_array.append(round(float(diesel_thisyear['dg_lit_cons__sum']) * 2.6))
                    except:
                        diesel_thisyear_array.append(0)
                for month in range(4, 13):
                    diesel_thisyear = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    try:
                        diesel_thisyear_array.append(round(float(diesel_thisyear['dg_lit_cons__sum']) * 2.6))
                    except:
                        diesel_thisyear_array.append(0)
                thisyear.append(sum(diesel_thisyear_array))
                
                diesel_lastyear_array = []
                for month in range(1, 4):
                    diesel_lastyear = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    try:
                        diesel_lastyear_array.append(round(float(diesel_lastyear['dg_lit_cons__sum']) * 2.6))
                    except:
                        diesel_lastyear_array.append(0)
                for month in range(4, 13):
                    diesel_lastyear = Cost_DG.objects.filter(dg_date__year = current_year - 1, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    try:
                        diesel_lastyear_array.append(round(float(diesel_lastyear['dg_lit_cons__sum']) * 2.6))
                    except:
                        diesel_lastyear_array.append(0)
                lastyear.append(sum(diesel_lastyear_array))
                
            # Petrol
            if data['assourcename'] == 'Petrol':
                data_names.append(data['assourcename'])
                try:
                    petrol_thismonth = (Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = current_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons')))['pt_lit_cons__sum']
                    thismonth.append(round(float(petrol_thismonth) * 2.32))
                except:
                    thismonth.append(0)
                
                try:
                    if current_month == 1:
                        petrol_lastmonth = (Cost_Petrol.objects.filter(pt_date__year = previous_year, pt_date__month = 12, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons')))['pt_lit_cons__sum']
                    else:
                        petrol_lastmonth = (Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = current_month - 1, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons')))['pt_lit_cons__sum']
                    lastmonth.append(round(float(petrol_lastmonth) * 2.32))
                except:
                    lastmonth.append(0)
                
                petrol_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    petrol_6months = Cost_Petrol.objects.filter(pt_date__year = target_year, pt_date__month = target_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    try:
                        petrol_6months_array.append(round(float(petrol_6months['pt_lit_cons__sum']) * 2.32))
                    except:
                        petrol_6months_array.append(0)
                last6months.append(sum(petrol_6months_array))
                
                petrol_thisyear_array = []
                for month in range(1, 4):
                    petrol_thisyear = Cost_Petrol.objects.filter(pt_date__year = current_year + 1, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    try:
                        petrol_thisyear_array.append(round(float(petrol_thisyear['pt_lit_cons__sum']) * 2.32))
                    except:
                        petrol_thisyear_array.append(0)
                for month in range(4, 13):
                    petrol_thisyear = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    try:
                        petrol_thisyear_array.append(round(float(petrol_thisyear['pt_lit_cons__sum']) * 2.32))
                    except:
                        petrol_thisyear_array.append(0)
                thisyear.append(sum(petrol_thisyear_array))
                
                petrol_lastyear_array = []
                for month in range(1, 4):
                    petrol_lastyear = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    try:
                        petrol_lastyear_array.append(round(float(petrol_lastyear['pt_lit_cons__sum']) * 2.32))
                    except:
                        petrol_lastyear_array.append(0)
                for month in range(4, 13):
                    petrol_lastyear = Cost_Petrol.objects.filter(pt_date__year = current_year - 1, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    try:
                        petrol_lastyear_array.append(round(float(petrol_lastyear['pt_lit_cons__sum']) * 2.32))
                    except:
                        petrol_lastyear_array.append(0)
                lastyear.append(sum(petrol_lastyear_array))
                
            # LPG
            if data['assourcename'] == 'LPG':
                data_names.append(data['assourcename'])
                try:
                    lpg_thismonth = (Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = current_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons')))['LPG_kg_cons__sum']
                    thismonth.append(round(float(lpg_thismonth) * 2.19))
                except:
                    thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        lpg_lastmonth = (Cost_LPG.objects.filter(LPG_date__year = previous_year, LPG_date__month = 12, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons')))['LPG_kg_cons__sum']
                    else:
                        lpg_lastmonth = (Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = current_month - 1, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons')))['LPG_kg_cons__sum']
                    lastmonth.append(round(float(lpg_lastmonth) * 2.19))
                except:
                   lastmonth.append(0)
                
                
                lpg_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    lpg_6months = Cost_LPG.objects.filter(LPG_date__year = target_year, LPG_date__month = target_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    try:
                        lpg_6months_array.append(round(float(lpg_6months['LPG_kg_cons__sum']) * 2.19))
                    except:
                        lpg_6months_array.append(0)
                last6months.append(sum(lpg_6months_array))
                
                lpg_thisyear_array = []
                for month in range(1, 4):
                    lpg_thisyear = Cost_LPG.objects.filter(LPG_date__year = current_year + 1, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    try:
                        lpg_thisyear_array.append(round(float(lpg_thisyear['LPG_kg_cons__sum']) * 2.19))
                    except:
                        lpg_thisyear_array.append(0)
                for month in range(4, 13):
                    lpg_thisyear = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    try:
                        lpg_thisyear_array.append(round(float(lpg_thisyear['LPG_kg_cons__sum']) * 2.19))
                    except:
                        lpg_thisyear_array.append(0)
                thisyear.append(sum(lpg_thisyear_array))
                
                lpg_lastyear_array = []
                for month in range(1, 4):
                    lpg_lastyear = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    try:
                        lpg_lastyear_array.append(round(float(lpg_lastyear['LPG_kg_cons__sum']) * 2.19))
                    except:
                        lpg_lastyear_array.append(0)
                for month in range(4, 13):
                    lpg_lastyear = Cost_LPG.objects.filter(LPG_date__year = current_year - 1, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    try:
                        lpg_lastyear_array.append(round(float(lpg_lastyear['LPG_kg_cons__sum']) * 2.19))
                    except:
                        lpg_lastyear_array.append(0)
                lastyear.append(sum(lpg_lastyear_array))
                
            # CNG
            if data['assourcename'] == 'CNG':
                data_names.append(data['assourcename'])
                try:
                    cng_thismonth = (Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = current_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons')))['CNG_kg_cons__sum']
                    thismonth.append(round(float(cng_thismonth) * 0.614))
                except:
                    thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        cng_lastmonth = (Cost_CNG.objects.filter(CNG_date__year = previous_year, CNG_date__month = 12, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons')))['CNG_kg_cons__sum']
                    else:
                        cng_lastmonth = (Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = current_month - 1, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons')))['CNG_kg_cons__sum']
                    lastmonth.append(round(float(cng_lastmonth) * 0.614))
                except:
                    lastmonth.append(0)
                
                
                cng_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    cng_6months = Cost_CNG.objects.filter(CNG_date__year = target_year, CNG_date__month = target_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    try:
                        cng_6months_array.append(round(float(cng_6months['CNG_kg_cons__sum']) * 0.614))
                    except:
                        cng_6months_array.append(0)
                last6months.append(sum(cng_6months_array))
                
                cng_thisyear_array = []
                for month in range(1, 4):
                    cng_thisyear = Cost_CNG.objects.filter(CNG_date__year = current_year + 1, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    try:
                        cng_thisyear_array.append(round(float(cng_thisyear['CNG_kg_cons__sum']) * 0.614))
                    except:
                        cng_thisyear_array.append(0)
                for month in range(4, 13):
                    cng_thisyear = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    try:
                        cng_thisyear_array.append(round(float(cng_thisyear['CNG_kg_cons__sum']) * 0.614))
                    except:
                        cng_thisyear_array.append(0)
                thisyear.append(sum(cng_thisyear_array))
                
                cng_lastyear_array = []
                for month in range(1, 4):
                    cng_lastyear = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    try:
                        cng_lastyear_array.append(round(float(cng_lastyear['CNG_kg_cons__sum']) * 0.614))
                    except:
                        cng_lastyear_array.append(0)
                for month in range(4, 13):
                    cng_lastyear = Cost_CNG.objects.filter(CNG_date__year = current_year - 1, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    try:
                        cng_lastyear_array.append(round(float(cng_lastyear['CNG_kg_cons__sum']) * 0.614))
                    except:
                        cng_lastyear_array.append(0)
                lastyear.append(sum(cng_lastyear_array))
                
            # PNG
            if data['assourcename'] == 'PNG':
                data_names.append(data['assourcename'])
                try:
                    png_thismonth = (Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = current_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons')))['PNG_kg_cons__sum']
                    thismonth.append(round(float(png_thismonth) * 0.056))
                except:
                    thismonth.append(0)
                
                
                try:
                    if current_month == 1:
                        png_lastmonth = (Cost_PNG.objects.filter(PNG_date__year = previous_year, PNG_date__month = 12, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons')))['PNG_kg_cons__sum']
                    else:
                        png_lastmonth = (Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = current_month - 1, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons')))['PNG_kg_cons__sum']
                    lastmonth.append(round(float(png_lastmonth) * 0.056))
                except:
                    lastmonth.append(0)
                
                
                png_6months_array = []
                for month in range(1, 7):
                    target_month = current_month - month
                    target_year = current_year
                    if target_month <= 0:
                        target_month += 12
                        target_year -= 1
                    png_6months = Cost_PNG.objects.filter(PNG_date__year = target_year, PNG_date__month = target_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    try:
                        png_6months_array.append(round(float(png_6months['PNG_kg_cons__sum']) * 0.056))
                    except:
                        png_6months_array.append(0)
                last6months.append(sum(png_6months_array))
                
                png_thisyear_array = []
                for month in range(1, 4):
                    png_thisyear = Cost_PNG.objects.filter(PNG_date__year = current_year + 1, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    try:
                        png_thisyear_array.append(round(float(png_thisyear['PNG_kg_cons__sum']) * 0.056))
                    except:
                        png_thisyear_array.append(0)
                for month in range(4, 13):
                    png_thisyear = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    try:
                        png_thisyear_array.append(round(float(png_thisyear['PNG_kg_cons__sum']) * 0.056))
                    except:
                        png_thisyear_array.append(0)
                thisyear.append(sum(png_thisyear_array))
                
                png_lastyear_array = []
                for month in range(1, 4):
                    png_lastyear = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    try:
                        png_lastyear_array.append(round(float(png_lastyear['PNG_kg_cons__sum']) * 0.056))
                    except:
                        png_lastyear_array.append(0)
                for month in range(4, 13):
                    png_lastyear = Cost_PNG.objects.filter(PNG_date__year = current_year - 1, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    try:
                        png_lastyear_array.append(round(float(png_lastyear['PNG_kg_cons__sum']) * 0.056))
                    except:
                        png_lastyear_array.append(0)
                lastyear.append(sum(png_lastyear_array))
                
        if filetype == 'thismonth':
            values = thismonth
        if filetype == 'lastmonth':
            values = lastmonth
        if filetype == 'last6months':
            values = last6months
        if filetype == 'thisyear':
            values = thisyear
        if filetype == 'lastyear':
            values = lastyear
                
        my_dict = {
            'values' : values,
            'Data_Names' : data_names
        }
        return JsonResponse(my_dict, safe=False)