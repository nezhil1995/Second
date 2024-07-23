from django.http import JsonResponse
import json
from source_management.models import AddSource
from meter_data.models import Masterdatatable
from general_settings.models import Timingtable
from datetime import datetime, timedelta
from costestimator.models import Cost_Wind, cost_water, Cost_DG, Cost_LPG, Cost_PNG, Cost_CNG, Cost_Petrol
from django.db.models import Sum
import calendar
from rest_framework.decorators import authentication_classes, permission_classes, api_view,renderer_classes
from encrypt.renderers import CustomAesRenderer
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import IsAuthenticated

# @renderer_classes([CustomAesRenderer])
@api_view(['GET', 'POST', 'PUT', 'DELETE'])
@authentication_classes([TokenAuthentication])
@permission_classes([IsAuthenticated])
def carbongraph(request):

    unitname = request.GET['plantname']
    request_data = json.loads(request.body)

    current_time = datetime.now()
    timing = list(Timingtable.objects.filter(ttplntname = unitname).values('ttdaystarttime'))[0]['ttdaystarttime']
    srcname = list(AddSource.objects.filter(asplantname = unitname).values('assourcename'))
    
    if request_data['type'] == 'date basis':
        date = request_data['date']
        month = int(request_data['month'])
        year = int(request_data['year'])
        monthname = int(date[5:7])
        start_time = datetime.combine(current_time.date(), timing)
        end_time = start_time + timedelta(days=1)
        label_array = []; carbon = []
        while start_time < end_time:
            label_array.append(start_time.strftime("%H-%M"))
            start_time += timedelta(hours = 1)
        try:
            wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = year, Wind_month = calendar.month_name[month], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
        except:
            wind_percentage = 0
        for source in srcname:
            if source['assourcename'] == 'Transformer1':
                transdata = []
                masterdata = list(Masterdatatable.objects.filter(mtdate = date, mtplntlctn = unitname, mtgrpname = 'Incomer', mtcategory = 'Secondary', mtsrcname = source['assourcename']).values('mth1ec', 'mth2ec', 'mth3ec', 'mth4ec', 'mth5ec', 'mth6ec', 'mth7ec', 'mth8ec', 'mth9ec', 'mth10ec', 'mth11ec', 'mth12ec', 'mth13ec', 'mth14ec', 'mth15ec', 'mth16ec', 'mth17ec', 'mth18ec', 'mth19ec', 'mth20ec', 'mth21ec', 'mth22ec', 'mth23ec', 'mth24ec'))
                for i in range(1, 25):
                    hours = 'mth' + str(i) + 'ec'
                    for energy in masterdata:
                        energydata = energy[hours]
                        wind = (float(energydata) * wind_percentage) / 100
                        transenergy = float(energydata) - wind
                        try:
                            transdata.append(round(float(transenergy) * 0.7132))
                        except:
                            transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            # if source['assourcename'] == 'DG':
            #     transdata = []
            #     masterdata = list(Masterdatatable.objects.filter(mtdate = date, mtplntlctn = unitname, mtgrpname = 'Incomer', mtcategory = 'Secondary', mtsrcname = source['assourcename']).values('mth1ec', 'mth2ec', 'mth3ec', 'mth4ec', 'mth5ec', 'mth6ec', 'mth7ec', 'mth8ec', 'mth9ec', 'mth10ec', 'mth11ec', 'mth12ec', 'mth13ec', 'mth14ec', 'mth15ec', 'mth16ec', 'mth17ec', 'mth18ec', 'mth19ec', 'mth20ec', 'mth21ec', 'mth22ec', 'mth23ec', 'mth24ec'))
            #     for i in range(1, 25):
            #         hours = 'mth' + str(i) + 'ec'
            #         for energy in masterdata:
            #             try:
            #                 transdata.append(round(float(energy[hours]) * 2.6))
            #             except:
            #                 transdata.append(0)
            #     my_dict = {
            #         "name": source['assourcename'],
            #         "data": transdata
            #     }
            #     carbon.append(my_dict)

            if source['assourcename'] == 'Solar Energy':
                transdata = []
                masterdata = list(Masterdatatable.objects.filter(mtdate = date, mtplntlctn = unitname, mtgrpname = 'Incomer', mtcategory = 'Secondary', mtsrcname = source['assourcename']).values('mth1ec', 'mth2ec', 'mth3ec', 'mth4ec', 'mth5ec', 'mth6ec', 'mth7ec', 'mth8ec', 'mth9ec', 'mth10ec', 'mth11ec', 'mth12ec', 'mth13ec', 'mth14ec', 'mth15ec', 'mth16ec', 'mth17ec', 'mth18ec', 'mth19ec', 'mth20ec', 'mth21ec', 'mth22ec', 'mth23ec', 'mth24ec'))
                for i in range(1, 25):
                    hours = 'mth' + str(i) + 'ec'
                    for energy in masterdata:
                        try:
                            transdata.append(round(float(energy[hours]) * 0.041))
                        except:
                            transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Wind':
                transdata = []
                masterdata = list(Masterdatatable.objects.filter(mtdate = date, mtplntlctn = unitname, mtgrpname = 'Incomer', mtcategory = 'Secondary', mtsrcname = 'Transformer1').values('mth1ec', 'mth2ec', 'mth3ec', 'mth4ec', 'mth5ec', 'mth6ec', 'mth7ec', 'mth8ec', 'mth9ec', 'mth10ec', 'mth11ec', 'mth12ec', 'mth13ec', 'mth14ec', 'mth15ec', 'mth16ec', 'mth17ec', 'mth18ec', 'mth19ec', 'mth20ec', 'mth21ec', 'mth22ec', 'mth23ec', 'mth24ec'))
                for i in range(1, 25):
                    hours = 'mth' + str(i) + 'ec'
                    for energy in masterdata:
                        energydata = energy[hours]
                        wind = (float(energydata) * wind_percentage) / 100
                        try:
                            transdata.append(round(float(wind) * 0.59))
                        except:
                            transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

        carbomemission = {
            "label": label_array,
            "carbon": carbon
        }

    if request_data['type'] == 'month basis':

        month = int(request_data['month'])
        year = int(request_data['year'])
        start_date = datetime(year, month, 1)
        if month == 12:
            end_date = datetime(year, month, 31)
        else:
            end_date = (datetime(year, month + 1, 1) - timedelta(days = 1))
        label_array = []; carbon = []
        while start_date <= end_date:
            label_array.append(start_date.strftime("%Y-%m-%d"))
            start_date += timedelta(days=1)
        try:
            wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = year, Wind_month = calendar.month_name[month], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
        except:
            wind_percentage = 0
        for source in srcname:
            if source['assourcename'] == 'Transformer1':
                transdata = []
                for dates in label_array:
                    masterdata = Masterdatatable.objects.filter(mtdate = dates, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    transenergy = float(masterdata) - winddata
                    try:
                        transdata.append(round(float(transenergy) * 0.7132))
                    except:
                        transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Wind':
                transdata = []
                for dates in label_array:
                    masterdata = Masterdatatable.objects.filter(mtdate = dates, mtplntlctn = unitname, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    try:
                        transdata.append(round(float(winddata) * 0.59))
                    except:
                        transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Solar Energy':
                transdata = []
                for dates in label_array:
                    masterdata = Masterdatatable.objects.filter(mtdate = dates, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    try:
                        transdata.append(round(float(masterdata) * 0.041))
                    except:
                        transdata.append(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'DG':
                transdata = []
                for dates in label_array:
                    diesel = Cost_DG.objects.filter(dg_date = dates, dg_plantname = unitname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                    if diesel is None:
                        diesel = 0
                    try:
                        transdata.append(round(float(diesel) * 2.6))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Water':
                transdata = []
                for dates in label_array:
                    water = cost_water.objects.filter(wt_date = dates, wt_plantname = unitname, wt_type = 'Borewell').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if water is None:
                        water = 0
                    try:
                        transdata.append(round(float(water) * 0.59))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'LPG':
                transdata = []
                for dates in label_array:
                    lpg = Cost_LPG.objects.filter(LPG_date = dates, LPG_plantname = unitname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                    if lpg is None:
                        lpg = 0
                    try:
                        transdata.append(round(float(lpg) * 2.19))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'CNG':
                transdata = []
                for dates in label_array:
                    cng = Cost_CNG.objects.filter(CNG_date = dates, CNG_plantname = unitname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                    if cng is None:
                        cng = 0
                    try:
                        transdata.append(round(float(cng) * 0.614))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'PNG':
                transdata = []
                for dates in label_array:
                    png = Cost_PNG.objects.filter(PNG_date = dates, PNG_plantname = unitname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                    if png is None:
                        png = 0
                    try:
                        transdata.append(round(float(png) * 0.056))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Petrol':
                transdata = []
                for dates in label_array:
                    petrol = Cost_Petrol.objects.filter(pt_date = dates, pt_plantname = unitname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                    if petrol is None:
                        petrol = 0
                    try:
                        transdata.append(round(float(petrol) * 2.32))
                    except:
                        transdata(0)
                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

        carbomemission = {
            "label": label_array,
            "carbon": carbon
        }

    if request_data['type'] == 'year basis':
        year = int(request_data['year'])
        years = year + 1
        label_array = []; carbon = []

        for months in range(4, 13):
            monthname = calendar.month_name[months]
            label_array.append(str(monthname[:3]) +' '+ str(year))

        for months in range(1, 4):
            monthname = calendar.month_name[months]
            label_array.append(str(monthname[:3]) +' '+ str(years))

        for source in srcname:
            if source['assourcename'] == 'Transformer1':
                transdata = []
                for months in range(4, 13):
                    try:
                        wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = year, Wind_month = calendar.month_name[months], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
                    except:
                        wind_percentage = 0
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = year, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    energydata = float(masterdata) - winddata
                    try:
                        transdata.append(round(energydata * 0.7132))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    try:
                        wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = years, Wind_month = calendar.month_name[months], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
                    except:
                        wind_percentage = 0
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = years, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    energydata = float(masterdata) - winddata
                    try:
                        transdata.append(round(energydata * 0.7132))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Wind':
                transdata = []
                for months in range(4, 13):
                    try:
                        wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = year, Wind_month = calendar.month_name[months], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
                    except:
                        wind_percentage = 0
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = year, mtplntlctn = unitname, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    try:
                        transdata.append(round(winddata * 0.59))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    try:
                        wind_percentage = float(list(Cost_Wind.objects.filter(Wind_year = years, Wind_month = calendar.month_name[months], Wind_plantname = unitname).values('Wind_percentage'))[0]['Wind_percentage'])
                    except:
                        wind_percentage = 0
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = years, mtplntlctn = unitname, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    winddata = (float(masterdata) * wind_percentage) / 100
                    try:
                        transdata.append(round(winddata * 0.59))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Solar Energy':
                transdata = []
                for months in range(4, 13):
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = year, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    try:
                        transdata.append(round(float(masterdata) * 0.041))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    masterdata = Masterdatatable.objects.filter(mtdate__month = months, mtdate__year = years, mtplntlctn = unitname, mtsrcname = source['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))['mtenergycons__sum']
                    if masterdata is None:
                        masterdata = 0
                    try:
                        transdata.append(round(float(masterdata) * 0.041))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'DG':
                transdata = []
                for months in range(4, 13):
                    diesel = Cost_DG.objects.filter(dg_date__month = months, dg_date__year = year, dg_plantname = unitname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                    if diesel is None:
                        diesel = 0
                    try:
                        transdata.append(round(float(diesel) * 2.6))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    diesel = Cost_DG.objects.filter(dg_date__month = months, dg_date__year = years, dg_plantname = unitname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))['dg_lit_cons__sum']
                    if diesel is None:
                        diesel = 0
                    try:
                        transdata.append(round(float(diesel) * 2.6))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Petrol':
                transdata = []
                for months in range(4, 13):
                    petrol = Cost_Petrol.objects.filter(pt_date__month = months, pt_date__year = year, pt_plantname = unitname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                    if petrol is None:
                        petrol = 0
                    try:
                        transdata.append(round(float(petrol) * 2.32))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    petrol = Cost_Petrol.objects.filter(pt_date__month = months, pt_date__year = years, pt_plantname = unitname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))['pt_lit_cons__sum']
                    if petrol is None:
                        petrol = 0
                    try:
                        transdata.append(round(float(petrol) * 2.32))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'LPG':
                transdata = []
                for months in range(4, 13):
                    lpg = Cost_LPG.objects.filter(LPG_date__month = months, LPG_date__year = year, LPG_plantname = unitname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                    if lpg is None:
                        lpg = 0
                    try:
                        transdata.append(round(float(lpg) * 2.19))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    lpg = Cost_LPG.objects.filter(LPG_date__month = months, LPG_date__year = years, LPG_plantname = unitname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))['LPG_kg_cons__sum']
                    if lpg is None:
                        lpg = 0
                    try:
                        transdata.append(round(float(lpg) * 2.19))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'CNG':
                transdata = []
                for months in range(4, 13):
                    cng = Cost_CNG.objects.filter(CNG_date__month = months, CNG_date__year = year, CNG_plantname = unitname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                    if cng is None:
                        cng = 0
                    try:
                        transdata.append(round(float(cng) * 0.614))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    cng = Cost_CNG.objects.filter(CNG_date__month = months, CNG_date__year = years, CNG_plantname = unitname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))['CNG_kg_cons__sum']
                    if cng is None:
                        cng = 0
                    try:
                        transdata.append(round(float(cng) * 0.614))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'PNG':
                transdata = []
                for months in range(4, 13):
                    png = Cost_PNG.objects.filter(PNG_date__month = months, PNG_date__year = year, PNG_plantname = unitname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))['PNG_kg_cons__sum']
                    if png is None:
                        png = 0
                    try:
                        transdata.append(round(float(png) * 0.056))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    png = Cost_PNG.objects.filter(dg_date__month = months, dg_date__year = years, dg_plantname = unitname).values('PNG_kg_cons').aggregate(Sum('NG_kg_cons'))['PNG_kg_cons__sum']
                    if png is None:
                        png = 0
                    try:
                        transdata.append(round(float(png) * 0.056))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

            if source['assourcename'] == 'Water':
                transdata = []
                for months in range(4, 13):
                    water = cost_water.objects.filter(wt_date__month = months, wt_date__year = year, wt_plantname = unitname, wt_type = 'Borewell').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if water is None:
                        water = 0
                    try:
                        transdata.append(round(float(water) * 0.59))
                    except:
                        transdata.append(0)

                for months in range(1, 4):
                    water = cost_water.objects.filter(wt_date__month = months, wt_date__year = years, wt_plantname = unitname, wt_type = 'Borewell').values('wt_consume').aggregate(Sum('wt_consume'))['wt_consume__sum']
                    if water is None:
                        water = 0
                    try:
                        transdata.append(round(float(water) * 0.59))
                    except:
                        transdata.append(0)

                my_dict = {
                    "name": source['assourcename'],
                    "data": transdata
                }
                carbon.append(my_dict)

        carbomemission = {
            "label": label_array,
            "carbon": carbon
        }

    return JsonResponse(carbomemission, safe=False)