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
def YearlyEmission(request):
    if request.method == 'POST':
        plantname = request.GET['plantname']
        request_data = json.loads(request.body)
        source = request_data['Source']
        year = request_data['Year']
        
        yearly_emission = []
        
        if source == 'Transformer':
            for month in range(4, 13):
                transformer = Masterdatatable.objects.filter(mtdate__year = year, mtdate__month = month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    yearly_emission.append(round(float(transformer['mtenergycons__sum']) * 0.7132))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                transformer = Masterdatatable.objects.filter(mtdate__year = int(year) + 1, mtdate__month = month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    yearly_emission.append(round(float(transformer['mtenergycons__sum']) * 0.7132))
                except:
                    yearly_emission.append(0)
        
        if source == 'Solar Energy':
            for month in range(4, 13):
                solar = Masterdatatable.objects.filter(mtdate__year = year, mtdate__month = month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    yearly_emission.append(round(float(solar['mtenergycons__sum']) * 0.041))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                solar = Masterdatatable.objects.filter(mtdate__year = int(year) + 1, mtdate__month = month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    yearly_emission.append(round(float(solar['mtenergycons__sum']) * 0.041))
                except:
                    yearly_emission.append(0)
        
        if source == 'Diesel':
            for month in range(4, 13):
                diesel = Cost_DG.objects.filter(dg_date__year = year, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                try:
                    yearly_emission.append(round(float(diesel['dg_lit_cons__sum']) * 2.6))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                diesel = Cost_DG.objects.filter(dg_date__year = int(year) + 1, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                try:
                    yearly_emission.append(round(float(diesel['dg_lit_cons__sum']) * 2.6))
                except:
                    yearly_emission.append(0)
                    
        if source == 'Petrol':
            for month in range(4, 13):
                petrol = Cost_Petrol.objects.filter(pt_date__year = year, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                try:
                    yearly_emission.append(round(float(petrol['pt_lit_cons__sum']) * 2.32))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                petrol = Cost_Petrol.objects.filter(pt_date__year = int(year) + 1, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                try:
                    yearly_emission.append(round(float(petrol['pt_lit_cons__sum']) * 2.32))
                except:
                    yearly_emission.append(0)
                    
        if source == 'LPG':
            for month in range(4, 13):
                lpg = Cost_LPG.objects.filter(LPG_date__year = year, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                try:
                    yearly_emission.append(round(float(lpg['LPG_kg_cons__sum']) * 2.19))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                lpg = Cost_LPG.objects.filter(LPG_date__year = int(year) + 1, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                try:
                    yearly_emission.append(round(float(lpg['LPG_kg_cons__sum']) * 2.19))
                except:
                    yearly_emission.append(0)
                    
        if source == 'CNG':
            for month in range(4, 13):
                cng = Cost_CNG.objects.filter(CNG_date__year = year, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                try:
                    yearly_emission.append(round(float(cng['CNG_kg_cons__sum']) * 0.614))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                cng = Cost_CNG.objects.filter(CNG_date__year = int(year) + 1, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                try:
                    yearly_emission.append(round(float(cng['CNG_kg_cons__sum']) * 0.614))
                except:
                    yearly_emission.append(0)
                    
        if source == 'PNG':
            for month in range(4, 13):
                png = Cost_PNG.objects.filter(PNG_date__year = year, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                try:
                    yearly_emission.append(round(float(png['PNG_kg_cons__sum']) * 0.614))
                except:
                    yearly_emission.append(0)
            for month in range(1, 4):
                png = Cost_PNG.objects.filter(PNG_date__year = int(year) + 1, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                try:
                    yearly_emission.append(round(float(png['PNG_kg_cons__sum']) * 0.614))
                except:
                    yearly_emission.append(0)
                    
        monthname = []
        for month in range(4, 13):
            months = datetime(int(year), month, 1).strftime("%b %y")
            monthname.append(months)
        for month in range(1, 4):
            months = datetime(int(year) + 1, month, 1).strftime("%b %y")
            monthname.append(months)
                    
        my_dict = {
            'Values' : yearly_emission,
            "Month" : monthname
        }
                    
        
    return JsonResponse(my_dict, safe=False)