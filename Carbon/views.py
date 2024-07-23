import json
import os
from click import style
from django.http import JsonResponse, HttpResponse
from django.shortcuts import render
from REnergy.settings import BASE_DIR
from source_management.models import AddSource
from meter_data.models import Masterdatatable
from costestimator.models import Cost_DG, Cost_Petrol, Cost_CNG, Cost_LPG, Cost_PNG
from django.db.models.aggregates import Sum
from django.views.decorators.csrf import csrf_exempt
from encrypt.renderers import CustomAesRenderer
from datetime import datetime, timedelta
from rest_framework.decorators import authentication_classes, permission_classes, api_view,renderer_classes
from encrypt.renderers import CustomAesRenderer
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import IsAuthenticated

# @renderer_classes([CustomAesRenderer])
@api_view(['GET', 'POST', 'PUT', 'DELETE'])
@authentication_classes([TokenAuthentication])
@permission_classes([IsAuthenticated])
# @csrf_exempt
def CarbonSummary(request):
    
    if request.method == 'POST':
        plantname = request.GET['plantname']
        request_data = json.loads(request.body)
        type = request_data['Type']
        month_data = request_data['Month']
        year_data = request_data['Year']
        
        if type == 'Month Basis':
            
            monthbasis = []
            
            transformer = Masterdatatable.objects.filter(mtdate__year = year_data, mtdate__month = month_data, mtplntlctn = plantname, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))
            try:
                transformer_cons = round(transformer['mtenergycons__sum'])
            except:
                transformer_cons = 0
            try:
                transformer_emis = round(float(transformer['mtenergycons__sum']) * 0.7132)
            except:
                transformer_emis = 0
            transformer_data = {
                "description": "Purchased electricity from Electricity Authorities(non-renewable)",
                "units": "kWh",
                "consumption": transformer_cons,
                "ghg": "CO2",
                "Namghg": 0.7132,
                "unitcon": "Kg CO2/kWh",
                "emission": transformer_emis,
                "id": 1
            }
            monthbasis.append(transformer_data)
            
            wind_data = {
                "description": "Purchased electricity from Third Party(wind-renewable)",
                "units": "kWh",
                "consumption": 0,
                "ghg": "CO2 equ",
                "Namghg": 0.59,
                "unitcon": "Kg CO2/kL",
                "emission": 0,
                "id": 2
            }
            monthbasis.append(wind_data)
            
            solar = Masterdatatable.objects.filter(mtdate__year = year_data, mtdate__month = month_data, mtplntlctn = plantname, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer').values('mtenergycons').aggregate(Sum('mtenergycons'))
            try:
                solar_cons = round(solar['mtenergycons__sum'])
            except:
                solar_cons = 0
            try:
                solar_emis = round(float(solar['mtenergycons__sum']) * 0.041)
            except:
                solar_emis = 0
            solar_data = {
                "description": "Purchased electricity from Second Party(solar-renewable)",
                "units": "kWh",
                "consumption": solar_cons,
                "ghg": "CO2",
                "Namghg": 0.041,
                "unitcon": "Kg CO2/kWh",
                "emission": solar_emis,
                "id": 3
            }
            monthbasis.append(solar_data)
            
            diesel = Cost_DG.objects.filter(dg_date__year = year_data, dg_date__month = month_data, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
            try:
                diesel_cons = round(diesel['dg_lit_cons__sum'])
            except:
                diesel_cons = 0
            try:
                diesel_emis = round(float(diesel['dg_lit_cons__sum']) * 2.6)
            except:
                diesel_emis = 0
            diesel_data = {
                "description": "Diesel",
                "units": "Litres",
                "consumption": diesel_cons,
                "ghg": "CO2",
                "Namghg": 2.6,
                "unitcon": "Kg CO2/L",
                "emission": diesel_emis,
                "id": 4
            }
            monthbasis.append(diesel_data)
            
            petrol = Cost_Petrol.objects.filter(pt_date__year = year_data, pt_date__month = month_data, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
            try:
                petrol_cons = round(petrol['pt_lit_cons__sum'])
            except:
                petrol_cons = 0
            try:
                petrol_emis = round(float(petrol['pt_lit_cons__sum']) * 2.32)
            except:
                petrol_emis = 0
            petrol_data = {
                "description": "Petrol",
                "units": "Litres",
                "consumption": petrol_cons,
                "ghg": "CO2",
                "Namghg": 2.32,
                "unitcon": "Kg CO2/L",
                "emission": petrol_emis,
                "id": 5
            }
            monthbasis.append(petrol_data)
            
            lpg = Cost_LPG.objects.filter(LPG_date__year = year_data, LPG_date__month = month_data, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
            try:
                lpg_cons = round(lpg['LPG_kg_cons__sum'])
            except:
                lpg_cons = 0
            try:
                lpg_emis = round(float(lpg['LPG_kg_cons__sum']) * 2.19)
            except:
                lpg_emis = 0
            lpg_data = {
                "description": "LPG",
                "units": "Litres",
                "consumption": lpg_cons,
                "ghg": "CO2",
                "Namghg": 2.19,
                "unitcon": "Kg CO2/L",
                "emission": lpg_emis,
                "id": 6
            }
            monthbasis.append(lpg_data)
            
            cng = Cost_CNG.objects.filter(CNG_date__year = year_data, CNG_date__month = month_data, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
            try:
                cng_cons = round(cng['CNG_kg_cons__sum'])
            except:
                cng_cons = 0
            try:
                cng_emis = round(float(cng['CNG_kg_cons__sum']) * 0.614)
            except:
                cng_emis = 0
            cng_data = {
                "description": "CNG",
                "units": "Kg",
                "consumption": cng_cons,
                "ghg": "CO2",
                "Namghg": 0.614,
                "unitcon": "Kg CO2/Kg",
                "emission": cng_emis,
                "id": 7
            }
            monthbasis.append(cng_data)
            
            png = Cost_PNG.objects.filter(PNG_date__year = year_data, PNG_date__month = month_data, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
            try:
                png_cons = round(png['PNG_kg_cons__sum'])
            except:
                png_cons = 0
            try:
                png_emis = round(float(png['PNG_kg_cons__sum']) * 0.056)
            except:
                png_emis = 0
            png_data = {
                "description": "PNG",
                "units": "Kg",
                "consumption": png_cons,
                "ghg": "CO2",
                "Namghg": 0.056,
                "unitcon": "Kg CO2/Kg",
                "emission": png_emis,
                "id": 8
            }
            monthbasis.append(png_data)
            
            return JsonResponse(monthbasis, safe=False)
            
        if type == 'Year Basis':
            
            yearbasis = []
            
            transformer_consumption = []
            transformer_emission = []
            for month in range(4, 13):
                transformer = Masterdatatable.objects.filter(mtdate__year = year_data, mtdate__month = month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    transformer_consumption.append(round(transformer['mtenergycons__sum']))
                except:
                    transformer_consumption.append(0)
                try:
                    transformer_emission.append(round(float(transformer['mtenergycons__sum']) * 0.7132))
                except:
                    transformer_emission.append(0)
                
            for month in range(1, 4):
                transformer = Masterdatatable.objects.filter(mtdate__year = int(year_data) + 1, mtdate__month = month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    transformer_consumption.append(round(transformer['mtenergycons__sum']))
                except:
                    transformer_consumption.append(0)
                try:
                    transformer_emission.append(round(float(transformer['mtenergycons__sum']) * 0.7132))
                except:
                    transformer_emission.append(0)
                    
            transformer_data = {
                "description": "Purchased electricity from Electricity Authorities(non-renewable)",
                "units": "kWh",
                "consumption": sum(transformer_consumption),
                "ghg": "CO2",
                "Namghg": 0.7132,
                "unitcon": "Kg CO2/kWh",
                "emission": sum(transformer_emission),
                "id": 1
            }
            yearbasis.append(transformer_data)
            
            wind_data = {
                "description": "Purchased electricity from Third Party(wind-renewable)",
                "units": "kWh",
                "consumption": 0,
                "ghg": "CO2 equ",
                "Namghg": 0.59,
                "unitcon": "Kg CO2/kL",
                "emission": 0,
                "id": 2
            }
            yearbasis.append(wind_data)
            
            solar_consumption = []
            solar_emission = []
            for month in range(4, 13):
                solar = Masterdatatable.objects.filter(mtdate__year = year_data, mtdate__month = month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    solar_consumption.append(round(solar['mtenergycons__sum']))
                except:
                    solar_consumption.append(0)
                try:
                    solar_emission.append(round(float(solar['mtenergycons__sum']) * 0.041))
                except:
                    solar_emission.append(0)
                
            for month in range(1, 4):
                solar = Masterdatatable.objects.filter(mtdate__year = int(year_data) + 1, mtdate__month = month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    solar_consumption.append(round(solar['mtenergycons__sum']))
                except:
                    solar_consumption.append(0)
                try:
                    solar_emission.append(round(float(solar['mtenergycons__sum']) * 0.041))
                except:
                    solar_emission.append(0)
                    
            solar_data = {
                "description": "Purchased electricity from Second Party(solar-renewable)",
                "units": "kWh",
                "consumption": sum(solar_consumption),
                "ghg": "CO2",
                "Namghg": 0.041,
                "unitcon": "Kg CO2/kWh",
                "emission": sum(solar_emission),
                "id": 3
            }
            yearbasis.append(solar_data)
            
            diesel_consumption = []
            diesel_emission = []
            for month in range(4, 13):
                diesel = Cost_DG.objects.filter(dg_date__year = year_data, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                try:
                    diesel_consumption.append(round(diesel['dg_lit_cons__sum']))
                except:
                    diesel_consumption.append(0)
                try:
                    diesel_emission.append(round(float(diesel['dg_lit_cons__sum']) * 2.6))
                except:
                    diesel_emission.append(0)
                
            for month in range(1, 4):
                diesel = Cost_DG.objects.filter(dg_date__year = int(year_data) + 1, dg_date__month = month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                try:
                    diesel_consumption.append(round(diesel['dg_lit_cons__sum']))
                except:
                    diesel_consumption.append(0)
                try:
                    diesel_emission.append(round(float(diesel['dg_lit_cons__sum']) * 2.6))
                except:
                    diesel_emission.append(0)
                    
            diesel_data = {
                "description": "Diesel",
                "units": "Litres",
                "consumption": sum(diesel_consumption),
                "ghg": "CO2",
                "Namghg": 2.6,
                "unitcon": "Kg CO2/L",
                "emission": sum(diesel_emission),
                "id": 4
            }
            yearbasis.append(diesel_data)
            
            petrol_consumption = []
            petrol_emission = []
            for month in range(4, 13):
                petrol = Cost_Petrol.objects.filter(pt_date__year = year_data, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                try:
                    petrol_consumption.append(round(petrol['pt_lit_cons__sum']))
                except:
                    petrol_consumption.append(0)
                try:
                    petrol_emission.append(round(float(petrol['pt_lit_cons__sum']) * 2.32))
                except:
                    petrol_emission.append(0)
                
            for month in range(1, 4):
                petrol = Cost_Petrol.objects.filter(pt_date__year = int(year_data) + 1, pt_date__month = month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                try:
                    petrol_consumption.append(round(petrol['pt_lit_cons__sum']))
                except:
                    petrol_consumption.append(0)
                try:
                    petrol_emission.append(round(float(petrol['pt_lit_cons__sum']) * 2.32))
                except:
                    petrol_emission.append(0)
                    
            petrol_data = {
                "description": "Petrol",
                "units": "Litres",
                "consumption": sum(petrol_consumption),
                "ghg": "CO2",
                "Namghg": 2.32,
                "unitcon": "Kg CO2/L",
                "emission": sum(petrol_emission),
                "id": 5
            }
            yearbasis.append(petrol_data)
            
            lpg_consumption = []
            lpg_emission = []
            for month in range(4, 13):
                lpg = Cost_LPG.objects.filter(LPG_date__year = year_data, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                try:
                    lpg_consumption.append(round(lpg['LPG_kg_cons__sum']))
                except:
                    lpg_consumption.append(0)
                try:
                    lpg_emission.append(round(float(lpg['LPG_kg_cons__sum']) * 2.19))
                except:
                    lpg_emission.append(0)
                
            for month in range(1, 4):
                lpg = Cost_LPG.objects.filter(LPG_date__year = int(year_data) + 1, LPG_date__month = month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                try:
                    lpg_consumption.append(round(lpg['LPG_kg_cons__sum']))
                except:
                    lpg_consumption.append(0)
                try:
                    lpg_emission.append(round(float(lpg['LPG_kg_cons__sum']) * 2.19))
                except:
                    lpg_emission.append(0)
                    
            lpg_data = {
                "description": "LPG",
                "units": "Litres",
                "consumption": sum(lpg_consumption),
                "ghg": "CO2",
                "Namghg": 2.19,
                "unitcon": "Kg CO2/L",
                "emission": sum(lpg_emission),
                "id": 6
            }
            yearbasis.append(lpg_data)
            
            cng_consumption = []
            cng_emission = []
            for month in range(4, 13):
                cng = Cost_CNG.objects.filter(CNG_date__year = year_data, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                try:
                    cng_consumption.append(round(cng['CNG_kg_cons__sum']))
                except:
                    cng_consumption.append(0)
                try:
                    cng_emission.append(round(float(cng['CNG_kg_cons__sum']) * 0.614))
                except:
                    cng_emission.append(0)
                
            for month in range(1, 4):
                cng = Cost_CNG.objects.filter(CNG_date__year = int(year_data) + 1, CNG_date__month = month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                try:
                    cng_consumption.append(round(cng['CNG_kg_cons__sum']))
                except:
                    cng_consumption.append(0)
                try:
                    cng_emission.append(round(float(cng['CNG_kg_cons__sum']) * 0.614))
                except:
                    cng_emission.append(0)
                    
            cng_data = {
                "description": "CNG",
                "units": "Kg",
                "consumption": sum(cng_consumption),
                "ghg": "CO2",
                "Namghg": 0.614,
                "unitcon": "Kg CO2/Kg",
                "emission": sum(cng_emission),
                "id": 7
            }
            yearbasis.append(cng_data)
            
            png_consumption = []
            png_emission = []
            for month in range(4, 13):
                png = Cost_PNG.objects.filter(PNG_date__year = year_data, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                try:
                    png_consumption.append(round(png['PNG_kg_cons__sum']))
                except:
                    png_consumption.append(0)
                try:
                    png_emission.append(round(float(png['PNG_kg_cons__sum']) * 0.056))
                except:
                    png_emission.append(0)
                
            for month in range(1, 4):
                png = Cost_PNG.objects.filter(PNG_date__year = int(year_data) + 1, PNG_date__month = month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                try:
                    png_consumption.append(round(png['PNG_kg_cons__sum']))
                except:
                    png_consumption.append(0)
                try:
                    png_emission.append(round(float(png['PNG_kg_cons__sum']) * 0.056))
                except:
                    png_emission.append(0)
                    
            png_data = {
                "description": "PNG",
                "units": "Kg",
                "consumption": sum(png_consumption),
                "ghg": "CO2",
                "Namghg": 0.056,
                "unitcon": "Kg CO2/Kg",
                "emission": sum(png_emission),
                "id": 8
            }
            yearbasis.append(png_data)

            return JsonResponse(yearbasis, safe=False)