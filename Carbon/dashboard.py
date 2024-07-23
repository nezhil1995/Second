from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt
from encrypt.renderers import CustomAesRenderer
from datetime import datetime
from costestimator.models import Cost_CNG, Cost_LPG, Cost_DG, Cost_PNG, Cost_Petrol
from django.db.models.aggregates import Sum
from meter_data.models import Masterdatatable
from source_management.models import AddSource
from rest_framework.decorators import authentication_classes, permission_classes, api_view,renderer_classes
from encrypt.renderers import CustomAesRenderer
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import IsAuthenticated

# @renderer_classes([CustomAesRenderer])
@api_view(['GET', 'POST', 'PUT', 'DELETE'])
@authentication_classes([TokenAuthentication])
@permission_classes([IsAuthenticated])
# @csrf_exempt
def Dashboard(request):
    if request.method == 'GET':
        
        plantname = request.GET['plantname']
        
        current_date = datetime.now()
        current_month = current_date.month
        current_year = current_date.year
        
        # FUEL
        fuel_array = []
        try:
            dg_data = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = current_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
            fuel_array.append(float(dg_data['dg_lit_cons__sum']) * 2.6)
        except:
            fuel_array.append(0)
        try:
            pt_data = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = current_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
            fuel_array.append(float(pt_data['pt_lit_cons__sum']) * 2.32)
        except:
            fuel_array.append(0)
        try: 
            lpg_data = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = current_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
            fuel_array.append(float(lpg_data['LPG_kg_cons__sum']) * 2.19)
        except:
            fuel_array.append(0)
        try:
            cng_data = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = current_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
            fuel_array.append(float(cng_data['CNG_kg_cons__sum']) * 0.614)
        except:
            fuel_array.append(0)
        try:
            png_data = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = current_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
            fuel_array.append(float(png_data['PNG_kg_cons__sum']) * 0.056)
        except:
            fuel_array.append(0)
        
        # ELECTRICITY
        electricity_array = []
        sourcename = AddSource.objects.filter(asplantname = plantname).values('assourcename')
        for data in sourcename:
            if data['assourcename'] == 'Transformer1':
                transformer = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    electricity_array.append(float(transformer['mtenergycons__sum']) * 0.7132)
                except:
                    electricity_array.append(0)
            if data['assourcename'] == 'Solar Energy':
                solar = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = data['assourcename'], mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                try:
                    electricity_array.append(float(solar['mtenergycons__sum']) * 0.041)
                except:
                    electricity_array.append(0)
        
        # Line graph
        # Scope1
        emission1_array = []
        start_month = 4; end_month = 12
        if current_month >= start_month:
            while start_month <= end_month:
                monthwise_emission = []
                try:
                    dg_data = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    monthwise_emission.append(float(dg_data['dg_lit_cons__sum']) * 2.6)
                except:
                    monthwise_emission.append(0)
                try:
                    pt_data = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    monthwise_emission.append(float(pt_data['pt_lit_cons__sum']) * 2.32)
                except:
                    monthwise_emission.append(0)
                try:
                    lpg_data = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    monthwise_emission.append(float(lpg_data['LPG_kg_cons__sum']) * 2.19)
                except:
                    monthwise_emission.append(0)
                try:
                    cng_data = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    monthwise_emission.append(float(cng_data['CNG_kg_cons__sum']) * 0.614)
                except:
                    monthwise_emission.append(0)
                try:
                    png_data = Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    monthwise_emission.append(float(png_data['PNG_kg_cons__sum']) * 0.056)
                except:
                    monthwise_emission.append(0)
                emission1_array.append(round(sum(monthwise_emission)))
                start_month += 1
            
            start_month = 1
            while start_month < 4:
                monthwise_emission = []
                try:
                    dg_data = Cost_DG.objects.filter(dg_date__year = current_year + 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    monthwise_emission.append(float(dg_data['dg_lit_cons__sum']) * 2.6)
                except:
                    monthwise_emission.append(0)
                try:
                    pt_data = Cost_Petrol.objects.filter(pt_date__year = current_year + 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    monthwise_emission.append(float(pt_data['pt_lit_cons__sum']) * 2.32)
                except:
                    monthwise_emission.append(0)
                try:
                    lpg_data = Cost_LPG.objects.filter(LPG_date__year = current_year + 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    monthwise_emission.append(float(lpg_data['LPG_kg_cons__sum']) * 2.19)
                except:
                    monthwise_emission.append(0)
                try:
                    cng_data = Cost_CNG.objects.filter(CNG_date__year = current_year + 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    monthwise_emission.append(float(cng_data['CNG_kg_cons__sum']) * 0.614)
                except:
                    monthwise_emission.append(0)
                try:
                    png_data = Cost_PNG.objects.filter(PNG_date__year = current_year + 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    monthwise_emission.append(float(png_data['PNG_kg_cons__sum']) * 0.056)
                except:
                    monthwise_emission.append(0)
                emission1_array.append(round(sum(monthwise_emission)))
                start_month += 1
            
        else:
            while start_month <= end_month:
                monthwise_emission = []
                try:
                    dg_data = Cost_DG.objects.filter(dg_date__year = current_year - 1, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    monthwise_emission.append(float(dg_data['dg_lit_cons__sum']) * 2.6)
                except:
                    monthwise_emission.append(0)
                try:
                    pt_data = Cost_Petrol.objects.filter(pt_date__year = current_year - 1, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    monthwise_emission.append(float(pt_data['pt_lit_cons__sum']) * 2.32)
                except:
                    monthwise_emission.append(0)
                try:
                    lpg_data = Cost_LPG.objects.filter(LPG_date__year = current_year - 1, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    monthwise_emission.append(float(lpg_data['LPG_kg_cons__sum']) * 2.19)
                except:
                    monthwise_emission.append(0)
                try:
                    cng_data = Cost_CNG.objects.filter(CNG_date__year = current_year - 1, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    monthwise_emission.append(float(cng_data['CNG_kg_cons__sum']) * 0.614)
                except:
                    monthwise_emission.append(0)
                try:
                    png_data = Cost_PNG.objects.filter(PNG_date__year = current_year - 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    monthwise_emission.append(float(png_data['PNG_kg_cons__sum']) * 0.056)
                except:
                    monthwise_emission.append(0)
                emission1_array.append(round(sum(monthwise_emission)))
                
                start_month += 1
            start_month = 1
            while start_month < 4:
                try:
                    dg_data = Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = start_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons'))
                    monthwise_emission.append(float(dg_data['dg_lit_cons__sum']) * 2.6)
                except:
                    monthwise_emission.append(0)
                try:
                    pt_data = Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = start_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons'))
                    monthwise_emission.append(float(pt_data['pt_lit_cons__sum']) * 2.32)
                except:
                    monthwise_emission.append(0)
                try:
                    lpg_data = Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = start_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons'))
                    monthwise_emission.append(float(lpg_data['LPG_kg_cons__sum']) * 2.19)
                except:
                    monthwise_emission.append(0)
                try:
                    cng_data = Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = start_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons'))
                    monthwise_emission.append(float(cng_data['CNG_kg_cons__sum']) * 0.614)
                except:
                    monthwise_emission.append(0)
                try:
                    png_data = Cost_PNG.objects.filter(PNG_date__year = current_year - 1, PNG_date__month = start_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons'))
                    monthwise_emission.append(float(png_data['PNG_kg_cons__sum']) * 0.056)
                except:
                    monthwise_emission.append(0)
                emission1_array.append(round(sum(monthwise_emission)))
                
                start_month += 1

        scope2_monthly_emission = []
        start_month = 4; end_month = 12
        if current_month >= start_month:
            while start_month <= end_month:
                array1 = []
                try:
                    trans_emission = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array1.append(float(trans_emission['mtenergycons__sum']) * 0.7132)
                except:
                    array1.append(0)
                try:
                    solar_emission = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array1.append(float(solar_emission['mtenergycons__sum']) * 0.041)
                except:
                    array1.append(0)
                scope2_monthly_emission.append(round(sum(array1)))
                start_month += 1
            start_month = 1
            while start_month < 4:
                array2 = []
                try:
                    trans_emission = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array2.append(float(trans_emission['mtenergycons__sum']) * 0.7132)
                except:
                    array2.append(0)
                try:
                    solar_emission = Masterdatatable.objects.filter(mtdate__year = current_year + 1, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array2.append(float(solar_emission['mtenergycons__sum']) * 0.041)
                except:
                    array2.append(0)
                scope2_monthly_emission.append(round(sum(array2)))
                start_month += 1
        else:
            while start_month <= end_month:
                array1 = []
                try:
                    trans_emission = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array1.append(float(trans_emission['mtenergycons__sum']) * 0.7132)
                except:
                    array1.append(0)
                try:
                    solar_emission = Masterdatatable.objects.filter(mtdate__year = current_year - 1, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array1.append(float(solar_emission['mtenergycons__sum']) * 0.041)
                except:
                    array1.append(0)
                scope2_monthly_emission.append(round(sum(array1)))
                start_month += 1
            start_month = 1
            while start_month < 4:
                array2 = []
                try:
                    trans_emission = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array2.append(float(trans_emission['mtenergycons__sum']) * 0.7132)
                except:
                    array2.append(0)
                try:
                    solar_emission = Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = start_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons'))
                    array2.append(float(solar_emission['mtenergycons__sum']) * 0.041)
                except:
                    array2.append(0)
                scope2_monthly_emission.append(round(sum(array2)))
                start_month += 1
        
        # Scope3
        scope3_monthly_emission = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        
        # Donut yearly chart
        scope1 = sum(emission1_array)
        scope2 = sum(scope2_monthly_emission)
        scope3 = 0
        
        # PieChart current month based
        # Diesel
        diesel = ( Cost_DG.objects.filter(dg_date__year = current_year, dg_date__month = current_month, dg_plantname = plantname).values('dg_lit_cons').aggregate(Sum('dg_lit_cons')))['dg_lit_cons__sum']
        if diesel is None:
            diesel = 0
        
        # Petrol
        petrol = (Cost_Petrol.objects.filter(pt_date__year = current_year, pt_date__month = current_month, pt_plantname = plantname).values('pt_lit_cons').aggregate(Sum('pt_lit_cons')))['pt_lit_cons__sum']
        if petrol is None:
            petrol = 0
        
        # LPG
        lpg =( Cost_LPG.objects.filter(LPG_date__year = current_year, LPG_date__month = current_month, LPG_plantname = plantname).values('LPG_kg_cons').aggregate(Sum('LPG_kg_cons')))['LPG_kg_cons__sum']
        if lpg is None:
            lpg = 0
        
        # CNG 
        cng = ( Cost_CNG.objects.filter(CNG_date__year = current_year, CNG_date__month = current_month, CNG_plantname = plantname).values('CNG_kg_cons').aggregate(Sum('CNG_kg_cons')))['CNG_kg_cons__sum']
        if cng is None:
            cng = 0
        
        # PNG 
        png =( Cost_PNG.objects.filter(PNG_date__year = current_year, PNG_date__month = current_month, PNG_plantname = plantname).values('PNG_kg_cons').aggregate(Sum('PNG_kg_cons')))['PNG_kg_cons__sum']
        if png is None:
            png =0
        
        # Transformer1
        
        transformer1 =( Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = 'Transformer1', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
        transformer1 = float(transformer1) * 0.7132
        if transformer1 is None:
            transformer1 = 0
        
        # Solar Energy
        solarenergy =( Masterdatatable.objects.filter(mtdate__year = current_year, mtdate__month = current_month, mtsrcname = 'Solar Energy', mtcategory = 'Secondary', mtgrpname = 'Incomer', mtplntlctn = plantname).values('mtenergycons').aggregate(Sum('mtenergycons')))['mtenergycons__sum']
        if solarenergy is None:
            solarenergy = 0
        
        # Final
        
        total_energy_emission = round(sum(electricity_array))
        total_fuel_emission = round(sum(fuel_array))
        waste = 0
        scope_data = [scope1, scope2, scope3]
        scope_name = ['S1', 'S2', 'S3']
        emission_data = [round(transformer1), round(float(solarenergy) * 0.041), round(float(diesel) * 2.6), round(float(petrol) * 2.32), round(float(lpg) * 2.19), round(float(cng) * 0.614), round(float(png) * 0.056)]
        emission_name = ['Transformer', 'Solar Energy', 'Diesel', 'Petrol', 'LPG', 'CNG', 'PNG']
        monthly_emission_name = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        monthly_emission_s1 = emission1_array
        monthly_emission_s2 = scope2_monthly_emission
        monthly_emission_s3 = scope3_monthly_emission
        
        mydict = {
            'Energy_Emission' : total_energy_emission,
            'Fuel_Emission' : total_fuel_emission,
            'Waste': waste,
            'Scope_data': scope_data,
            'Scope_name' : scope_name,
            'Emission_name' : emission_name,
            'Emission_data' : emission_data,
            'Monthly_emission_name' : monthly_emission_name,
            'Monthly_emission_S1' : monthly_emission_s1,
            'Monthly_emission_S2' : monthly_emission_s2,
            'Monthly_emission_S3' : monthly_emission_s3
        }
        return JsonResponse(mydict, safe=False)