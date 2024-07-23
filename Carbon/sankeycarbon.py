from django.http import JsonResponse
from meter_data.models import Masterdatatable
from source_management.models import AddSource
from datetime import datetime , date, timedelta
from group_management.models import AddGroup
from meter_management.models import AddMeter
from django.views.decorators.csrf import csrf_exempt
from encrypt.renderers import CustomAesRenderer
from workflow_dashboard.models import WorkflowData
from django.db.models import Sum as add
import json
from rest_framework.decorators import authentication_classes, permission_classes, api_view,renderer_classes
from encrypt.renderers import CustomAesRenderer
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import IsAuthenticated

# @renderer_classes([CustomAesRenderer])
@api_view(['GET', 'POST', 'PUT', 'DELETE'])
@authentication_classes([TokenAuthentication])
@permission_classes([IsAuthenticated])
# @csrf_exempt
def SankeyCarbon(request):
    
    if request.method == 'POST':
        request_data = json.loads(request.body)
        request_type = request_data['Type']

        if request_type == 'Date Basis':
            Plantname = request.GET['plantname']   
            current_date = request_data['Date']  
            link = []; nodes = []; Solar_Consumption = 0; total_grpValue = 0; DG_Consumption = 0
            nodes.append({"name": "Total Emission"})
            sourcenames = AddSource.objects.filter(asplantname=Plantname).distinct('assourcename').values('assourcename')
            groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            meternames = AddMeter.objects.filter(amplantname = Plantname).distinct('ammetername').values('ammetername')
            node_groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            #Total to Incomer
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':
                    node_srcnm = 'Transformer'
                    total_energy = Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    if total_energy is not None:
                        nodes.append({"name": node_srcnm})
                        link.append({"source":"Total Emission",
                                    "target":node_srcnm,
                                    "value":float(total_energy) * 0.7132})
                else:
                    if(srcname['assourcename'] == 'Solar Energy'):
                        try:
                            Solar_Consumption = round(float(Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":Solar_Consumption*0.041
                                        })
                        except:
                            Solar_Consumption = 0
                    elif(srcname['assourcename'] == 'DG'):
                        try:
                            DG_Consumption = round(float(Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":DG_Consumption*2.6
                                        })
                        except:
                            DG_Consumption = 0
                    else:
                        nodes.append({"name": srcname['assourcename']})
                        link.append({"source":"Total Emission",
                                    "target":srcname['assourcename'],
                                    "value":Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']})
            #Append Group Name to nodes
            for data in node_groupnames:
                if data['aggroupname'] != 'Incomer':
                    nodes.append({"name":data['aggroupname']})
                    grp_val = Masterdatatable.objects.filter(mtdate = current_date, mtgrpname = data['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    if grp_val is None:
                        total_grpValue = 0
                    else:
                        total_grpValue += grp_val
            #Individual Group Percent
            grpPercent = []
            for grupnme in groupnames:
                if grupnme['aggroupname'] != 'Incomer':
                    try:
                        grpVal = Masterdatatable.objects.filter(mtdate = current_date, mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    except:
                        grpVal = 0
                    grpPercent.append({'Groupname':grupnme['aggroupname'], 
                                       'percent': round((float(grpVal)/float(total_grpValue))*100, 2), 
                                       'solar_value': round(Solar_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 0) / 100)),
                                       'dg_value': round(DG_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 2) / 100))
                                       })
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':   
                    node_srcnm = 'Transformer'
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                diffvalue = data['solar_value'] + data['dg_value']
                                link.append({"source":node_srcnm,
                                            "target":grupnme['aggroupname'],
                                            "value":(int(Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']) - diffvalue) * 0.7132})
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100)) })
                elif srcname['assourcename'] == 'Solar Energy':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100))*0.041 })
                elif srcname['assourcename'] == 'DG':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (DG_Consumption/100))*2.6 })
                else:
                    for grupnme in groupnames:
                        if grupnme['aggroupname'] != 'Incomer':
                            total_energy = Masterdatatable.objects.filter(mtdate = current_date, mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if total_energy is None:
                                total_energy = 0
                            link.append({"source":srcname['assourcename'],
                                        "target":grupnme['aggroupname'],
                                        "value":total_energy})
            # #Group to Meters
            for grupnme in groupnames:
                meternames = AddMeter.objects.filter(ammetergroup = grupnme ['aggroupname'], amplantname = Plantname).distinct('ammetername').values('ammetername')
                if grupnme['aggroupname'] != 'Incomer':
                    for mtrnme in meternames:
                        if mtrnme['ammetername'] != 'EB_Incomer':
                            energy_cons = Masterdatatable.objects.filter(mtdate = current_date, mtgrpname = grupnme['aggroupname'], mtmtrname = mtrnme['ammetername'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if energy_cons is not None:
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":float(energy_cons) * 0.7132})
                            else:
                                energy_cons = 0
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":energy_cons})

        #################################################################################################################################################################

        if request_type == 'Month Basis':
            Plantname = request.GET['plantname']   
            request_month = request_data['Month'] 
            request_year = request_data['Year']
            link = []; nodes = []; Solar_Consumption = 0; total_grpValue = 0; DG_Consumption = 0
            nodes.append({"name": "Total Emission"})
            sourcenames = AddSource.objects.filter(asplantname=Plantname).distinct('assourcename').values('assourcename')
            groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            meternames = AddMeter.objects.filter(amplantname = Plantname).distinct('ammetername').values('ammetername')
            node_groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            #Total to Incomer
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':
                    node_srcnm = 'Transformer'
                    nodes.append({"name": node_srcnm})
                    link.append({"source":"Total Emission",
                                "target":node_srcnm,
                                "value":float(Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']) * 0.7132})
                else:
                    if(srcname['assourcename'] == 'Solar Energy'):
                        try:
                            Solar_Consumption = round(float(Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":Solar_Consumption*0.041
                                        })
                        except:
                            Solar_Consumption = 0
                    elif(srcname['assourcename'] == 'DG'):
                        try:
                            DG_Consumption = round(float(Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":DG_Consumption*2.6
                                        })
                        except:
                            DG_Consumption = 0
                    else:
                        nodes.append({"name": srcname['assourcename']})
                        link.append({"source":"Total Emission",
                                    "target":srcname['assourcename'],
                                    "value":Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']})
            #Append Group Name to nodes
            for data in node_groupnames:
                if data['aggroupname'] != 'Incomer':
                    nodes.append({"name":data['aggroupname']})
                    grp_val = Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtgrpname = data['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    if grp_val is None:
                        total_grpValue = 0
                    else:
                        total_grpValue += grp_val
            #Individual Group Percent
            grpPercent = []
            for grupnme in groupnames:
                if grupnme['aggroupname'] != 'Incomer':
                    grpVal = Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    grpPercent.append({'Groupname':grupnme['aggroupname'], 
                                       'percent': round((float(grpVal)/float(total_grpValue))*100, 2), 
                                       'solar_value': round(Solar_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 0) / 100)),
                                       'dg_value': round(DG_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 2) / 100))
                                       })
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':   
                    node_srcnm = 'Transformer'
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                diffvalue = data['solar_value'] + data['dg_value']
                                link.append({"source":node_srcnm,
                                            "target":grupnme['aggroupname'],
                                            "value":(int(Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']) - diffvalue) * 0.7132})
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100)) })
                elif srcname['assourcename'] == 'Solar Energy':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100))*0.041 })
                elif srcname['assourcename'] == 'DG':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (DG_Consumption/100))*2.6 })
                else:
                    for grupnme in groupnames:
                        if grupnme['aggroupname'] != 'Incomer':
                            total_energy = Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if total_energy is None:
                                total_energy = 0
                            link.append({"source":srcname['assourcename'],
                                        "target":grupnme['aggroupname'],
                                        "value":total_energy})
            # #Group to Meters
            for grupnme in groupnames:
                meternames = AddMeter.objects.filter(ammetergroup = grupnme ['aggroupname'], amplantname = Plantname).distinct('ammetername').values('ammetername')
                if grupnme['aggroupname'] != 'Incomer':
                    for mtrnme in meternames:
                        if mtrnme['ammetername'] != 'EB_Incomer':
                            energy_cons = Masterdatatable.objects.filter(mtdate__month = request_month, mtdate__year = request_year, mtgrpname = grupnme['aggroupname'], mtmtrname = mtrnme['ammetername'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if energy_cons is not None:
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":energy_cons})
                            else:
                                energy_cons = 0
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":energy_cons})
                    
        #################################################################################################################################################################
                            
        if request_type == 'Year Basis':
            Plantname = request.GET['plantname']
            request_year = request_data['Year']
            start_date = datetime(request_year, 4, 1).strftime("%Y-%m-%d")
            end_date = datetime(request_year + 1, 3, 31).strftime("%Y-%m-%d")
            link = []; nodes = []; Solar_Consumption = 0; total_grpValue = 0; DG_Consumption = 0
            nodes.append({"name": "Total Emission"})
            sourcenames = AddSource.objects.filter(asplantname=Plantname).distinct('assourcename').values('assourcename')
            groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            meternames = AddMeter.objects.filter(amplantname = Plantname).distinct('ammetername').values('ammetername')
            node_groupnames = AddGroup.objects.filter(agplantname=Plantname).distinct('aggroupname').values('aggroupname')
            #Total to Incomer
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':
                    node_srcnm = 'Transformer'
                    nodes.append({"name": node_srcnm})
                    link.append({"source":"Total Emission",
                                "target":node_srcnm,
                                "value":float(Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']) * 0.7132})
                else:
                    if(srcname['assourcename'] == 'Solar Energy'):
                        try:
                            Solar_Consumption = round(float(Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":Solar_Consumption*0.041
                                        })
                        except:
                            Solar_Consumption = 0
                    elif(srcname['assourcename'] == 'DG'):
                        try:
                            DG_Consumption = round(float(Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']), 0)
                            nodes.append({"name": srcname['assourcename']})
                            link.append({"source":"Total Emission",
                                        "target":srcname['assourcename'],
                                        "value":DG_Consumption*2.6
                                        })
                        except:
                            DG_Consumption = 0
                    else:
                        nodes.append({"name": srcname['assourcename']})
                        link.append({"source":"Total Emission",
                                    "target":srcname['assourcename'],
                                    "value":Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = 'Incomer', mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']})
            #Append Group Name to nodes
            for data in node_groupnames:
                if data['aggroupname'] != 'Incomer':
                    nodes.append({"name":data['aggroupname']})
                    grp_val = Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtgrpname = data['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    if grp_val is None:
                        total_grpValue = 0
                    else:
                        total_grpValue += grp_val
            #Individual Group Percent
            grpPercent = []
            for grupnme in groupnames:
                if grupnme['aggroupname'] != 'Incomer':
                    grpVal = Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                    grpPercent.append({'Groupname':grupnme['aggroupname'], 
                                       'percent': round((float(grpVal)/float(total_grpValue))*100, 2), 
                                       'solar_value': round(Solar_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 0) / 100)),
                                       'dg_value': round(DG_Consumption*(round((float(grpVal)/float(total_grpValue))*100, 2) / 100))
                                       })
            for srcname in sourcenames:
                if srcname['assourcename'] == 'Transformer1':   
                    node_srcnm = 'Transformer'
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                diffvalue = data['solar_value'] + data['dg_value']
                                link.append({"source":node_srcnm,
                                            "target":grupnme['aggroupname'],
                                            "value":(int(Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']) - diffvalue) * 0.7132})
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100)) })
                elif srcname['assourcename'] == 'Solar Energy':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (Solar_Consumption/100))*0.041 })
                elif srcname['assourcename'] == 'DG':
                    for grupnme in groupnames:
                        for data in grpPercent:
                            if data['Groupname'] == grupnme['aggroupname']:
                                link.append({"source":srcname['assourcename'],
                                            "target":grupnme['aggroupname'],
                                            "value":round(data['percent'] * (DG_Consumption/100))*2.6 })
                else:
                    for grupnme in groupnames:
                        if grupnme['aggroupname'] != 'Incomer':
                            total_energy = Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtsrcname = srcname['assourcename'], mtgrpname = grupnme['aggroupname'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if total_energy is None:
                                total_energy = 0
                            link.append({"source":srcname['assourcename'],
                                        "target":grupnme['aggroupname'],
                                        "value":total_energy})
            # #Group to Meters
            for grupnme in groupnames:
                meternames = AddMeter.objects.filter(ammetergroup = grupnme ['aggroupname'], amplantname = Plantname).distinct('ammetername').values('ammetername')
                if grupnme['aggroupname'] != 'Incomer':
                    for mtrnme in meternames:
                        if mtrnme['ammetername'] != 'EB_Incomer':
                            energy_cons = Masterdatatable.objects.filter(mtdate__range = (start_date, end_date), mtgrpname = grupnme['aggroupname'], mtmtrname = mtrnme['ammetername'], mtcategory = 'Secondary', mtplntlctn = Plantname).aggregate(add('mtenergycons'))['mtenergycons__sum']
                            if energy_cons is not None:
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":energy_cons})
                            else:
                                energy_cons = 0
                                nodes.append({"name": mtrnme['ammetername']})
                                link.append({"source":grupnme['aggroupname'],
                                            "target":mtrnme['ammetername'],
                                            "value":energy_cons})
        
    return JsonResponse({"node": nodes, "link":link}, safe=False)