from django.urls import path
from Carbon import views, dashboard, analytics, yearlyemission, sankeycarbon, carbonexcel,sourcegraph

urlpatterns = [
    path('rcarbonsummary/',views.CarbonSummary),
    path('rcarbon/',dashboard.Dashboard),
    path('analyticsrcarbon/',analytics.Analytics),
    path('yearlyemission/', yearlyemission.YearlyEmission),
    path('sankeycarbon/', sankeycarbon.SankeyCarbon),
    path('carbonreport/',carbonexcel.sendmail),
    path('carbonsrc/', sourcegraph.carbongraph)
]