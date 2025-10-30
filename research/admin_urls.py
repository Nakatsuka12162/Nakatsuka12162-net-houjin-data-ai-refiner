from django.urls import path
from . import admin_views

urlpatterns = [
    path('', admin_views.admin_dashboard, name='admin_dashboard'),
    path('login/', admin_views.admin_login, name='admin_login'),
    path('logout/', admin_views.admin_logout, name='admin_logout'),
    
    path('companies/', admin_views.company_list, name='company_list'),
    path('companies/<int:pk>/', admin_views.company_detail, name='company_detail'),
    path('companies/<int:pk>/delete/', admin_views.delete_company, name='delete_company'),
    
    path('executives/', admin_views.executive_list, name='executive_list'),
    path('offices/', admin_views.office_list, name='office_list'),
    path('history/', admin_views.history_list, name='history_list'),
    path('scraping-history/', admin_views.scraping_history, name='scraping_history'),
    path('execution/<int:execution_id>/', admin_views.execution_detail, name='execution_detail'),
    path('execution/<int:execution_id>/download/', admin_views.export_execution_data, name='export_execution_data'),
    
    path('export/companies.csv', admin_views.export_companies_csv, name='export_companies'),
    path('export/companies_detailed.csv', admin_views.export_companies_detailed_csv, name='export_companies_detailed'),
    path('export/companies.xlsx', admin_views.export_companies_excel, name='export_companies_excel'),
    path('export/company/<int:company_id>/detailed.csv', admin_views.export_single_company_detailed_csv, name='export_single_company_detailed'),
    path('api/scrape/', admin_views.trigger_scraping, name='trigger_scraping'),
    path('api/stats/', admin_views.api_stats, name='api_stats'),
]
