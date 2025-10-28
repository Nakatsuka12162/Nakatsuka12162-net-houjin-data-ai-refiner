from django.shortcuts import render, get_object_or_404, redirect
from django.http import JsonResponse, HttpResponse
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.core.paginator import Paginator
from django.db.models import Q, Count
from django.views.decorators.csrf import csrf_exempt
from .models import Company, Executive, Office, ResearchHistory, ExecutionHistory
import csv
import threading
from django.utils import timezone

def admin_login(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(request, username=username, password=password)
        if user and user.is_staff:
            login(request, user)
            return redirect('admin_dashboard')
        messages.error(request, 'Invalid credentials')
    return render(request, 'admin/login.html')

def admin_logout(request):
    logout(request)
    return redirect('admin_login')

@login_required
def admin_dashboard(request):
    stats = {
        'total_companies': Company.objects.count(),
        'total_executives': Executive.objects.count(),
        'total_offices': Office.objects.count(),
        'recent_changes': ResearchHistory.objects.order_by('-timestamp')[:10],
        'industry_breakdown': Company.objects.values('industry').annotate(count=Count('id')).order_by('-count')[:5],
        'recent_companies': Company.objects.order_by('-created_at')[:5],
    }
    return render(request, 'admin/dashboard.html', stats)

@login_required
def company_list(request):
    search = request.GET.get('search', '')
    industry = request.GET.get('industry', '')
    
    companies = Company.objects.all()
    if search:
        companies = companies.filter(
            Q(company_name__icontains=search) |
            Q(corporate_number__icontains=search) |
            Q(representative_name__icontains=search)
        )
    if industry:
        companies = companies.filter(industry=industry)
    
    paginator = Paginator(companies.order_by('-updated_at'), 25)
    page = paginator.get_page(request.GET.get('page'))
    
    industries = Company.objects.values_list('industry', flat=True).distinct()
    
    return render(request, 'admin/company_list.html', {
        'companies': page,
        'search': search,
        'industry': industry,
        'industries': industries,
    })

@login_required
def company_detail(request, pk):
    company = get_object_or_404(Company, pk=pk)
    executives = company.executives.all().order_by('order')
    offices = company.offices.all().order_by('order')
    history = ResearchHistory.objects.filter(corporate_number=company.corporate_number).order_by('-timestamp')[:20]
    
    if request.method == 'POST':
        # Update company
        for field in ['company_name', 'representative_name', 'industry', 'address', 'phone', 'revenue', 'employee_count']:
            if field in request.POST:
                setattr(company, field, request.POST[field])
        company.save()
        messages.success(request, 'Company updated successfully')
        return redirect('company_detail', pk=pk)
    
    return render(request, 'admin/company_detail.html', {
        'company': company,
        'executives': executives,
        'offices': offices,
        'history': history,
    })

@login_required
def executive_list(request):
    search = request.GET.get('search', '')
    executives = Executive.objects.select_related('company').all()
    
    if search:
        executives = executives.filter(
            Q(name__icontains=search) |
            Q(position__icontains=search) |
            Q(company__company_name__icontains=search)
        )
    
    paginator = Paginator(executives.order_by('company__company_name'), 50)
    page = paginator.get_page(request.GET.get('page'))
    
    return render(request, 'admin/executive_list.html', {
        'executives': page,
        'search': search,
    })

@login_required
def office_list(request):
    search = request.GET.get('search', '')
    offices = Office.objects.select_related('company').all()
    
    if search:
        offices = offices.filter(
            Q(name__icontains=search) |
            Q(address__icontains=search) |
            Q(company__company_name__icontains=search)
        )
    
    paginator = Paginator(offices.order_by('company__company_name'), 50)
    page = paginator.get_page(request.GET.get('page'))
    
    return render(request, 'admin/office_list.html', {
        'offices': page,
        'search': search,
    })

@login_required
def history_list(request):
    search = request.GET.get('search', '')
    history = ResearchHistory.objects.all()
    
    if search:
        history = history.filter(
            Q(corporate_number__icontains=search) |
            Q(changed_field__icontains=search)
        )
    
    paginator = Paginator(history.order_by('-timestamp'), 100)
    page = paginator.get_page(request.GET.get('page'))
    
    return render(request, 'admin/history_list.html', {
        'history': page,
        'search': search,
    })

@login_required
def export_companies_csv(request):
    response = HttpResponse(content_type='text/csv; charset=utf-8-sig')  # UTF-8 with BOM
    response['Content-Disposition'] = 'attachment; filename="companies.csv"'
    
    writer = csv.writer(response)
    writer.writerow([
        'Corporate Number', 'Company Name', 'Representative', 'Industry', 'Employees', 'Revenue', 'Address',
        'Phone', 'Established', 'Capital', 'Business Content', 'Updated'
    ])
    
    for company in Company.objects.all():
        writer.writerow([
            f"'{company.corporate_number}" if company.corporate_number else '',  # Add quote to force text
            company.company_name or '',
            company.representative_name or '',
            company.industry or '',
            company.employee_count or '',
            company.revenue or '',
            company.address or '',
            company.phone or '',
            company.established or '',
            company.capital or '',
            company.business_content or '',
            company.updated_at.strftime('%Y-%m-%d %H:%M:%S') if company.updated_at else ''
        ])
    
    return response

@login_required
def export_companies_excel(request):
    """Simple Excel export"""
    try:
        import openpyxl
        from io import BytesIO
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Companies"
        
        # Headers
        headers = ['Corporate Number', 'Company Name', 'Representative', 'Industry', 'Employees', 'Revenue', 'Address', 'Phone', 'Updated']
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        
        # Data
        for row, company in enumerate(Company.objects.all(), 2):
            # Format corporate number as text to prevent scientific notation
            corp_cell = ws.cell(row=row, column=1, value=company.corporate_number or '')
            corp_cell.number_format = '@'  # Text format
            
            ws.cell(row=row, column=2, value=company.company_name or '')
            ws.cell(row=row, column=3, value=company.representative_name or '')
            ws.cell(row=row, column=4, value=company.industry or '')
            ws.cell(row=row, column=5, value=company.employee_count or '')
            ws.cell(row=row, column=6, value=company.revenue or '')
            ws.cell(row=row, column=7, value=company.address or '')
            ws.cell(row=row, column=8, value=company.phone or '')
            ws.cell(row=row, column=9, value=company.updated_at.strftime('%Y-%m-%d %H:%M:%S') if company.updated_at else '')
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="companies.xlsx"'
        return response
        
    except ImportError:
        return export_companies_csv(request)

@login_required
def scraping_history(request):
    """View for scraping execution history"""
    executions = ExecutionHistory.objects.all()
    
    paginator = Paginator(executions.order_by('-started_at'), 20)
    page = paginator.get_page(request.GET.get('page'))
    
    return render(request, 'admin/scraping_history.html', {
        'executions': page,
    })

@login_required
def execution_detail(request, execution_id):
    """View execution details and allow Excel download"""
    execution = get_object_or_404(ExecutionHistory, id=execution_id)
    
    # Get companies processed in this execution (approximate by time range)
    if execution.completed_at:
        companies = Company.objects.filter(
            updated_at__gte=execution.started_at,
            updated_at__lte=execution.completed_at
        ).order_by('-updated_at')
    else:
        companies = Company.objects.filter(
            updated_at__gte=execution.started_at
        ).order_by('-updated_at')[:execution.processed_companies]
    
    return render(request, 'admin/execution_detail.html', {
        'execution': execution,
        'companies': companies[:50],  # Limit display
        'total_companies': companies.count(),
    })

@login_required
def export_execution_data(request, execution_id):
    """Export all data from specific execution"""
    execution = get_object_or_404(ExecutionHistory, id=execution_id)
    
    # Get companies from this execution
    if execution.completed_at:
        companies = Company.objects.filter(
            updated_at__gte=execution.started_at,
            updated_at__lte=execution.completed_at
        ).order_by('-updated_at')
    else:
        companies = Company.objects.filter(
            updated_at__gte=execution.started_at
        ).order_by('-updated_at')[:execution.processed_companies]
    
    try:
        import openpyxl
        from openpyxl.styles import NamedStyle
        from io import BytesIO
        
        wb = openpyxl.Workbook()
        
        # Create text style to prevent scientific notation
        text_style = NamedStyle(name="text_style")
        text_style.number_format = '@'  # Text format
        
        # Companies sheet
        ws1 = wb.active
        ws1.title = "Companies"
        
        headers = ['Corporate Number', 'Company Name', 'Representative', 'Industry', 'Employees', 'Revenue', 'Address', 'Phone', 'Updated']
        for col, header in enumerate(headers, 1):
            ws1.cell(row=1, column=col, value=header)
        
        for row, company in enumerate(companies, 2):
            # Format corporate number as text to prevent scientific notation
            corp_cell = ws1.cell(row=row, column=1, value=company.corporate_number or '')
            corp_cell.number_format = '@'
            
            ws1.cell(row=row, column=2, value=company.company_name or '')
            ws1.cell(row=row, column=3, value=company.representative_name or '')
            ws1.cell(row=row, column=4, value=company.industry or '')
            ws1.cell(row=row, column=5, value=company.employee_count or '')
            ws1.cell(row=row, column=6, value=company.revenue or '')
            ws1.cell(row=row, column=7, value=company.address or '')
            ws1.cell(row=row, column=8, value=company.phone or '')
            ws1.cell(row=row, column=9, value=company.updated_at.strftime('%Y-%m-%d %H:%M:%S') if company.updated_at else '')
        
        # Executives sheet
        ws2 = wb.create_sheet("Executives")
        exec_headers = ['Corporate Number', 'Company Name', 'Position', 'Name', 'Name Kana', 'Order']
        for col, header in enumerate(exec_headers, 1):
            ws2.cell(row=1, column=col, value=header)
        
        exec_row = 2
        for company in companies:
            for executive in company.executives.all():
                corp_cell = ws2.cell(row=exec_row, column=1, value=company.corporate_number)
                corp_cell.number_format = '@'
                
                ws2.cell(row=exec_row, column=2, value=company.company_name)
                ws2.cell(row=exec_row, column=3, value=executive.position)
                ws2.cell(row=exec_row, column=4, value=executive.name)
                ws2.cell(row=exec_row, column=5, value=executive.name_kana)
                ws2.cell(row=exec_row, column=6, value=executive.order)
                exec_row += 1
        
        # Offices sheet
        ws3 = wb.create_sheet("Offices")
        office_headers = ['Corporate Number', 'Company Name', 'Office Name', 'Postal Code', 'Address', 'Phone', 'Business Content', 'Order']
        for col, header in enumerate(office_headers, 1):
            ws3.cell(row=1, column=col, value=header)
        
        office_row = 2
        for company in companies:
            for office in company.offices.all():
                corp_cell = ws3.cell(row=office_row, column=1, value=company.corporate_number)
                corp_cell.number_format = '@'
                
                ws3.cell(row=office_row, column=2, value=company.company_name)
                ws3.cell(row=office_row, column=3, value=office.name)
                ws3.cell(row=office_row, column=4, value=office.postal_code)
                ws3.cell(row=office_row, column=5, value=office.address)
                ws3.cell(row=office_row, column=6, value=office.phone)
                ws3.cell(row=office_row, column=7, value=office.business_content)
                ws3.cell(row=office_row, column=8, value=office.order)
                office_row += 1
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"execution_{execution_id}_{execution.started_at.strftime('%Y%m%d_%H%M%S')}.xlsx"
        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
        
    except ImportError:
        # Fallback to CSV with proper encoding
        response = HttpResponse(content_type='text/csv; charset=utf-8-sig')  # UTF-8 with BOM
        filename = f"execution_{execution_id}_{execution.started_at.strftime('%Y%m%d_%H%M%S')}.csv"
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        
        writer = csv.writer(response)
        writer.writerow(['Corporate Number', 'Company Name', 'Representative', 'Industry', 'Employees', 'Revenue', 'Address', 'Updated'])
        
        for company in companies:
            writer.writerow([
                f"'{company.corporate_number}" if company.corporate_number else '',  # Add quote to force text
                company.company_name or '',
                company.representative_name or '',
                company.industry or '',
                company.employee_count or '',
                company.revenue or '',
                company.address or '',
                company.updated_at.strftime('%Y-%m-%d %H:%M:%S') if company.updated_at else ''
            ])
        
        return response

@login_required
@csrf_exempt
def trigger_scraping(request):
    if request.method == 'POST':
        try:
            import json
            
            # Parse configuration from request body
            config = {}
            if request.content_type == 'application/json' and request.body:
                config = json.loads(request.body)
            
            # Create execution record with user configuration
            execution = ExecutionHistory.objects.create(
                spreadsheet_range=config.get('spreadsheet_range', '会社リスト!A3:D'),
                update_google_sheets=config.get('update_google_sheets', True),
                description=config.get('description', ''),
                max_companies=config.get('max_companies', None)
            )
            
            # Start background scraping
            def background_scrape():
                try:
                    from .scraper import CompanyScraper
                    scraper = CompanyScraper()
                    
                    # Check if user provided custom configuration
                    if (execution.spreadsheet_range != '会社リスト!A3:D' or 
                        not execution.update_google_sheets or 
                        execution.max_companies or 
                        execution.description):
                        # Use configured scraping
                        scraper.user_range = execution.spreadsheet_range
                        scraper.user_update_sheets = execution.update_google_sheets
                        scraper.user_max_companies = execution.max_companies
                        result = scraper.scrape_companies_with_config()
                    else:
                        # Use default scraping
                        result = scraper.scrape_companies()
                    
                    execution.status = 'completed'
                    execution.processed_companies = result.get('processed', 0)
                    execution.total_companies = result.get('total', result.get('processed', 0))
                    
                    # Store logs for debugging
                    logs = result.get('logs', [])
                    if logs:
                        execution.error_message = '\n'.join(logs[-50:])  # Last 50 log entries
                    
                    execution.completed_at = timezone.now()
                    execution.save()
                    
                except Exception as e:
                    import traceback
                    error_details = f"Error: {str(e)}\nTraceback: {traceback.format_exc()}"
                    execution.status = 'failed'
                    execution.error_message = error_details
                    execution.completed_at = timezone.now()
                    execution.save()
            
            thread = threading.Thread(target=background_scrape)
            thread.daemon = True
            thread.start()
            
            return JsonResponse({
                'status': 'started', 
                'execution_id': execution.id,
                'message': 'Scraping started with user configuration',
                'config': {
                    'range': execution.spreadsheet_range,
                    'update_sheets': execution.update_google_sheets,
                    'description': execution.description,
                    'max_companies': execution.max_companies
                }
            })
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})
    return JsonResponse({'status': 'error', 'message': 'POST required'})

@login_required
def api_stats(request):
    return JsonResponse({
        'companies': Company.objects.count(),
        'executives': Executive.objects.count(),
        'offices': Office.objects.count(),
        'recent_changes': ResearchHistory.objects.count(),
    })
