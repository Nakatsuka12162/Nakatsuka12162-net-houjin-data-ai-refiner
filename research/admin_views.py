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
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="companies.csv"'
    
    writer = csv.writer(response)
    writer.writerow(['法人番号', '会社名', '代表者', '業種', '従業員数', '売上高', '住所'])
    
    for company in Company.objects.all():
        writer.writerow([
            company.corporate_number,
            company.company_name,
            company.representative_name,
            company.industry,
            company.employee_count,
            company.revenue,
            company.address
        ])
    
    return response

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
                    
                    # Apply user configuration
                    scraper.user_range = execution.spreadsheet_range
                    scraper.user_update_sheets = execution.update_google_sheets
                    scraper.user_max_companies = execution.max_companies
                    
                    result = scraper.scrape_companies_with_config()
                    execution.status = 'completed'
                    execution.processed_companies = result.get('processed', 0)
                    execution.total_companies = result.get('total', 0)
                    execution.completed_at = timezone.now()
                    execution.save()
                except Exception as e:
                    execution.status = 'failed'
                    execution.error_message = str(e)
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
