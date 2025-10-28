from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from .models import Company, Executive, Office, ResearchHistory, ExecutionHistory
from .scraper import CompanyScraper
import json
import threading
from django.utils import timezone

def home(request):
    return JsonResponse({"message": "Company Research Django API is running"})

@csrf_exempt
def run_scraping(request):
    if request.method == 'POST':
        try:
            # Create execution record
            execution = ExecutionHistory.objects.create()
            
            # Start background scraping
            def background_scrape():
                try:
                    scraper = CompanyScraper()
                    result = scraper.scrape_companies()
                    execution.status = 'completed'
                    execution.processed_companies = result.get('processed', 0)
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
                "status": "started", 
                "execution_id": execution.id,
                "message": "Scraping started in background"
            })
        except Exception as e:
            return JsonResponse({"status": "error", "message": str(e)}, status=500)
    return JsonResponse({"status": "error", "message": "Only POST method allowed"}, status=405)

def health_check(request):
    return JsonResponse({"status": "healthy"})
