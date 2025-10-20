from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from .models import Company, Executive, Office, ResearchHistory
from .scraper import CompanyScraper
import json

def home(request):
    return JsonResponse({"message": "Company Research Django API is running"})

@csrf_exempt
def run_scraping(request):
    if request.method == 'POST':
        try:
            scraper = CompanyScraper()
            result = scraper.scrape_companies()
            return JsonResponse({"status": "success", "message": "Scraping completed successfully", "result": result})
        except Exception as e:
            return JsonResponse({"status": "error", "message": str(e)}, status=500)
    return JsonResponse({"status": "error", "message": "Only POST method allowed"}, status=405)

def health_check(request):
    return JsonResponse({"status": "healthy"})
