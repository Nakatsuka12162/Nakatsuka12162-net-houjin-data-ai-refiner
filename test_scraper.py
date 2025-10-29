#!/usr/bin/env python
"""Test script to verify scraper outputs detailed information"""
import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'company_research.settings')
django.setup()

from research.scraper import CompanyScraper
from research.models import Company, Executive, Office

def test_scraper():
    print("=" * 60)
    print("Testing Company Scraper - Detailed Output")
    print("=" * 60)
    
    scraper = CompanyScraper()
    
    # Test with limited companies
    scraper.user_max_companies = 1
    scraper.user_update_sheets = True
    
    print("\n1. Starting scraping process...")
    result = scraper.scrape_companies()
    
    print(f"\n2. Scraping Result:")
    print(f"   - Processed: {result.get('processed', 0)}")
    print(f"   - Total: {result.get('total', 0)}")
    print(f"   - Message: {result.get('message', '')}")
    
    print("\n3. Database Check:")
    companies = Company.objects.all()[:1]
    for company in companies:
        print(f"\n   Company: {company.company_name}")
        print(f"   Corporate Number: {company.corporate_number}")
        print(f"   Representative: {company.representative_name}")
        print(f"   Industry: {company.industry}")
        print(f"   Revenue: {company.revenue}")
        print(f"   Employees: {company.employee_count}")
        print(f"   Business Content: {company.business_content[:100] if company.business_content else 'N/A'}...")
        
        execs = company.executives.all()
        print(f"\n   Executives ({execs.count()}):")
        for exec in execs[:5]:
            print(f"     - {exec.position}: {exec.name}")
        
        offices = company.offices.all()
        print(f"\n   Offices ({offices.count()}):")
        for office in offices[:5]:
            print(f"     - {office.name}: {office.address[:50] if office.address else 'N/A'}...")
    
    print("\n4. Recent Logs:")
    for log in result.get('logs', [])[-10:]:
        print(f"   {log}")
    
    print("\n" + "=" * 60)
    print("Test Complete!")
    print("=" * 60)

if __name__ == '__main__':
    test_scraper()
