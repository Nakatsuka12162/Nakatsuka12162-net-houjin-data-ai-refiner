import threading
from django.utils import timezone
from .models import ExecutionHistory
from .scraper import CompanyScraper

# In-memory worker storage
active_workers = {}

class BackgroundScraper:
    def __init__(self, execution_id):
        self.execution_id = execution_id
        self.scraper = CompanyScraper()
        
    def run(self):
        """Run scraping in background thread"""
        try:
            execution = ExecutionHistory.objects.get(id=self.execution_id)
            
            # Get companies data
            data = self.scraper.upload_prompt()
            if "values" not in data:
                execution.status = 'failed'
                execution.error_message = "No data found in Google Sheets"
                execution.completed_at = timezone.now()
                execution.save()
                return
            
            # Count total companies
            companies = []
            for row in data["values"]:
                while len(row) < 4:
                    row.append("")
                corp_no = (row[0] or "").strip()
                if corp_no:
                    companies.append({
                        'corp_no': corp_no,
                        'name': row[1],
                        'addr': row[2],
                        'extra': row[3]
                    })
            
            execution.total_companies = len(companies)
            execution.save()
            
            if not companies:
                execution.status = 'failed'
                execution.error_message = "No companies to process"
                execution.completed_at = timezone.now()
                execution.save()
                return
            
            # Process companies in batches
            processed = 0
            batch_size = 3
            
            for batch_start in range(0, len(companies), batch_size):
                batch = companies[batch_start:batch_start + batch_size]
                
                try:
                    parsed_list = self.scraper.call_openai_batch(batch)
                    if parsed_list:
                        saved_count = self.scraper.save_to_database_bulk(parsed_list)
                        processed += saved_count
                        
                        # Update progress
                        execution.processed_companies = processed
                        execution.save()
                        
                except Exception as e:
                    print(f"Batch error: {e}")
                    continue
            
            # Mark as completed
            execution.status = 'completed'
            execution.processed_companies = processed
            execution.completed_at = timezone.now()
            execution.save()
            
        except Exception as e:
            execution = ExecutionHistory.objects.get(id=self.execution_id)
            execution.status = 'failed'
            execution.error_message = str(e)
            execution.completed_at = timezone.now()
            execution.save()
        finally:
            # Remove from active workers
            if self.execution_id in active_workers:
                del active_workers[self.execution_id]

def start_background_scraping():
    """Start scraping in background thread"""
    # Create execution record
    execution = ExecutionHistory.objects.create()
    
    # Start background worker
    worker = BackgroundScraper(execution.id)
    thread = threading.Thread(target=worker.run)
    thread.daemon = True
    thread.start()
    
    # Store in memory
    active_workers[execution.id] = {
        'thread': thread,
        'worker': worker,
        'started_at': timezone.now()
    }
    
    return execution.id

def get_execution_status(execution_id):
    """Get current execution status"""
    try:
        execution = ExecutionHistory.objects.get(id=execution_id)
        return {
            'id': execution.id,
            'status': execution.status,
            'started_at': execution.started_at,
            'completed_at': execution.completed_at,
            'total_companies': execution.total_companies,
            'processed_companies': execution.processed_companies,
            'error_message': execution.error_message,
            'duration': execution.duration.total_seconds() if execution.duration else 0,
            'is_active': execution_id in active_workers
        }
    except ExecutionHistory.DoesNotExist:
        return None

def get_latest_executions(limit=10):
    """Get latest execution history"""
    executions = ExecutionHistory.objects.all()[:limit]
    return [get_execution_status(ex.id) for ex in executions]
