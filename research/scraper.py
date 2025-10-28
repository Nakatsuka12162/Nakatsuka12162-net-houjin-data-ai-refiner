import requests
import json
import logging
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread
from openai import OpenAI
from gspread_formatting import format_cell_range, CellFormat, Color
from gspread.exceptions import WorksheetNotFound, APIError
from django.conf import settings
from django.db import transaction
from .models import Company, Executive, Office, ResearchHistory

# Setup logger
logger = logging.getLogger('scraper')

# Color constants for Google Sheets
COLOR_I   = (220, 230, 241)  # 薄青
COLOR_II  = (226, 239, 218)  # 薄緑
COLOR_III = (252, 228, 214)  # 薄オレンジ
COLOR_IV  = (248, 203, 173)  # 薄ピンク
COLOR_VI  = (255, 242, 204)  # 薄黄
COLOR_VII = (217, 217, 217)  # グレー(URL)

class CompanyScraper:
    def __init__(self):
        self.api_key = settings.API_KEY
        self.spreadsheet_id = settings.SPREADSHEET_ID
        self.openai_api_key = settings.OPEN_AI_API_KEY
        self.credentials_info = settings.CREDENTIALS_INFO
        
        # Initialize OpenAI client with error handling
        if self.openai_api_key:
            self.openai_client = OpenAI(api_key=self.openai_api_key)
        else:
            self.openai_client = None
        
        # User configuration defaults
        self.user_range = "会社リスト!A3:D"
        self.user_update_sheets = True
        self.user_max_companies = None
        
        # Log collection
        self.logs = []
        
        # Prompts
        self.prompt_text1 = """
各会社の調査においては、まず必ず提示された企業法人番号を利用してGoogleで検索してください。
https://info.gbiz.go.jp/hojin/ichiran?hojinBango=
の末尾に会社法人番号を追加すると、会社に関する情報が表示されます。
ここに基本的な情報があるので、これを基本的に参考にしてください。
次の URL を検索します。提示URLに表示されない情報は、再びインターネット検索で補完されます。
調査及び対照の最優先基準は、**会社法人番号（法人番号）**とします。企業法人番号は決して変更されない。会社名・住所は変更される可能性がありますので、これらを根拠に推測・確定してください。
出力形式はJSONのみであり、説明文やコメントは必要ありません。必ず指定されたJSONスキーマに従って納品してください（ファイル以外の形式は不可）。
年齢計算の基準日は2025年9月時点とし、「50代」のような数表示は避け、可能な限り**具体的な年齢（例：52歳）**で記載してください。
調査は正確さを最優先に、慎重に実施してください。
>>>>>>
"""
        
        self.prompt_text2 = """{
  "基本法人情報（識別・概要）": {
    "企業法人番号": "",
    "会社名": "",
    "会社名かな": "",
    "英文企業名": "",
    "代表者名": "",
    "代表者かな": "",
    "代表者年齢": "",
    "代表者生年月日": "",
    "代表者出身大学": "",
    "郵便番号": "",
    "住所": "",
    "電話番号": "",
    "登記住所": "",
    "FAX番号": "",
    "URL": "",
    "創業": "",
    "設立": "",
    "資本金": "",
    "出資金": "",
    "会員数": "",
    "組合員数": "",
    "上場市場": "",
    "証券コード": "",
    "決算期": ""
  },
  "経営・財務情報": {
    "売上高": "",
    "純利益": "",
    "預金量": "",
    "従業員数": "",
    "平均年齢": "",
    "平均年収": "",
    "役員数": "",
    "株主数": "",
    "取引銀行": ""
  },
  "事業・業務内容": {
    "業種": "",
    "事業内容": "",
    "主要事業": "",
    "事業エリア": "",
    "系列": "",
    "販売先": "",
    "仕入先": ""
  },
  "役員名簿": {
    "役職名１": "", "役員名１": "", "ふりがな１": "",
    "役職名２": "", "役員名２": "", "ふりがな２": "",
    "役職名３": "", "役員名３": "", "ふりがな３": "",
    "役職名４": "", "役員名４": "", "ふりがな４": "",
    "役職名５": "", "役員名５": "", "ふりがな５": ""
  },
  "拠点・展開規模": {
    "事業所数": "",
    "店舗数": ""
  },
  "拠点・事業所一覧": {
    "事業所名１": "", "郵便番号１": "", "住所１": "", "電話番号１": "", "扱い品目・業務内容１": "",
    "事業所名２": "", "郵便番号２": "", "住所２": "", "電話番号２": "", "扱い品目・業務内容２": "",
    "事業所名３": "", "郵便番号３": "", "住所３": "", "電話番号３": "", "扱い品目・業務内容３": ""
  },
  "URL": {
    "会社概要ページURL": "",
    "拠点・事業所ページURL": "",
    "組織図ページURL": "",
    "関係会社ページURL": ""
  }
}"""

    def log(self, message, level='INFO'):
        """Add log message to both logger and internal collection"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}"
        self.logs.append(log_entry)
        
        if level == 'ERROR':
            logger.error(message)
        elif level == 'WARNING':
            logger.warning(message)
        else:
            logger.info(message)

    def call_openai_batch(self, companies_batch):
        """Process companies - compatible with existing background_worker.py"""
        results = []
        for company in companies_batch:
            result = self.call_openai_single(company)
            if result:
                results.append(result)
        return results
    
    def save_to_database_bulk(self, parsed_data_list):
        """Bulk save - compatible with existing background_worker.py"""
        saved_count = 0
        for parsed_data in parsed_data_list:
            try:
                count = self.save_to_database_single(parsed_data)
                saved_count += count
            except Exception as e:
                self.log(f"Error saving company: {e}", 'ERROR')
                continue
        return saved_count

    def call_openai_single(self, company):
        """Process one company per call for better JSON reliability"""
        if not self.openai_client:
            self.log("OpenAI API key is not set", 'ERROR')
            raise ValueError("OpenAI API key is not set.")
        
        self.log(f"Processing company: {company['corp_no']} - {company['name']}")
        
        # Single company prompt
        company_text = f"企業法人番号: {company['corp_no']}\n会社名: {company['name']}\n所在地: {company['addr']}\n補足: {company.get('extra', '')}"
        final_prompt = self.prompt_text1 + company_text + "\n\n以下のJSON形式で返してください:\n" + self.prompt_text2
        
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "あなたは会社情報を正確にJSON形式で出力するアシスタントです。JSONのみを返し、説明文は不要です。"},
                    {"role": "user", "content": final_prompt}
                ],
                temperature=0,
                max_tokens=4096
            )
            
            content = response.choices[0].message.content.strip()
            
            # Clean JSON response
            if content.startswith('```json'):
                content = content[7:]
            if content.endswith('```'):
                content = content[:-3]
            content = content.strip()
            
            try:
                parsed = json.loads(content)
                # Ensure corporate number is set
                if "基本法人情報（識別・概要）" in parsed:
                    parsed["基本法人情報（識別・概要）"]["企業法人番号"] = company['corp_no']
                self.log(f"Successfully parsed company: {company['corp_no']}")
                return parsed
            except json.JSONDecodeError as e:
                self.log(f"JSON parsing failed for {company['corp_no']}: {e}", 'ERROR')
                self.log(f"Raw response: {content[:500]}...", 'ERROR')
                return None
                
        except Exception as e:
            self.log(f"OpenAI API call failed for {company['corp_no']}: {e}", 'ERROR')
            return None

    def upload_prompt(self):
        """Get data from Google Sheets"""
        RANGE = self.user_range
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.spreadsheet_id}/values/{RANGE}?key={self.api_key}"
        try:
            response = requests.get(url)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            self.log(f"Failed to fetch Google Sheets data: {e}", 'ERROR')
            return {}

    def to_zenkaku(self, num: int) -> str:
        s = str(num)
        table = str.maketrans("0123456789", "０１２３４５６７８９")
        return s.translate(table)

    def pick(self, d: dict, base: str, i: int, default=""):
        return (d.get(f"{base}{i}") or d.get(f"{base}{self.to_zenkaku(i)}") or default)

    def extract_roles(self, parsed: dict):
        src = parsed.get("役員名簿", {}) or {}
        out = []
        i = 1
        while i <= 20:  # Limit to prevent infinite loops
            # Try both half-width and full-width numbers
            role_key = f"役職名{i}"
            name_key = f"役員名{i}"
            kana_key = f"ふりがな{i}"
            
            # If half-width doesn't exist, try full-width
            if role_key not in src:
                role_key = f"役職名{self.to_zenkaku(i)}"
            if name_key not in src:
                name_key = f"役員名{self.to_zenkaku(i)}"
            if kana_key not in src:
                kana_key = f"ふりがな{self.to_zenkaku(i)}"
            
            role = src.get(role_key, "")
            name = src.get(name_key, "")
            kana = src.get(kana_key, "")
            
            if role or name:
                out.append({
                    "役職名": role,
                    "役員名": name,
                    "ふりがな": kana,
                })
            elif not role and not name and i > 5:  # Stop if no data after first 5
                break
            i += 1
        return out

    def extract_locations(self, parsed: dict):
        src = parsed.get("拠点・事業所一覧", {}) or {}
        out = []
        i = 1
        while i <= 20:  # Limit to prevent infinite loops
            # Try both half-width and full-width numbers
            name_key = f"事業所名{i}"
            postal_key = f"郵便番号{i}"
            addr_key = f"住所{i}"
            phone_key = f"電話番号{i}"
            content_key = f"扱い品目・業務内容{i}"
            
            # If half-width doesn't exist, try full-width
            if name_key not in src:
                name_key = f"事業所名{self.to_zenkaku(i)}"
            if postal_key not in src:
                postal_key = f"郵便番号{self.to_zenkaku(i)}"
            if addr_key not in src:
                addr_key = f"住所{self.to_zenkaku(i)}"
            if phone_key not in src:
                phone_key = f"電話番号{self.to_zenkaku(i)}"
            if content_key not in src:
                content_key = f"扱い品目・業務内容{self.to_zenkaku(i)}"
            
            name = src.get(name_key, "")
            postal = src.get(postal_key, "")
            addr = src.get(addr_key, "")
            phone = src.get(phone_key, "")
            content = src.get(content_key, "")
            
            if name or addr:
                out.append({
                    "事業所名": name,
                    "郵便番号": postal,
                    "住所": addr,
                    "電話番号": phone,
                    "扱い品目・業務内容": content,
                })
            elif not name and not addr and i > 3:  # Stop if no data after first 3
                break
            i += 1
        return out

    @transaction.atomic
    def save_to_database_single(self, parsed_data):
        """Save single company to database"""
        info = parsed_data.get("基本法人情報（識別・概要）", {}) or {}
        corp_no = info.get("企業法人番号", "").strip()
        if not corp_no:
            self.log("No corporate number found in parsed data", 'ERROR')
            return 0
        
        fin = parsed_data.get("経営・財務情報", {}) or {}
        biz = parsed_data.get("事業・業務内容", {}) or {}
        scale = parsed_data.get("拠点・展開規模", {}) or {}
        urls = parsed_data.get("URL", {}) or {}
        
        company_data = {
            'company_name': info.get("会社名", ""),
            'company_name_kana': info.get("会社名かな", ""),
            'english_name': info.get("英文企業名", ""),
            'representative_name': info.get("代表者名", ""),
            'representative_kana': info.get("代表者かな", ""),
            'representative_age': info.get("代表者年齢", ""),
            'representative_birth': info.get("代表者生年月日", ""),
            'representative_university': info.get("代表者出身大学", ""),
            'postal_code': info.get("郵便番号", ""),
            'address': info.get("住所", ""),
            'phone': info.get("電話番号", ""),
            'registered_address': info.get("登記住所", ""),
            'fax': info.get("FAX番号", ""),
            'url': info.get("URL", ""),
            'founded': info.get("創業", ""),
            'established': info.get("設立", ""),
            'capital': info.get("資本金", ""),
            'investment': info.get("出資金", ""),
            'member_count': info.get("会員数", ""),
            'union_member_count': info.get("組合員数", ""),
            'stock_market': info.get("上場市場", ""),
            'stock_code': info.get("証券コード", ""),
            'fiscal_year_end': info.get("決算期", ""),
            'revenue': fin.get("売上高", ""),
            'net_profit': fin.get("純利益", ""),
            'deposits': fin.get("預金量", ""),
            'employee_count': fin.get("従業員数", ""),
            'average_age': fin.get("平均年齢", ""),
            'average_salary': fin.get("平均年収", ""),
            'executive_count': fin.get("役員数", ""),
            'shareholder_count': fin.get("株主数", ""),
            'main_bank': fin.get("取引銀行", ""),
            'industry': biz.get("業種", ""),
            'business_content': biz.get("事業内容", ""),
            'main_business': biz.get("主要事業", ""),
            'business_area': biz.get("事業エリア", ""),
            'group_affiliation': biz.get("系列", ""),
            'sales_destination': biz.get("販売先", ""),
            'supplier': biz.get("仕入先", ""),
            'office_count': scale.get("事業所数", ""),
            'store_count': scale.get("店舗数", ""),
            'company_overview_url': urls.get("会社概要ページURL", ""),
            'office_list_url': urls.get("拠点・事業所ページURL", ""),
            'organization_chart_url': urls.get("組織図ページURL", ""),
            'related_companies_url': urls.get("関係会社ページURL", ""),
        }
        
        company, created = Company.objects.update_or_create(
            corporate_number=corp_no,
            defaults=company_data
        )
        
        # Delete old related data
        company.executives.all().delete()
        company.offices.all().delete()
        
        # Create executives
        roles = self.extract_roles(parsed_data)
        self.log(f"Extracted {len(roles)} executives for {corp_no}")
        
        # Debug: print the raw data structure
        role_data = parsed_data.get("役員名簿", {})
        self.log(f"Raw role data keys: {list(role_data.keys())}")
        
        for i, role in enumerate(roles, 1):
            if role["役職名"] or role["役員名"]:
                Executive.objects.create(
                    company=company,
                    position=role["役職名"],
                    name=role["役員名"],
                    name_kana=role["ふりがな"],
                    order=i
                )
                self.log(f"Created executive: {role['役職名']} - {role['役員名']}")
        
        # Create offices
        locations = self.extract_locations(parsed_data)
        self.log(f"Extracted {len(locations)} offices for {corp_no}")
        
        # Debug: print the raw data structure
        location_data = parsed_data.get("拠点・事業所一覧", {})
        self.log(f"Raw location data keys: {list(location_data.keys())}")
        
        for i, location in enumerate(locations, 1):
            if location["事業所名"] or location["住所"]:
                Office.objects.create(
                    company=company,
                    name=location["事業所名"],
                    postal_code=location["郵便番号"],
                    address=location["住所"],
                    phone=location["電話番号"],
                    business_content=location["扱い品目・業務内容"],
                    order=i
                )
                self.log(f"Created office: {location['事業所名']} - {location['住所']}")
                
        
        self.log(f"Saved company to database: {corp_no} - {company.company_name}")
        return 1

    def scrape_companies(self):
        """Main scraping method - one company at a time for reliability"""
        try:
            self.log("Starting scraping process")
            
            # Get data from Google Sheets
            data = self.upload_prompt()
            self.log(f"Google Sheets API response keys: {list(data.keys())}")
            
            if "values" not in data:
                error_msg = f"No 'values' key in response: {data}"
                self.log(error_msg, 'ERROR')
                return {"processed": 0, "total": 0, "message": error_msg, "logs": self.logs}

            # Prepare companies
            companies = []
            self.log(f"Processing {len(data['values'])} rows from sheets")
            
            for i, row in enumerate(data["values"]):
                while len(row) < 4:
                    row.append("")
                corp_no, name, addr, extra = row
                corp_no = (corp_no or "").strip()
                
                if corp_no:
                    companies.append({
                        'corp_no': corp_no,
                        'name': name,
                        'addr': addr,
                        'extra': extra,
                        'index': i
                    })
                    self.log(f"Row {i}: Added company {corp_no} - {name}")
                else:
                    self.log(f"Row {i}: Skipped (no corporate number)")

            self.log(f"Found {len(companies)} valid companies")
            
            if not companies:
                error_msg = "No companies with valid corporate numbers found"
                self.log(error_msg, 'ERROR')
                return {"processed": 0, "total": 0, "message": error_msg, "logs": self.logs}

            # Apply max_companies limit if set
            if self.user_max_companies and self.user_max_companies > 0:
                companies = companies[:self.user_max_companies]
                self.log(f"Limited to {len(companies)} companies")

            processed = 0
            
            # Setup Google Sheets for output (if enabled)
            sh = None
            if self.user_update_sheets:
                try:
                    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
                    creds = Credentials.from_service_account_info(self.credentials_info, scopes=scopes)
                    gc = gspread.authorize(creds)
                    sh = gc.open_by_key(self.spreadsheet_id)
                    self.log("Google Sheets connection established for output")
                except Exception as e:
                    self.log(f"Failed to setup Google Sheets output: {e}", 'WARNING')
                    sh = None
            
            # Process companies one by one
            for company in companies:
                try:
                    self.log(f"Processing company {processed + 1}/{len(companies)}: {company['corp_no']}")
                    
                    # Call OpenAI for single company
                    parsed_data = self.call_openai_single(company)
                    
                    if not parsed_data:
                        self.log(f"No parsed data for {company['corp_no']}", 'WARNING')
                        continue
                    
                    # Save to database
                    saved_count = self.save_to_database_single(parsed_data)
                    if saved_count > 0:
                        processed += saved_count
                        self.log(f"Saved company to database: {company['corp_no']}")
                        
                        # Update Google Sheets if enabled
                        if sh and self.user_update_sheets:
                            try:
                                self.update_single_sheet(sh, parsed_data, company['corp_no'], company['name'], company['index'])
                                self.log(f"Updated Google Sheet for: {company['corp_no']}")
                            except Exception as e:
                                self.log(f"Sheet update error for {company['corp_no']}: {e}", 'WARNING')
                    
                except Exception as e:
                    self.log(f"Error processing company {company['corp_no']}: {e}", 'ERROR')
                    continue

            success_msg = f"Successfully processed {processed}/{len(companies)} companies"
            self.log(success_msg)
            return {"processed": processed, "total": len(companies), "message": success_msg, "logs": self.logs}
            
        except Exception as e:
            error_msg = f"Scraping failed: {str(e)}"
            self.log(error_msg, 'ERROR')
            return {"processed": 0, "total": 0, "message": error_msg, "logs": self.logs}

    def update_single_sheet(self, sh, parsed_data, corp_no, company_name, index):
        """Update single sheet with company data"""
        try:
            info = parsed_data.get("基本法人情報（識別・概要）", {}) or {}
            corp_no_from_json = (info.get("企業法人番号") or corp_no).strip()
            company_name = info.get("会社名", company_name or f"Company{index+1}")
            
            ws = self.get_or_create_company_ws(sh, corp_no_from_json, company_name)
            self.write_simple_form_to_sheet(ws, parsed_data)
        except Exception as e:
            self.log(f"Sheet error for {corp_no}: {e}", 'ERROR')
    
    def write_simple_form_to_sheet(self, ws, parsed):
        """Simple sheet writing without complex formatting"""
        rows = []
        
        info = parsed.get("基本法人情報（識別・概要）", {}) or {}
        rows.append(["基本情報", "企業法人番号", info.get("企業法人番号", "")])
        rows.append(["", "会社名", info.get("会社名", "")])
        rows.append(["", "代表者名", info.get("代表者名", "")])
        rows.append(["", "住所", info.get("住所", "")])
        rows.append(["", "電話番号", info.get("電話番号", "")])
        rows.append(["", "設立", info.get("設立", "")])
        rows.append(["", "資本金", info.get("資本金", "")])
        
        fin = parsed.get("経営・財務情報", {}) or {}
        rows.append(["財務情報", "売上高", fin.get("売上高", "")])
        rows.append(["", "従業員数", fin.get("従業員数", "")])
        rows.append(["", "平均年収", fin.get("平均年収", "")])
        
        biz = parsed.get("事業・業務内容", {}) or {}
        rows.append(["事業情報", "業種", biz.get("業種", "")])
        rows.append(["", "事業内容", biz.get("事業内容", "")])
        
        roles = self.extract_roles(parsed)
        for i, role in enumerate(roles[:10], 1):
            rows.append([f"役員{i}", role["役職名"], role["役員名"]])
        
        locs = self.extract_locations(parsed)
        for i, loc in enumerate(locs[:10], 1):
            rows.append([f"拠点{i}", loc["事業所名"], loc["住所"]])
        
        # Single update - no formatting
        ws.clear()
        ws.update("A1", rows, value_input_option='RAW')

    def get_or_create_company_ws(self, sh, corp_no: str, company_name: str):
        """Get or create worksheet for company"""
        existing_titles = [w.title for w in sh.worksheets()]
        
        # Try to find existing sheet by corporate number
        for ws in sh.worksheets():
            try:
                values = ws.get("A1:C10")
                for row in values:
                    if len(row) >= 3 and row[1] == "企業法人番号" and row[2] == corp_no:
                        return ws
            except Exception:
                continue
        
        # Create new sheet
        title = self.unique_title(corp_no if corp_no else (company_name or "Company"), existing_titles)
        ws = sh.add_worksheet(title=title, rows="1000", cols="6")
        return ws

    def unique_title(self, base: str, existing_titles):
        """Generate unique sheet title"""
        safe = (base or "Company")[:100]
        if safe not in existing_titles:
            return safe
        idx = 2
        while True:
            cand = f"{safe}_{idx}"
            if cand not in existing_titles:
                return cand
            idx += 1
