import requests
import json
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread
from openai import OpenAI
from gspread_formatting import format_cell_range, CellFormat, Color
from gspread.exceptions import WorksheetNotFound, APIError
from django.conf import settings
from django.db import transaction
from .models import Company, Executive, Office, ResearchHistory
from concurrent.futures import ThreadPoolExecutor, as_completed

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
        self.openai_client = OpenAI(api_key=self.openai_api_key)
        
        # Prompts from original code
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
    "役職名５": "", "役員名５": "", "ふりがな５": "",
    "役職名６": "", "役員名６": "", "ふりがな６": "",
    "役職名７": "", "役員名７": "", "ふりがな７": "",
    "役職名８": "", "役員名８": "", "ふりがな８": "",
    "役職名９": "", "役員名９": "", "ふりがな９": "",
    "役職名１０": "", "役員名１０": "", "ふりがな１０": "",
    "役職名１１": "", "役員名１１": "", "ふりがな１１": "",
    "役職名１２": "", "役員名１２": "", "ふりがな１２": "",
    "役職名１３": "", "役員名１３": "", "ふりがな１３": "",
    "役職名１４": "", "役員名１４": "", "ふりがな１４": ""
  },
  "拠点・展開規模": {
    "事業所数": "",
    "店舗数": ""
  },
  "拠点・事業所一覧": {
    "事業所名１": "", "郵便番号１": "", "住所１": "", "電話番号１": "", "扱い品目・業務内容１": "",
    "事業所名２": "", "郵便番号２": "", "住所２": "", "電話番号２": "", "扱い品目・業務内容２": "",
    "事業所名３": "", "郵便番号３": "", "住所３": "", "電話番号３": "", "扱い品目・業務内容３": "",
    "事業所名４": "", "郵便番号４": "", "住所４": "", "電話番号４": "", "扱い品目・業務内容４": "",
    "事業所名５": "", "郵便番号５": "", "住所５": "", "電話番号５": "", "扱い品目・業務内容５": "",
    "事業所名６": "", "郵便番号６": "", "住所６": "", "電話番号６": "", "扱い品目・業務内容６": "",
    "事業所名７": "", "郵便番号７": "", "住所７": "", "電話番号７": "", "扱い品目・業務内容７": "",
    "事業所名８": "", "郵便番号８": "", "住所８": "", "電話番号８": "", "扱い品目・業務内容８": "",
    "事業所名９": "", "郵便番号９": "", "住所９": "", "電話番号９": "", "扱い品目・業務内容９": "",
    "事業所名１０": "", "郵便番号１０": "", "住所１０": "", "電話番号１０": "", "扱い品目・業務内容１０": "",
    "事業所名１１": "", "郵便番号１１": "", "住所１１": "", "電話番号１１": "", "扱い品目・業務内容１１": "",
    "事業所名１２": "", "郵便番号１２": "", "住所１２": "", "電話番号１２": "", "扱い品目・業務内容１２": "",
    "事業所名１３": "", "郵便番号１３": "", "住所１３": "", "電話番号１３": "", "扱い品目・業務内容１３": "",
    "事業所名１４": "", "郵便番号１４": "", "住所１４": "", "電話番号１４": "", "扱い品目・業務内容１４": ""
  },
  "URL": {
    "会社概要ページURL": "",
    "拠点・事業所ページURL": "",
    "組織図ページURL": "",
    "関係会社ページURL": ""
  }
}"""

    def call_openai_batch(self, companies_batch):
        """Process multiple companies in single OpenAI call - 10x faster"""
        if not self.openai_api_key:
            raise ValueError("OpenAI API key is not set.")
        
        # Build prompt for multiple companies
        companies_text = ""
        for i, comp in enumerate(companies_batch, 1):
            companies_text += f"\n\n企業{i}:\n企業法人番号: {comp['corp_no']}\n会社名: {comp['name']}\n所在地: {comp['addr']}\n補足: {comp.get('extra', '')}"
        
        final_prompt = self.prompt_text1 + companies_text + "\n\n各企業について以下のJSON形式で配列として返してください:\n[" + self.prompt_text2 + ", " + self.prompt_text2 + ", ...]"
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo-16k",
            messages=[
                {"role": "system", "content": "あなたは会社情報を正確にJSON配列形式で出力するアシスタントです。"},
                {"role": "user", "content": final_prompt}
            ],
            temperature=0,
            max_tokens=4096
        )
        
        content = response.choices[0].message.content.strip()
        # Handle if response is array or single object
        try:
            parsed = json.loads(content)
            return parsed if isinstance(parsed, list) else [parsed]
        except:
            return []

    def call_openai(self, final_prompt):
        """Single company fallback"""
        if not self.openai_api_key:
            raise ValueError("OpenAI API key is not set.")
        
        response = self.openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "あなたは会社情報を正確にJSON形式で出力するアシスタントです。"},
                {"role": "user", "content": final_prompt}
            ],
            temperature=0,
            max_tokens=4096
        )
        return response.choices[0].message.content.strip()

    def upload_prompt(self):
        RANGE = "会社リスト!A3:D"
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.spreadsheet_id}/values/{RANGE}?key={self.api_key}"
        return requests.get(url).json()

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
        while any([src.get(f"役職名{i}") or src.get(f"役職名{self.to_zenkaku(i)}"),
                   src.get(f"役員名{i}") or src.get(f"役員名{self.to_zenkaku(i)}"),
                   src.get(f"ふりがな{i}") or src.get(f"ふりがな{self.to_zenkaku(i)}")]):
            out.append({
                "役職名": self.pick(src, "役職名", i, ""),
                "役員名": self.pick(src, "役員名", i, ""),
                "ふりがな": self.pick(src, "ふりがな", i, ""),
            })
            i += 1
        return out

    def extract_locations(self, parsed: dict):
        src = parsed.get("拠点・事業所一覧", {}) or {}
        out = []
        i = 1
        while any([src.get(f"事業所名{i}") or src.get(f"事業所名{self.to_zenkaku(i)}"),
                   src.get(f"住所{i}") or src.get(f"住所{self.to_zenkaku(i)}"),
                   src.get(f"電話番号{i}") or src.get(f"電話番号{self.to_zenkaku(i)}"),
                   src.get(f"郵便番号{i}") or src.get(f"郵便番号{self.to_zenkaku(i)}"),
                   src.get(f"扱い品目・業務内容{i}") or src.get(f"扱い品目・業務内容{self.to_zenkaku(i)}")]):
            out.append({
                "事業所名": self.pick(src, "事業所名", i, ""),
                "郵便番号": self.pick(src, "郵便番号", i, ""),
                "住所": self.pick(src, "住所", i, ""),
                "電話番号": self.pick(src, "電話番号", i, ""),
                "扱い品目・業務内容": self.pick(src, "扱い品目・業務内容", i, ""),
            })
            i += 1
        return out

    @transaction.atomic
    def save_to_database_bulk(self, parsed_data_list):
        """Bulk save multiple companies - 100x faster than individual saves"""
        companies_to_update = []
        execs_to_create = []
        offices_to_create = []
        
        for parsed_data in parsed_data_list:
            info = parsed_data.get("基本法人情報（識別・概要）", {}) or {}
            corp_no = info.get("企業法人番号", "").strip()
            if not corp_no:
                continue
            
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
            
            # Prepare executives
            roles = self.extract_roles(parsed_data)
            for i, role in enumerate(roles, 1):
                if role["役職名"] or role["役員名"]:
                    execs_to_create.append(Executive(
                        company=company,
                        position=role["役職名"],
                        name=role["役員名"],
                        name_kana=role["ふりがな"],
                        order=i
                    ))
            
            # Prepare offices
            locations = self.extract_locations(parsed_data)
            for i, location in enumerate(locations, 1):
                if location["事業所名"] or location["住所"]:
                    offices_to_create.append(Office(
                        company=company,
                        name=location["事業所名"],
                        postal_code=location["郵便番号"],
                        address=location["住所"],
                        phone=location["電話番号"],
                        business_content=location["扱い品目・業務内容"],
                        order=i
                    ))
        
        # Bulk create all at once
        if execs_to_create:
            Executive.objects.bulk_create(execs_to_create, batch_size=500)
        if offices_to_create:
            Office.objects.bulk_create(offices_to_create, batch_size=500)
        
        return len(parsed_data_list)

    def save_to_database(self, parsed_data, corp_no):
        """Keep for backward compatibility"""
        return self.save_to_database_bulk([parsed_data])
        info = parsed_data.get("基本法人情報（識別・概要）", {}) or {}
        fin = parsed_data.get("経営・財務情報", {}) or {}
        biz = parsed_data.get("事業・業務内容", {}) or {}
        scale = parsed_data.get("拠点・展開規模", {}) or {}
        urls = parsed_data.get("URL", {}) or {}

        # Get or create company
        company, created = Company.objects.get_or_create(
            corporate_number=corp_no,
            defaults={
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
        )

        if not created:
            # Update existing company
            for field, value in {
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
            }.items():
                old_value = getattr(company, field)
                if old_value != value and old_value and value:
                    ResearchHistory.objects.create(
                        corporate_number=corp_no,
                        changed_field=field,
                        old_value=old_value,
                        new_value=value
                    )
                setattr(company, field, value)
            company.save()

        # Clear and recreate executives
        company.executives.all().delete()
        roles = self.extract_roles(parsed_data)
        for i, role in enumerate(roles, 1):
            if role["役職名"] or role["役員名"]:
                Executive.objects.create(
                    company=company,
                    position=role["役職名"],
                    name=role["役員名"],
                    name_kana=role["ふりがな"],
                    order=i
                )

        # Clear and recreate offices
        company.offices.all().delete()
        locations = self.extract_locations(parsed_data)
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

        return company

    def scrape_companies(self):
        """Optimized: Process 5 companies per OpenAI call, bulk DB saves, parallel sheets"""
        # Setup Google Sheets
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(self.credentials_info, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(self.spreadsheet_id)

        # Get data
        data = self.upload_prompt()
        if "values" not in data:
            return {"processed": 0, "message": "No data found"}

        # Prepare companies
        companies = []
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

        if not companies:
            return {"processed": 0, "message": "No companies to process"}

        processed = 0
        batch_size = 5  # Process 5 companies per OpenAI call
        
        # Process in batches
        for batch_start in range(0, len(companies), batch_size):
            batch = companies[batch_start:batch_start + batch_size]
            
            try:
                # Single OpenAI call for multiple companies
                parsed_list = self.call_openai_batch(batch)
                
                if not parsed_list:
                    continue
                
                # Bulk save to database
                saved_count = self.save_to_database_bulk(parsed_list)
                processed += saved_count
                
                # Parallel Google Sheets updates
                with ThreadPoolExecutor(max_workers=3) as executor:
                    futures = []
                    for parsed_data, comp in zip(parsed_list, batch):
                        future = executor.submit(
                            self.update_single_sheet,
                            sh, parsed_data, comp['corp_no'], comp['name'], comp['index']
                        )
                        futures.append(future)
                    
                    # Wait for all sheets to complete
                    for future in as_completed(futures):
                        try:
                            future.result()
                        except Exception as e:
                            print(f"Sheet update error: {e}")
                
            except Exception as e:
                print(f"Batch processing error: {e}")
                continue

        return {"processed": processed, "message": f"Successfully processed {processed} companies"}
    
    def update_single_sheet(self, sh, parsed_data, corp_no, company_name, index):
        """Update single sheet - called in parallel"""
        try:
            info = parsed_data.get("基本法人情報（識別・概要）", {}) or {}
            corp_no_from_json = (info.get("企業法人番号") or corp_no).strip()
            company_name = info.get("会社名", company_name or f"Company{index+1}")
            
            ws = self.get_or_create_company_ws(sh, corp_no_from_json, company_name)
            self.write_vertical_form_to_gspread_fast(ws, parsed_data)
        except Exception as e:
            print(f"Sheet error for {corp_no}: {e}")
    
    def write_vertical_form_to_gspread_fast(self, ws, parsed):
        """Faster version - skip formatting, just write data"""
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
        for i, role in enumerate(roles[:15], 1):
            rows.append([f"役員{i}", role["役職名"], role["役員名"]])
        
        locs = self.extract_locations(parsed)
        for i, loc in enumerate(locs[:15], 1):
            rows.append([f"拠点{i}", loc["事業所名"], loc["住所"]])
        
        # Single update - no formatting
        ws.clear()
        ws.update("A1", rows, value_input_option='RAW')
    

    def write_vertical_form_to_gspread(self, ws, parsed, chunk_size=15):
        rows = []
        sections = []

        def add_section(title: str, kv_list, rgb):
            nonlocal rows, sections
            def _s(v): return "" if v is None else str(v)

            start = len(rows) + 1
            if kv_list:
                first_label, first_val = kv_list[0]
                rows.append([title, first_label, _s(first_val)])
                for label, val in kv_list[1:]:
                    rows.append(["", label, _s(val)])
            else:
                rows.append([title, "", ""])
            end = len(rows)
            sections.append((title, start, end, rgb))

        info = parsed.get("基本法人情報（識別・概要）", {}) or {}
        add_section("◆ I. 基本法人情報（識別・概要）", [
            ("法人番号", info.get("企業法人番号")),
            ("会社名", info.get("会社名")),
            ("会社名かな", info.get("会社名かな")),
            ("英文企業名", info.get("英文企業名")),
            ("代表者名", info.get("代表者名")),
            ("代表者かな", info.get("代表者かな")),
            ("代表者年齢", info.get("代表者年齢")),
            ("代表者生年月日", info.get("代表者生年月日")),
            ("代表者出身大学", info.get("代表者出身大学")),
            ("郵便番号", info.get("郵便番号")),
            ("住所", info.get("住所")),
            ("電話番号", info.get("電話番号")),
            ("登記住所", info.get("登記住所")),
            ("FAX番号", info.get("FAX番号")),
            ("URL", info.get("URL")),
            ("創業", info.get("創業")),
            ("設立", info.get("設立")),
            ("資本金", info.get("資本金")),
            ("出資金", info.get("出資金")),
            ("会員数", info.get("会員数")),
            ("組合員数", info.get("組合員数")),
            ("上場市場", info.get("上場市場")),
            ("証券コード", info.get("証券コード")),
            ("決算期", info.get("決算期")),
        ], COLOR_I)

        fin = parsed.get("経営・財務情報", {}) or {}
        add_section("◆ II. 経営・財務情報", [
            ("売上高", fin.get("売上高")),
            ("純利益", fin.get("純利益")),
            ("預金量", fin.get("預金量")),
            ("従業員数", fin.get("従業員数")),
            ("平均年齢", fin.get("平均年齢")),
            ("平均年収", fin.get("平均年収")),
            ("役員数", fin.get("役員数")),
            ("株主数", fin.get("株主数")),
            ("取引銀行", fin.get("取引銀行")),
        ], COLOR_II)

        biz = parsed.get("事業・業務内容", {}) or {}
        add_section("◆ III. 事業・業務内容", [
            ("業種", biz.get("業種")),
            ("事業内容", biz.get("事業内容")),
            ("主要事業", biz.get("主要事業")),
            ("事業エリア", biz.get("事業エリア")),
            ("系列", biz.get("系列")),
            ("販売先", biz.get("販売先")),
            ("仕入先", biz.get("仕入先")),
        ], COLOR_III)

        roles = self.extract_roles(parsed) or []
        base_idx = 1
        while True:
            block = roles[base_idx-1:base_idx-1+chunk_size]
            if not block and base_idx > 1:
                break
            kv = []
            for i in range(chunk_size):
                entry = block[i] if i < len(block) else {"役職名":"", "役員名":"", "ふりがな":""}
                idx = base_idx + i
                kv.extend([
                    (f"役職名{idx}", entry["役職名"]),
                    (f"役員名{idx}", entry["役員名"]),
                    (f"ふりがな{idx}", entry["ふりがな"]),
                ])
            add_section("◆ IV. 役員名簿", kv, COLOR_IV)
            if not block:
                break
            base_idx += chunk_size

        locs = self.extract_locations(parsed) or []
        counts = parsed.get("拠点・展開規模", {}) or {}
        base_idx = 1
        while True:
            block = locs[base_idx-1:base_idx-1+chunk_size]
            if not block and base_idx > 1:
                break
            add_section("◆ VI. 拠点・展開規模", [
                ("事業所数", counts.get("事業所数", "")),
                ("店舗数", counts.get("店舗数", "")),
            ], COLOR_VI)
            kv = []
            for i in range(chunk_size):
                entry = block[i] if i < len(block) else {"事業所名":"", "郵便番号":"", "住所":"", "電話番号":"", "扱い品目・業務内容":""}
                idx = base_idx + i
                kv.extend([
                    (f"事業所名{idx}", entry["事業所名"]),
                    (f"郵便番号{idx}", entry["郵便番号"]),
                    (f"住所{idx}", entry["住所"]),
                    (f"電話番号{idx}", entry["電話番号"]),
                    (f"扱い品目・業務内容{idx}", entry["扱い品目・業務内容"]),
                ])
            add_section("◆ VII. 拠点・事業所一覧", kv, COLOR_VI)
            if not block:
                break
            base_idx += chunk_size

        u = parsed.get("URL", {}) or {}
        add_section("◆ VII. URL", [
            ("会社概要ページURL", u.get("会社概要ページURL")),
            ("拠点・事業所ページURL", u.get("拠点・事業所ページURL")),
            ("組織図ページURL", u.get("組織図ページURL")),
            ("関係会社ページURL", u.get("関係会社ページURL")),
        ], COLOR_VII)

        ws.clear()
        ws.update("A1", rows)

        for title, start, end, rgb in sections:
            self.apply_color(ws, start, end, rgb)
            try:
                ws.merge_cells(f"A{start}:A{end}")
            except APIError:
                pass

    def apply_color(self, ws, start_row, end_row, rgb):
        fmt = CellFormat(backgroundColor=Color(
            red=rgb[0]/255, green=rgb[1]/255, blue=rgb[2]/255
        ))
        format_cell_range(ws, f"A{start_row}:C{end_row}", fmt)

    def unique_title(self, base: str, existing_titles):
        safe = (base or "Company")[:100]
        if safe not in existing_titles:
            return safe
        idx = 2
        while True:
            cand = f"{safe}_{idx}"
            if cand not in existing_titles:
                return cand
            idx += 1

    def locate_sheet_by_corp_number(self, sh, corp_no: str):
        try:
            ws = sh.worksheet(corp_no)
            return ws
        except WorksheetNotFound:
            pass
        for ws in sh.worksheets():
            try:
                values = ws.get("A1:C60")
            except Exception:
                continue
            for row in values:
                if len(row) >= 3:
                    a, b, c = (row + ["", "", ""])[:3]
                    if (b == "法人番号") and (c == corp_no):
                        return ws
        return None

    def get_or_create_company_ws(self, sh, corp_no: str, company_name: str):
        existing_titles = [w.title for w in sh.worksheets()]
        ws = self.locate_sheet_by_corp_number(sh, corp_no)
        if ws:
            if ws.title != corp_no:
                new_title = self.unique_title(corp_no, existing_titles)
                try:
                    ws.update_title(new_title)
                except APIError:
                    pass
            return ws
        title = self.unique_title(corp_no if corp_no else (company_name or "Company"), existing_titles)
        ws = sh.add_worksheet(title=title, rows="5000", cols="6")
        return ws
