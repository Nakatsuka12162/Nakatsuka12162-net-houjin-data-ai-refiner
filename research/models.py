from django.db import models

class Company(models.Model):
    # Basic Corporate Info
    corporate_number = models.CharField(max_length=13, unique=True, verbose_name="企業法人番号")
    company_name = models.CharField(max_length=200, verbose_name="会社名")
    company_name_kana = models.CharField(max_length=200, blank=True, verbose_name="会社名かな")
    english_name = models.CharField(max_length=200, blank=True, verbose_name="英文企業名")
    representative_name = models.CharField(max_length=100, blank=True, verbose_name="代表者名")
    representative_kana = models.CharField(max_length=100, blank=True, verbose_name="代表者かな")
    representative_age = models.CharField(max_length=10, blank=True, verbose_name="代表者年齢")
    representative_birth = models.CharField(max_length=20, blank=True, verbose_name="代表者生年月日")
    representative_university = models.CharField(max_length=100, blank=True, verbose_name="代表者出身大学")
    postal_code = models.CharField(max_length=10, blank=True, verbose_name="郵便番号")
    address = models.TextField(blank=True, verbose_name="住所")
    phone = models.CharField(max_length=20, blank=True, verbose_name="電話番号")
    registered_address = models.TextField(blank=True, verbose_name="登記住所")
    fax = models.CharField(max_length=20, blank=True, verbose_name="FAX番号")
    url = models.URLField(blank=True, verbose_name="URL")
    founded = models.CharField(max_length=20, blank=True, verbose_name="創業")
    established = models.CharField(max_length=20, blank=True, verbose_name="設立")
    capital = models.CharField(max_length=50, blank=True, verbose_name="資本金")
    investment = models.CharField(max_length=50, blank=True, verbose_name="出資金")
    member_count = models.CharField(max_length=20, blank=True, verbose_name="会員数")
    union_member_count = models.CharField(max_length=20, blank=True, verbose_name="組合員数")
    stock_market = models.CharField(max_length=50, blank=True, verbose_name="上場市場")
    stock_code = models.CharField(max_length=10, blank=True, verbose_name="証券コード")
    fiscal_year_end = models.CharField(max_length=20, blank=True, verbose_name="決算期")
    
    # Financial Info
    revenue = models.CharField(max_length=50, blank=True, verbose_name="売上高")
    net_profit = models.CharField(max_length=50, blank=True, verbose_name="純利益")
    deposits = models.CharField(max_length=50, blank=True, verbose_name="預金量")
    employee_count = models.CharField(max_length=20, blank=True, verbose_name="従業員数")
    average_age = models.CharField(max_length=10, blank=True, verbose_name="平均年齢")
    average_salary = models.CharField(max_length=50, blank=True, verbose_name="平均年収")
    executive_count = models.CharField(max_length=10, blank=True, verbose_name="役員数")
    shareholder_count = models.CharField(max_length=20, blank=True, verbose_name="株主数")
    main_bank = models.CharField(max_length=100, blank=True, verbose_name="取引銀行")
    
    # Business Info
    industry = models.CharField(max_length=100, blank=True, verbose_name="業種")
    business_content = models.TextField(blank=True, verbose_name="事業内容")
    main_business = models.TextField(blank=True, verbose_name="主要事業")
    business_area = models.CharField(max_length=100, blank=True, verbose_name="事業エリア")
    group_affiliation = models.CharField(max_length=100, blank=True, verbose_name="系列")
    sales_destination = models.TextField(blank=True, verbose_name="販売先")
    supplier = models.TextField(blank=True, verbose_name="仕入先")
    
    # Scale Info
    office_count = models.CharField(max_length=20, blank=True, verbose_name="事業所数")
    store_count = models.CharField(max_length=20, blank=True, verbose_name="店舗数")
    
    # URLs
    company_overview_url = models.URLField(blank=True, verbose_name="会社概要ページURL")
    office_list_url = models.URLField(blank=True, verbose_name="拠点・事業所ページURL")
    organization_chart_url = models.URLField(blank=True, verbose_name="組織図ページURL")
    related_companies_url = models.URLField(blank=True, verbose_name="関係会社ページURL")
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "会社"
        verbose_name_plural = "会社一覧"
    
    def __str__(self):
        return f"{self.company_name} ({self.corporate_number})"

class Executive(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE, related_name='executives')
    position = models.CharField(max_length=50, verbose_name="役職名")
    name = models.CharField(max_length=100, verbose_name="役員名")
    name_kana = models.CharField(max_length=100, blank=True, verbose_name="ふりがな")
    order = models.PositiveIntegerField(default=1, verbose_name="順序")
    
    class Meta:
        verbose_name = "役員"
        verbose_name_plural = "役員一覧"
        ordering = ['order']
    
    def __str__(self):
        return f"{self.company.company_name} - {self.position}: {self.name}"

class Office(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE, related_name='offices')
    name = models.CharField(max_length=100, verbose_name="事業所名")
    postal_code = models.CharField(max_length=10, blank=True, verbose_name="郵便番号")
    address = models.TextField(blank=True, verbose_name="住所")
    phone = models.CharField(max_length=20, blank=True, verbose_name="電話番号")
    business_content = models.TextField(blank=True, verbose_name="扱い品目・業務内容")
    order = models.PositiveIntegerField(default=1, verbose_name="順序")
    
    class Meta:
        verbose_name = "事業所"
        verbose_name_plural = "事業所一覧"
        ordering = ['order']
    
    def __str__(self):
        return f"{self.company.company_name} - {self.name}"

class ResearchHistory(models.Model):
    corporate_number = models.CharField(max_length=13, verbose_name="企業法人番号")
    changed_field = models.CharField(max_length=100, verbose_name="変更列")
    old_value = models.TextField(blank=True, verbose_name="変更前")
    new_value = models.TextField(blank=True, verbose_name="変更後")
    timestamp = models.DateTimeField(auto_now_add=True, verbose_name="日時")
    
    class Meta:
        verbose_name = "履歴"
        verbose_name_plural = "履歴一覧"
        ordering = ['-timestamp']
    
    def __str__(self):
        return f"{self.corporate_number} - {self.changed_field} ({self.timestamp})"
