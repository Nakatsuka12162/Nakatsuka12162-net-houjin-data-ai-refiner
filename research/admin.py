from django.contrib import admin
from .models import Company, Executive, Office, ResearchHistory

class ExecutiveInline(admin.TabularInline):
    model = Executive
    extra = 1
    fields = ('order', 'position', 'name', 'name_kana')

class OfficeInline(admin.TabularInline):
    model = Office
    extra = 1
    fields = ('order', 'name', 'postal_code', 'address', 'phone', 'business_content')

@admin.register(Company)
class CompanyAdmin(admin.ModelAdmin):
    list_display = ('company_name', 'corporate_number', 'representative_name', 'industry', 'updated_at')
    list_filter = ('industry', 'stock_market', 'created_at')
    search_fields = ('company_name', 'corporate_number', 'representative_name')
    readonly_fields = ('created_at', 'updated_at')
    inlines = [ExecutiveInline, OfficeInline]
    
    fieldsets = (
        ('基本情報', {
            'fields': ('corporate_number', 'company_name', 'company_name_kana', 'english_name')
        }),
        ('代表者情報', {
            'fields': ('representative_name', 'representative_kana', 'representative_age', 
                      'representative_birth', 'representative_university')
        }),
        ('連絡先', {
            'fields': ('postal_code', 'address', 'phone', 'fax', 'url', 'registered_address')
        }),
        ('会社詳細', {
            'fields': ('founded', 'established', 'capital', 'investment', 'member_count', 
                      'union_member_count', 'stock_market', 'stock_code', 'fiscal_year_end')
        }),
        ('財務情報', {
            'fields': ('revenue', 'net_profit', 'deposits', 'employee_count', 'average_age', 
                      'average_salary', 'executive_count', 'shareholder_count', 'main_bank')
        }),
        ('事業情報', {
            'fields': ('industry', 'business_content', 'main_business', 'business_area', 
                      'group_affiliation', 'sales_destination', 'supplier')
        }),
        ('規模', {
            'fields': ('office_count', 'store_count')
        }),
        ('関連URL', {
            'fields': ('company_overview_url', 'office_list_url', 'organization_chart_url', 
                      'related_companies_url')
        }),
        ('システム情報', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )

@admin.register(Executive)
class ExecutiveAdmin(admin.ModelAdmin):
    list_display = ('company', 'position', 'name', 'order')
    list_filter = ('position',)
    search_fields = ('company__company_name', 'name', 'position')

@admin.register(Office)
class OfficeAdmin(admin.ModelAdmin):
    list_display = ('company', 'name', 'address', 'order')
    search_fields = ('company__company_name', 'name', 'address')

@admin.register(ResearchHistory)
class ResearchHistoryAdmin(admin.ModelAdmin):
    list_display = ('corporate_number', 'changed_field', 'timestamp')
    list_filter = ('changed_field', 'timestamp')
    search_fields = ('corporate_number', 'changed_field')
    readonly_fields = ('timestamp',)
