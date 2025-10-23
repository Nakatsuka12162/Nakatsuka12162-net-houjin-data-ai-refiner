from django.contrib import admin
from django.contrib.admin import AdminSite
from django.utils.html import format_html
from .models import Company, Executive, Office, ResearchHistory

class CustomAdminSite(AdminSite):
    site_header = '🏢 Company Research System'
    site_title = 'Company Research'
    index_title = '📊 Dashboard'
    site_url = None  # Remove "View Site" link
    
    def index(self, request, extra_context=None):
        extra_context = extra_context or {}
        extra_context.update({
            'total_companies': Company.objects.count(),
            'total_executives': Executive.objects.count(),
            'total_offices': Office.objects.count(),
            'recent_changes': ResearchHistory.objects.order_by('-timestamp')[:5],
        })
        return super().index(request, extra_context)

# Create custom admin site instance
custom_admin_site = CustomAdminSite(name='custom_admin')

class ExecutiveInline(admin.TabularInline):
    model = Executive
    extra = 0
    fields = ('order', 'position', 'name', 'name_kana')

class OfficeInline(admin.TabularInline):
    model = Office
    extra = 0
    fields = ('order', 'name', 'postal_code', 'address', 'phone', 'business_content')

class CompanyAdmin(admin.ModelAdmin):
    list_display = ('company_badge', 'corporate_number', 'representative_name', 'industry_badge', 'updated_at')
    list_filter = ('industry', 'stock_market', 'created_at')
    search_fields = ('company_name', 'corporate_number', 'representative_name')
    readonly_fields = ('created_at', 'updated_at')
    inlines = [ExecutiveInline, OfficeInline]
    list_per_page = 50
    
    def company_badge(self, obj):
        return format_html('<strong style="color: #2196F3;">🏢 {}</strong>', obj.company_name)
    company_badge.short_description = '会社名'
    
    def industry_badge(self, obj):
        colors = {'製造業': '#4CAF50', '小売業': '#FF9800', 'サービス業': '#2196F3'}
        color = colors.get(obj.industry, '#757575')
        return format_html(
            '<span style="background: {}; color: white; padding: 3px 8px; border-radius: 3px;">{}</span>',
            color, obj.industry or '-'
        )
    industry_badge.short_description = '業種'
    
    fieldsets = (
        ('🏢 基本情報', {'fields': ('corporate_number', 'company_name', 'company_name_kana', 'english_name')}),
        ('👤 代表者情報', {'fields': ('representative_name', 'representative_kana', 'representative_age', 'representative_birth', 'representative_university'), 'classes': ('collapse',)}),
        ('📞 連絡先', {'fields': ('postal_code', 'address', 'phone', 'fax', 'url', 'registered_address'), 'classes': ('collapse',)}),
        ('🏛️ 会社詳細', {'fields': ('founded', 'established', 'capital', 'investment', 'member_count', 'union_member_count', 'stock_market', 'stock_code', 'fiscal_year_end'), 'classes': ('collapse',)}),
        ('💰 財務情報', {'fields': ('revenue', 'net_profit', 'deposits', 'employee_count', 'average_age', 'average_salary', 'executive_count', 'shareholder_count', 'main_bank'), 'classes': ('collapse',)}),
        ('💼 事業情報', {'fields': ('industry', 'business_content', 'main_business', 'business_area', 'group_affiliation', 'sales_destination', 'supplier'), 'classes': ('collapse',)}),
        ('📍 規模', {'fields': ('office_count', 'store_count'), 'classes': ('collapse',)}),
        ('🔗 関連URL', {'fields': ('company_overview_url', 'office_list_url', 'organization_chart_url', 'related_companies_url'), 'classes': ('collapse',)}),
        ('⚙️ システム情報', {'fields': ('created_at', 'updated_at'), 'classes': ('collapse',)}),
    )

class ExecutiveAdmin(admin.ModelAdmin):
    list_display = ('executive_badge', 'company_link', 'position', 'order')
    list_filter = ('position',)
    search_fields = ('company__company_name', 'name', 'position')
    
    def executive_badge(self, obj):
        return format_html('<strong style="color: #2196F3;">👤 {}</strong>', obj.name)
    executive_badge.short_description = '役員名'
    
    def company_link(self, obj):
        return format_html('<a href="../company/{}/change/">🏢 {}</a>', obj.company.id, obj.company.company_name)
    company_link.short_description = '会社'

class OfficeAdmin(admin.ModelAdmin):
    list_display = ('office_badge', 'company_link', 'address_short', 'order')
    search_fields = ('company__company_name', 'name', 'address')
    
    def office_badge(self, obj):
        return format_html('<strong style="color: #4CAF50;">📍 {}</strong>', obj.name)
    office_badge.short_description = '事業所名'
    
    def company_link(self, obj):
        return format_html('<a href="../company/{}/change/">🏢 {}</a>', obj.company.id, obj.company.company_name)
    company_link.short_description = '会社'
    
    def address_short(self, obj):
        return obj.address[:50] + '...' if len(obj.address) > 50 else obj.address
    address_short.short_description = '住所'

class ResearchHistoryAdmin(admin.ModelAdmin):
    list_display = ('timestamp_badge', 'corporate_number', 'changed_field', 'change_preview')
    list_filter = ('changed_field', 'timestamp')
    search_fields = ('corporate_number', 'changed_field')
    readonly_fields = ('timestamp',)
    date_hierarchy = 'timestamp'
    
    def timestamp_badge(self, obj):
        return format_html('<span style="color: #2196F3;">🕒 {}</span>', obj.timestamp.strftime('%Y-%m-%d %H:%M'))
    timestamp_badge.short_description = '日時'
    
    def change_preview(self, obj):
        old = obj.old_value[:30] + '...' if len(obj.old_value) > 30 else obj.old_value
        new = obj.new_value[:30] + '...' if len(obj.new_value) > 30 else obj.new_value
        return format_html('<span style="color: #f44336;">{}</span> → <span style="color: #4CAF50;">{}</span>', old, new)
    change_preview.short_description = '変更内容'

# Register models with custom admin site
custom_admin_site.register(Company, CompanyAdmin)
custom_admin_site.register(Executive, ExecutiveAdmin)
custom_admin_site.register(Office, OfficeAdmin)
custom_admin_site.register(ResearchHistory, ResearchHistoryAdmin)
