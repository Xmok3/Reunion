from django.contrib import admin
from .models import Guest
from django.utils.html import format_html
import openpyxl
from django.http import HttpResponse

@admin.register(Guest)
class GuestAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'email', 'phone_number', 'qr_code_value', 'qr_thumbnail')
    search_fields = ('full_name', 'email', 'phone_number', 'qr_code_value')

    def qr_thumbnail(self, obj):
        if obj.qr_image:
            return format_html('<img src="{}" width="50" />', obj.qr_image.url)
        return "-"
    qr_thumbnail.short_description = "QR Code"

    # Export to Excel button
    actions = ['export_as_excel']

    def export_as_excel(self, request, queryset):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Guests'
        sheet.append(['Full Name', 'Email', 'Phone Number', 'QR Code'])
        for guest in queryset:
            sheet.append([guest.full_name, guest.email, guest.phone_number, guest.qr_code_value])
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=guests.xlsx'
        workbook.save(response)
        return response
    export_as_excel.short_description = "Export Selected Guests to Excel"
