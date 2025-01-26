import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

def create_calendar(year):
    wb = Workbook()
    months = {
        1: 'Januar', 2: 'Februar', 3: 'März', 4: 'April',
        5: 'Mai', 6: 'Juni', 7: 'Juli', 8: 'August',
        9: 'September', 10: 'Oktober', 11: 'November', 12: 'Dezember'
    }
    
    weekdays = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    thick_right_border = Border(left=Side(style='thin'), right=Side(style='thick'), top=Side(style='thin'), bottom=Side(style='thin'))
    bold_font = Font(bold=True)
    
    wb.remove(wb.active)
    
    for month_num in range(1, 13):
        ws = wb.create_sheet(months[month_num])
        
        dv = DataValidation(
            type="custom",
            formula1='=COUNTIF($D3:$AG3,"u")+COUNTIF($D3:$AG3,"x")<=8',
            error="Für diesen Tag ist das Kontingent bereits aufgebraucht. Kein Eintrag mehr möglich.",
            errorTitle="Kein Eintrag möglich",
            allow_blank=True,
            showErrorMessage=True
        )
        ws.add_data_validation(dv)
        
        ws['A1'] = f"{months[month_num]}.{str(year)[-2:]}"
        ws.merge_cells('A1:B1')
        ws['A1'].alignment = Alignment(horizontal='center')
        ws['A1'].font = bold_font
        
        ws['B2'] = "Wochentag"
        ws['B2'].alignment = Alignment(horizontal='center')
        ws['B2'].font = bold_font
        
        ws['C2'] = months[month_num]
        ws['C2'].alignment = Alignment(horizontal='center')
        ws['C2'].font = bold_font
        
        first_day = datetime(year, month_num, 1)
        last_day = datetime(year, month_num + 1, 1) - timedelta(days=1) if month_num < 12 else datetime(year, 12, 31)
        dates = pd.date_range(first_day, last_day)
        
        row = 3
        kw_start_row = None
        current_kw = None
        
        for i, date in enumerate(dates):
            new_kw = date.isocalendar()[1]
            
            if i == 0 or date.weekday() == 0:
                if current_kw != new_kw:
                    if kw_start_row and kw_start_row < row - 1:
                        ws.merge_cells(f'A{kw_start_row}:A{row-1}')
                        ws[f'A{kw_start_row}'].alignment = Alignment(horizontal='center', vertical='center')
                        ws[f'A{kw_start_row}'].font = bold_font
                    
                    ws[f'A{row}'] = f'KW {new_kw}'
                    current_kw = new_kw
                    kw_start_row = row
            
            ws[f'B{row}'] = weekdays[date.weekday()]
            ws[f'C{row}'] = date.day
            ws[f'C{row}'].alignment = Alignment(horizontal='center')
            
            dv.add(f'D{row}:AG{row}')
            
            ws[f'B{row}'].font = bold_font
            ws[f'C{row}'].font = bold_font
            
            ws[f'A{row}'].border = thin_border
            ws[f'B{row}'].border = thin_border
            ws[f'C{row}'].border = thick_right_border
            
            row += 1
        
        if kw_start_row and kw_start_row < row - 1:
            ws.merge_cells(f'A{kw_start_row}:A{row-1}')
            ws[f'A{kw_start_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{kw_start_row}'].font = bold_font
        
        ws['A1'].border = thin_border
        ws['B1'].border = thin_border
        ws['A2'].border = thin_border
        ws['B2'].border = thin_border
        ws['C2'].border = thick_right_border
        
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 8

    wb.save(f'Kalender_{year}.xlsx')

if __name__ == '__main__':
    year = int(input("Bitte geben Sie das Jahr ein: "))
    create_calendar(year)