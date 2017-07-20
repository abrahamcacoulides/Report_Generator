import sys
import os
import django
import datetime
import openpyxl
from openpyxl.styles import PatternFill

sys.path.append('C:\\Users\\abraham.cacoulides\\PycharmProjects\\testwebpage\\')
os.environ['DJANGO_SETTINGS_MODULE'] = 'testwebpage.settings'
django.setup()

from prodfloor.models import Info,Times,Stops
from extra_functions import *
grayFill = PatternFill(start_color='CBFCC4',end_color = 'CBFCC4',fill_type = 'solid')

wb = openpyxl.load_workbook(r'C:\Users\abraham.cacoulides\Desktop\Morning Reports\Morning_Report_Template.xlsx')
data_wb = openpyxl.load_workbook(r'C:\Users\abraham.cacoulides\Desktop\Morning Reports\Data.xlsx')
sheet = wb.get_sheet_by_name('Jobs_Completed')
data_sheet = data_wb.get_sheet_by_name('Data')
job = Info.objects.all()
day = datetime.date(2017,7,13)
completed_before = datetime.datetime.combine(day, datetime.datetime.max.time())
completed_after = datetime.datetime.combine(day, datetime.datetime.min.time())
times = Times.objects.filter(end_time_4__gte=completed_after,end_time_4__lte=completed_before)
times_pks=[]
for i in times:
    times_pks.append(i.info.pk)
job = job.filter(pk__in=times_pks).filter(status='Complete')
completed_pks = []
for i in job:
    completed_pks.append(i.po)
completed_jobs_number = len(completed_pks)
completed_jobs = Info.objects.filter(po__in=completed_pks).order_by('po')#might not be needed
completed_jobs_m4000 = job.filter(job_type = '4000')
completed_jobs_m2000 = job.filter(job_type = '2000')
completed_jobs_elem = job.filter(job_type = 'elem')
c = 2
for obj_main in job:
    stations_dict = {'0': '-----',
                     '1': 'S1',
                     '2': 'S2',
                     '3': 'S3',
                     '4': 'S4',
                     '5': 'S5',
                     '6': 'S6',
                     '7': 'S7',
                     '8': 'S8',
                     '9': 'S9',
                     '10': 'S10',
                     '11': 'S11',
                     '12': 'S12',
                     '13': 'ELEM1',
                     '14': 'ELEM2'}
    job_type_dict = {'2000': 'M2000',
                     '4000': 'M4000',
                     'ELEM': 'Element'}
    all_jobs = Info.objects.filter(po=obj_main.po).order_by('po')
    total_elapsed_time = datetime.timedelta(0)
    total_stops = 0
    total_time_on_stop = datetime.timedelta(0)
    total_eff_time = datetime.timedelta(0)
    for obj in all_jobs:
        start = gettimes(obj.pk, 'start')
        end = gettimes(obj.pk, 'end')
        beginning_time = spentTime(obj.pk, 1)
        program_time = spentTime(obj.pk, 2)
        logic_time = spentTime(obj.pk, 3)
        ending_time = spentTime(obj.pk, 4)
        number_of_stops = stopsnumber(obj.pk)
        time_on_stop = timeonstop(obj.pk)
        elapsed_time = totaltime(obj.pk)
        eff_time = effectivetime(obj.pk)
        tech = gettech(obj.pk)
        category = categories(obj.pk)
        total_stops += number_of_stops
        total_elapsed_time += elapsed_time
        total_time_on_stop += time_on_stop
        total_eff_time += eff_time
        sheet['A' + str(c)] = obj.job_num
        sheet['B' + str(c)] = obj.po
        sheet['C' + str(c)] = job_type_dict[obj.job_type]
        sheet['D' + str(c)] = obj.status
        sheet['E' + str(c)] = stations_dict[obj.station]
        sheet['F' + str(c)] = str(start)
        sheet['G' + str(c)] = str(end)
        sheet['H' + str(c)] = beginning_time
        sheet['I' + str(c)] = program_time
        sheet['J' + str(c)] = logic_time
        sheet['K' + str(c)] = ending_time
        sheet['L' + str(c)] = str(elapsed_time).split('.', 2)[0]
        sheet['M' + str(c)] = number_of_stops
        sheet['N' + str(c)] = str(time_on_stop).split('.', 2)[0]
        sheet['O' + str(c)] = str(eff_time).split('.', 2)[0]
        sheet['P' + str(c)] = category
        sheet['Q' + str(c)] = tech
        c += 1
    sheet['A' + str(c)] = obj_main.job_num
    sheet['B' + str(c)] = obj_main.po
    sheet['C' + str(c)] = job_type_dict[obj_main.job_type]
    sheet['D' + str(c)] = obj_main.status
    sheet['E' + str(c)] = '-'
    sheet['F' + str(c)] = '-'
    sheet['G' + str(c)] = '-'
    sheet['H' + str(c)] = '-'
    sheet['I' + str(c)] = '-'
    sheet['J' + str(c)] = '-'
    sheet['K' + str(c)] = '-'
    sheet['L' + str(c)] = str(total_elapsed_time).split('.', 2)[0]
    sheet['M' + str(c)] = str(total_stops).split('.', 2)[0]
    sheet['N' + str(c)] = str(total_time_on_stop).split('.', 2)[0]
    sheet['O' + str(c)] = str(total_eff_time).split('.', 2)[0]
    sheet['P' + str(c)] = categories(obj_main.pk)
    sheet['Q' + str(c)] = '-'
    for rowOfCellObjects in sheet['A' + str(c):'Q' + str(c)]:
        for cellObj in rowOfCellObjects:
            cellObj.fill = grayFill
    c+=1

efficiency_2000 = efficiency(completed_jobs_m2000)
efficiency_4000 = efficiency(completed_jobs_m4000)
efficiency_elem = efficiency(completed_jobs_elem)
data_sheet['C3']= completed_jobs_m4000.count()
data_sheet['C6']= completed_jobs_m2000.count()
data_sheet['C9']= completed_jobs_elem.count()
data_sheet['B3']= efficiency_4000
data_sheet['B6']= efficiency_2000
data_sheet['B9']= efficiency_elem

wb.save(r'C:\Users\abraham.cacoulides\Desktop\Morning Reports\Morning_Report_'+ str(day) + '.xlsx')
data_wb.save(r'C:\Users\abraham.cacoulides\Desktop\Morning Reports\Data.xlsx')
