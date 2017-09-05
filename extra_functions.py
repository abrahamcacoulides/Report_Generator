import pytz
from django.contrib.admin.models import LogEntry, CHANGE, ADDITION
from django.contrib.contenttypes.models import ContentType
from django.utils import timezone
from prodfloor.models import Stops, Times, Features, Info
import datetime,copy

instance_time_zone = pytz.timezone('America/Monterrey')

def spentTime(pk,number):
    now = timezone.now()
    end = 0
    start = 0
    time_on_shift_end = timezone.timedelta(0)
    times = Times.objects.get(info_id=pk)
    job = Info.objects.get(pk=pk)
    status = job.status
    if status == 'Stoppped':
        status = job.prev_stage
    status_list = ['','Beginning','Program','Logic','Ending']
    stops_shift_end = Stops.objects.filter(info_id=pk,reason='Shift ended')
    if number == 1:
        start = times.start_time_1
        end = times.end_time_1
    elif number == 2:
        start = times.start_time_2
        end = times.end_time_2
    elif number == 3:
        start = times.start_time_3
        end = times.end_time_3
    elif number == 4:
        start = times.start_time_4
        end = times.end_time_4
    else:
        pass
    for stop in stops_shift_end:
        if stop.stop_start_time > start and stop.stop_end_time<end:#el inicio del stop debe de ser mayor quue el inicio del stage y el
            if stop.stop_start_time == stop.stop_end_time: #job is in shift end stop
                time_on_shift_end += now - stop.stop_start_time
            else:
                time_on_shift_end += stop.stop_end_time - stop.stop_start_time
    if end>start:#means that the stop_time has been set
        elapsed_time = (end-start)-time_on_shift_end
        return elapsed_time
    elif end == start and status != status_list[number]:#means that it has not being started
        elapsed_time = datetime.timedelta(0)
        return elapsed_time
    else:#job is being worked on
        elapsed_time = (now - start)-time_on_shift_end
        return elapsed_time

def stopsnumber(pk):
    stops = Stops.objects.filter(info_id=pk)
    numer_of_stops = 0
    for stop in stops:
        numer_of_stops+=1
    return numer_of_stops

def timeonstop(pk):
    now = timezone.now()
    stops_shiftend = Stops.objects.filter(info_id=pk).filter(reason='Shift ended')
    stops = Stops.objects.filter(info_id=pk).exclude(reason='Shift ended')
    timeinstop = timezone.timedelta(0)
    for stop in stops:
        if stops_shiftend:
            real_time_on_stop = stop.stop_end_time - stop.stop_start_time
            for stop_se in stops_shiftend:
                if stop.stop_start_time < stop_se.stop_start_time and stop.stop_end_time >= stop_se.stop_end_time:
                    shift_end_stop_elapsed_time = stop_se.stop_end_time - stop_se.stop_start_time
                    real_time_on_stop -= shift_end_stop_elapsed_time
            timeinstop += real_time_on_stop
        else:
            if stop.stop_end_time>stop.stop_start_time:
                timeinstop += stop.stop_end_time - stop.stop_start_time
            else:
                timeinstop += now - stop.stop_start_time
    return timeinstop

def timeonstop_1(pk):
    now = timezone.now()
    stop_obj = Stops.objects.get(pk=pk)
    job_obj = Info.objects.get(pk=stop_obj.info_id)
    stops_shiftend = Stops.objects.filter(info_id=job_obj.pk).filter(reason='Shift ended')
    timeinstop = timezone.timedelta(0)
    if stops_shiftend:
        real_time_on_stop = stop_obj.stop_end_time - stop_obj.stop_start_time
        for stop_se in stops_shiftend:
            if stop_obj.stop_start_time < stop_se.stop_start_time and stop_obj.stop_end_time >= stop_se.stop_end_time:
                shift_end_stop_elapsed_time = stop_se.stop_end_time - stop_se.stop_start_time
                real_time_on_stop -= shift_end_stop_elapsed_time
        timeinstop += real_time_on_stop
    else:
        if stop_obj.stop_end_time>stop_obj.stop_start_time:
            timeinstop += stop_obj.stop_end_time - stop_obj.stop_start_time
        else:
            timeinstop += now - stop_obj.stop_start_time
    return timeinstop

def totaltime(pk):
    now = timezone.now()
    end = 0
    start = 0
    times = Times.objects.get(info_id=pk)
    elapsed_time = datetime.timedelta(0)
    number = 1
    while number < 5:
        if number == 1:
            start = times.start_time_1
            end = times.end_time_1
        elif number == 2:
            start = times.start_time_2
            end = times.end_time_2
        elif number == 3:
            start = times.start_time_3
            end = times.end_time_3
        elif number == 4:
            start = times.start_time_4
            end = times.end_time_4
        if end > start:  # means that the stop_time has been set
            elapsed_time += (end - start)
        elif end == start and number != 1:  # means that it has not being started and is not in beginning stage
            elapsed_time += datetime.timedelta(0)
        else:  # job is being worked on
            elapsed_time += (now - start)
        number+=1
    return elapsed_time

def effectivetime(pk):#effective time spent on job
    now = timezone.now()
    stops_end_shift = Stops.objects.filter(info_id=pk, reason='Shift ended')
    reassign_stops = Stops.objects.filter(info_id=pk, reason='Job reassignment')
    stops = Stops.objects.filter(info_id=pk).exclude(reason__in=['Shift ended','Job reassignment'])
    stops_1 = Stops.objects.filter(info_id=pk).exclude(reason__in=['Shift ended',])
    timeinstop = timezone.timedelta(0)
    end = 0
    start = 0
    times = Times.objects.get(info_id=pk)
    elapsed_time = datetime.timedelta(0)
    number = 1
    while number < 5:
        if number == 1:
            start = times.start_time_1
            end = times.end_time_1
        elif number == 2:
            start = times.start_time_2
            end = times.end_time_2
        elif number == 3:
            start = times.start_time_3
            end = times.end_time_3
        elif number == 4:
            start = times.start_time_4
            end = times.end_time_4
        if end > start:  # means that the stop_time has been set
            elapsed_time += (end - start)
        elif end == start:  # means that it has not being started and is not in beginning stage
            elapsed_time += datetime.timedelta(0)
        else:  # job is being worked on
            elapsed_time += (now - start)
        number += 1
    for stop in stops:
        if stop.stop_end_time > stop.stop_start_time:
            timeinstop += stop.stop_end_time - stop.stop_start_time
        else:
            timeinstop += now - stop.stop_start_time
    for re_stop in reassign_stops:
        print('here')
        inside_another_stop = False
        if re_stop.stop_end_time > re_stop.stop_start_time:
            for stop in stops:
                if re_stop.stop_start_time >= stop.stop_start_time and re_stop.stop_end_time <= stop.stop_end_time:
                    inside_another_stop = True
        else:
            for stop in stops:
                if not (stop.stop_end_time >= stop.stop_start_time):
                    inside_another_stop = True
        if not inside_another_stop:
            if re_stop.stop_end_time > re_stop.stop_start_time:
                timeinstop += re_stop.stop_end_time - re_stop.stop_start_time
            else:
                timeinstop += now - re_stop.stop_start_time
    for es_stop in stops_end_shift:
        inside_another_stop = False
        if es_stop.stop_end_time > es_stop.stop_start_time:
            for stop in stops_1:
                if es_stop.stop_start_time > stop.stop_start_time and es_stop.stop_end_time <= stop.stop_end_time:
                    inside_another_stop = True
        else:
            for stop in stops_1:
                if not (stop.stop_end_time > stop.stop_start_time):
                    inside_another_stop = True
        if not inside_another_stop:
            if es_stop.stop_end_time > es_stop.stop_start_time:
                timeinstop += es_stop.stop_end_time - es_stop.stop_start_time
            else:
                timeinstop += now - es_stop.stop_start_time
    job = Info.objects.get(pk=pk)
    if job.status == 'Stopped':
        if any(stop.solution == 'Not available yet' for stop in reassign_stops):
            return datetime.timedelta(0)
    eff_time = elapsed_time - timeinstop
    return (eff_time)

def gettech(pk,*args, **kwargs):
    info = Info.objects.get(pk=pk)
    tech = info.Tech_name
    return tech

def categories(pk,*args, **kwargs):
    features_in_job = Features.objects.filter(info_id = pk)
    job = Info.objects.get(pk = pk)
    job_type = job.job_type
    element_dict = {
        'level_2': [],
        'level_3': [],
        'level_4': [],
        'level_5': [],
        'level_6': []}#pending
    m2000_dict = {
        'level_2': ['REAR', 'DUP', 'MOD', '2STARTERS', 'SHC', 'EMCO', 'R6'],
        'level_3': ['mView', 'iMon', 'LOC'],
        'level_4': ['MANUAL', 'OVL'],
        'level_5': ['CUST', 'MRL'],
        'level_6': ['TSSA']}
    m4000_dict = {
        'level_2': ['REAR', 'DUP', 'MOD', '2STARTERS'],
        'level_3': ['mView', 'iMon', 'LOC', 'SHORTF'],
        'level_4': ['MANUAL', 'OVL'],
        'level_5': ['CUST'],
        'level_6': ['TSSA']}
    if job_type == '2000':
        dict = copy.deepcopy(m2000_dict)
    elif job_type == '4000':
        dict = copy.deepcopy(m4000_dict)
    else:
        dict = copy.deepcopy(element_dict)
    if any(feature.features in dict['level_6'] for feature in features_in_job):
        category = 6
    elif any(feature.features in dict['level_5'] for feature in features_in_job):
        category = 5
    elif any(feature.features in dict['level_4'] for feature in features_in_job):
        category = 4
    elif any(feature.features in dict['level_3'] for feature in features_in_job):
        category = 3
    elif any(feature.features in dict['level_2'] for feature in features_in_job):
        category = 2
    else:
        category = 1
    return category

def gettimes(pk,B,*args, **kwargs):
    times = Times.objects.get(info_id=pk)
    if B == 'start':
        return datetime.date.__format__(times.start_time_1.astimezone(instance_time_zone),"%m/%d/%Y")
    elif B == 'end':
        if times.start_time_1 == times.end_time_4:
            return '-'
        else:
            return datetime.date.__format__(times.end_time_4.astimezone(instance_time_zone),"%m/%d/%Y")
    else:
        return 'N/A'

def efficiency(A):
    count = 0
    eff = 0
    for obj in A:
        category = categories(obj.pk)
        times_per_category = {
            '2000':
                {1: 4 * 60,
                 2: 4.5 * 60,
                 3: 5 * 60,
                 4: 6 * 60,
                 5: 7 * 60,
                 6: 8 * 60, },
            '4000':
                {1: 7 * 60,
                 2: 8 * 60,
                 3: 10 * 60,
                 4: 13 * 60,
                 5: 15 * 60,
                 6: 16 * 60, },
            'ELEM':
                {1: 3 * 60, },
        }
        expected_time = times_per_category[obj.job_type][category]
        jobs_for_this_po = Info.objects.filter(po=obj.po)
        eff_time_total = datetime.timedelta(0)
        for job in jobs_for_this_po:
            eff_time_total += effectivetime(job.pk)
        Efficiency = (expected_time / ((eff_time_total.seconds) / 60))
        count += 1
        eff += Efficiency
    if count<1:
        return '-'
    else:
        return(eff / count)
