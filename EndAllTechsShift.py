import sys
import os
import django
from django.utils import timezone

from Variables import *
from django import utils

#*****************************
sys.path.append(project_location)#folder for the app
#*****************************
os.environ['DJANGO_SETTINGS_MODULE'] = 'testwebpage.settings'
django.setup()
from django.contrib.auth.models import User
from prodfloor.models import Features,Info, Stops
from django.contrib.admin.models import LogEntry, ADDITION,CHANGE
from django.contrib.contenttypes.models import ContentType
from django.contrib.auth import logout

techs = User.objects.filter(groups__name='Technicians')


def EndShift(tech):
    tech_name = tech.first_name + ' ' + tech.last_name
    jobs = Info.objects.filter(Tech_name= tech_name).exclude(status='Complete').exclude(status='Reassigned')
    for obj in jobs:
        ID = obj.id
        po = obj.po
        stop_reason = 'Shift ended'
        time = timezone.now()
        description = 'The user ' + tech_name + ' ended his shift'
        if obj.status == 'Stopped':
            stops = Stops.objects.filter(info_id=obj.id)
            if any('Shift ended' in stop.reason and 'Not available yet' in stop.solution for stop in stops):
                pass
            else:
                stop = Stops(info_id=ID, reason=stop_reason, solution='Not available yet', extra_cause_1='N/A',
                             extra_cause_2='N/A', stop_start_time=time, stop_end_time=time,
                             reason_description=description, po=po)
                stop.save()
                obj.save()
                ct = ContentType.objects.get_for_model(obj)
                l = LogEntry.objects.log_action(
                    user_id=tech.pk,
                    content_type_id=ct.pk,
                    object_id=obj.pk,
                    object_repr=str(obj.po),
                    action_flag=CHANGE,
                    change_message='Job was stopped due to a shift end'
                )
                l.save()
        else:
            if obj.status != 'Reassigned':
                obj.prev_stage = obj.status
                obj.save()
            stop = Stops(info_id=ID, reason=stop_reason, solution='Not available yet', extra_cause_1='N/A',
                         extra_cause_2='N/A', stop_start_time=time, stop_end_time=time,
                         reason_description=description, po=po)
            stop.save()
            obj.status = 'Stopped'
            obj.save()
            ct = ContentType.objects.get_for_model(obj)
            l = LogEntry.objects.log_action(
                user_id=tech.pk,
                content_type_id=ct.pk,
                object_id=obj.pk,
                object_repr=str(obj.po),
                action_flag=CHANGE,
                change_message='Job was stopped due to a shift end'
            )
            l.save()

for tech in techs:
    EndShift(tech)