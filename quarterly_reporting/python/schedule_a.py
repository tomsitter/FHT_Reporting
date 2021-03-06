import csv
import os
from datetime import datetime
from collections import defaultdict, namedtuple
from operator import attrgetter
import pdb
import random
import string
from util import getfile
import numbers

__author__ = 'Tom Sitter'

Event = namedtuple('Event', ['date', 'value'])

date_format = '%m/%d/%Y'


filepath = r'C:\Users\SAFHT_AdminTom\Documents\PHI'

def calculate_change(events, score_type=int):
    """Given a dict of patients with a list of Event tuples as values, return the difference between the
       most recent and least recent scores"""
    changes = {} 
    
    for patient, scores in events.items():
        sorted_scores = sorted(scores, key=attrgetter('date'))
        if len(sorted_scores) == 1:
            changes[patient] = 'N/A'
        else:
            changes[patient] = score_type(sorted_scores[-1].value) - score_type(sorted_scores[0].value)
        
    return changes


def build_dict(filename=None, key=None, event_fields=None, encoding='utf-8-sig'):
    """Given a filename to a CSV file, and which key and values we want to extract, build a dict"""
    events = defaultdict(list)
    with open(filename, encoding=encoding) as f:
        reader = csv.DictReader(f)
        for row in reader:
            patient_key = get_patient_key(key, row)
            event = Event(date=datetime.strptime(row[event_fields.date], date_format).date(),
                          value=row[event_fields.value])
            events[patient_key].append(event)
    return events


def filter_events_by_code(events, init_code=None):
    """Given a list of events and an initial billing code,
    Return a sorted list of all billing events that occurred from the
    date of the most recent initial billing onward"""
    
    # Sort events by date and get the most recent initial visit date
    sorted_event_dates = sorted([event.date for event in events if event.value == init_code])
    
    if not sorted_event_dates:
        return None, []
    else:
        init_date = sorted_event_dates[-1]
        
    # Make sure patient has had a initial visit billed
    if not init_date:
        return None, []
        
    # Now go through all events and pull values for any dates occurring on or after the initial visit date
    return (init_date, [event
                        for event in sorted(events, key=attrgetter('date'))
                        if event.date >= init_date]
            )


def filter_events_by_date(events, init_date=None):
    return [event
            for event in sorted(events, key=attrgetter('date'))
            if event.date >= init_date]


def get_patient_status(start, end):
    if not isinstance(end, numbers.Real) or not isinstance(start, numbers.Real):
        return 'NA'
    elif end < start:
        return 'Improved'
    elif end == start:
        return 'Same'
    elif end > start:
        return 'Worsened'


def get_patient_key(key, row):
    if isinstance(key, tuple):
        return ' '.join([row[e] for e in key]).title()
    else:
        if row[key] == '':
            return random_identifier()
        else:
            return row[key]


def random_identifier():
    print('Generating Random identifier')
    return ''.join(
        [random.choice(string.digits) for _ in range(10)] +
        [random.choice(string.ascii_uppercase) for _ in range(2)]
    )


def weight_mgmt_change():
    billing_file = 'WM_Weight_20160401-20161231.csv'
    weights = build_dict(filename=os.path.join(filepath, billing_file),
                         key=('First Name', 'Last Name'),
                         event_fields=Event(date='Collection Date', value='Value'))
    for patient, events in weights.items():
        sorted_dates = sorted([event.date for event in events])
        if len(sorted_dates) > 1:
            start_date, end_date = sorted_dates[0], sorted_dates[-1]
            start = float([event.value for event in events if event.date == start_date][0])
            end = float([event.value for event in events if event.date == end_date][0])
            print(patient, ":", round((end-start)/start*100, 2), "%")
        else:
            print(patient, ": Only 1 weight")

def weight_mgmt_appt_type():
    
    billing_file = 'WM_AllVisits_20160401-20161231.csv'
    
    visits = build_dict(filename=os.path.join(filepath, billing_file),
                       key=('First Name', 'Last Name'),
                       event_fields=Event(date='Bill Date', value='PCode'))
    
    summary = {
        'only_INI': [],
        'WFU_WGS': [],
    }
    for patient, events in visits.items():
        only_INI = all(event.value=='WINI' for event in events)
        if only_INI:
            summary['only_INI'].append(events)
        else:
            summary['WFU_WGS'].append(events)

def mental_health(billing_file='MH_Billing_20160401-20160930.csv',
                  phq9_file = 'MH_PHQ9s_Q1-Q3_2016.csv',
                  gad7_file = 'MH_GAD7s_Q1-Q3_2016.csv',
                  fields=('Bill Date', 'PCode'),
                  init_code='MHINI',
                  event_filter=filter_events_by_code):
        
    bills = build_dict(filename=os.path.join(filepath, billing_file),
                       key='PHN',
                       event_fields=Event(date=fields[0], value=fields[1]))
    
    phq9s = build_dict(filename=os.path.join(filepath, phq9_file),
                       key='PHN',
                       event_fields=Event(date='Observation Date', value='Value'))
    
    gad7s = build_dict(filename=os.path.join(filepath, gad7_file),
                       key='PHN',
                       event_fields=Event(date='Observation Date', value='Value'))
    
    scores = defaultdict(dict)
    
    for patient, events in bills.items():
        
        init_date, filtered_events = event_filter(events, init_code=init_code)
        
        if not init_date:
            # print('Patient: ', patient, 'Status: N/A -- No Init Date')
            continue
            
        filtered_phq9_events = filter_events_by_date(phq9s[patient],
                                                     init_date=init_date)
        filtered_gad7_events = filter_events_by_date(gad7s[patient],
                                                     init_date=init_date)
        
        if len(filtered_phq9_events) >= 2:
            init_phq9, final_phq9 = filtered_phq9_events[0], filtered_phq9_events[-1]
            # print('Patient: ', patient, 'PHQ9 Status: ', get_patient_status(int(init_phq9.value),
                                                                            # int(final_phq9.value)))
            scores[patient]['phq9'] = get_patient_status(int(init_phq9.value), int(final_phq9.value))
            
        if len(filtered_gad7_events) >= 2:
            init_gad7, final_gad7 = filtered_gad7_events[0], filtered_gad7_events[-1]
            # print('Patient: ', patient, 'GAD7 Status: ', get_patient_status(int(init_gad7.value),
                                                                            # int(final_gad7.value)))
            scores[patient]['gad7'] = get_patient_status(int(init_gad7.value), int(final_gad7.value))
            
    return scores

    
def score_changes(billing_file=None,
                  key=('First Name', 'Last Name'),
                  event_fields=Event(date='Collection Date', value='Value'),
                  improvement='decrease'):

    
    if not billing_file:
        billing_file = getfile()   
    scores = build_dict(billing_file, key, event_fields)
        
    changes = calculate_change(scores, score_type=float)
    
    for pat, change in changes.items():
        if improvement == 'decrease':
            print(pat, get_patient_status(0, change))
        else:
            print(pat, get_patient_status(change, 0))
        
    return changes
            
                    
