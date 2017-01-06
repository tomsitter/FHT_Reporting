import csv
from collections import defaultdict, namedtuple, OrderedDict
from tkinter.filedialog import askopenfilename
from tkinter import Tk
from datetime import datetime as dt
import os
from operator import attrgetter


def getfile(initialdir=None, title=None):
    Tk().withdraw()
    
    filename = askopenfilename(initialdir=initialdir or os.path.expanduser("~"),
                               title=title or 'Select report file')
    
    return filename


def get_patient_key(key, row):
    if isinstance(key, tuple):
        return ' '.join([row[e] for e in key]).title()
    else:
        if row[key] == '':
            return random_identifier()
        else:
            return row[key]


def build_dict(filename=None, key=None, lab_fields=None, encoding='utf-8-sig', date_format = '%m/%d/%Y'):
    """Given a filename to a CSV file, and which key and values we want to extract, build a dict"""
    
    
    labs = defaultdict(list)
    with open(filename, encoding=encoding) as f:
        reader = csv.DictReader(f)
        for row in reader:
            patient_key = get_patient_key(key, row)
            lab = Lab(date=dt.strptime(row[lab_fields.date], date_format).date(),
                      value=row[lab_fields.value])
            labs[patient_key].append(lab)
    return labs

def summarize(value, categories):
    for category in categories:
        

def main():
    
    dm_file = getfile()
    Lab = namedtuple('Lab', ['date', 'value'])
    lab_fields = Lab(date='Collection Date', value='Normalized')
    
    labs = build_dict(filename=dm_file, key='PHN', lab_fields=lab_fields)
    
    categories = OrderedDict([('<6.5', 0), ('6.5-7', 0), ('7-7.5', 0), ('>7.5', 0)])
    
    for patient_labs in labs.values():
        # sorted_labs = sorted(patient_labs, key=attrgetter('date'))
        most_recent = max(patient_labs, key=attrgetter('date'))
        
        try:
            value = float(most_recent.value)
            if value < 6.5:
                categories['<6.5'] += 1
            elif value < 7:
                categories['6.5-7'] += 1
            elif value < 7.5:
                categories['7-7.5']+= 1
            else:
                categories['>7.5'] += 1
        except ValueError as e:
            print(e)
           
    
    for key, value in categories.items():
        print(key, ":", value)
        
        
        