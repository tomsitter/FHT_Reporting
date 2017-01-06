from collections import defaultdict, namedtuple, OrderedDict
import csv
from datetime import datetime as dt
from operator import attrgetter
import os
from tkinter.filedialog import askopenfilename
from tkinter import Tk

# Standardized tuple for storing observation date / value pairs
Observation = namedtuple('Observation', ['date', 'value'])


def getfilename(initialdir=None, title=None):
    """Opens a file dialog to let user select a data file. Defaults to home directory"""
    Tk().withdraw()
    
    filename = askopenfilename(initialdir=initialdir or os.path.expanduser("~"),
                               title=title or 'Select report file')
    
    return filename


def get_patient_key(key, row):
    """Given a row from a CSV DictReader, makes a str key for use in a dict given the key/keys passed to it
       If the row is empty for a given key, it generates a random identifier instead
    """
    if isinstance(key, tuple):
        return ' '.join([row[e] for e in key]).title()
    else:
        if row[key] == '':
            return random_identifier()
        else:
            return row[key]


def build_dict(filename=None, key=None, observation_fields=None, encoding='utf-8-sig', date_format = '%m/%d/%Y'):
    """Given a filename to a CSV file, and a Observation tuple with date and value pointing to the appropriate columns,
       return a dict of patients with values being a list of observations"""
    observations = defaultdict(list)
    with open(filename, encoding=encoding) as f:
        reader = csv.DictReader(f)
        for row in reader:
            patient_key = get_patient_key(key, row)
            observation = Observation(date=dt.strptime(row[observation_fields.date], date_format).date(),
                      value=row[observation_fields.value])
            observations[patient_key].append(observation)
    return observations
    
def filter_observations_by_code(observations, init_code=None):
    """Given a list of observations and an initial billing code,
    Return a sorted list of all observations that occurred from the
    date of the most recent initial billing onward"""
    
    # Sort observations by date and get the most recent initial visit date
    init_observation = max([observation for observation in observations 
                             if observation.value == init_code], 
                             key=attrgetter('date'))
    
    if not init_observation:
        return None, []
    else:
        init_date = init_observation.date
        
    # Make sure patient has had a initial visit billed
    if not init_date:
        return None, []
        
    # Now go through all observations and pull values for any dates occurring on or after the initial visit date
    return (init_date, [observation
                        for observation in sorted(observations, key=attrgetter('date'))
                        if observation.date >= init_date]
            )


def filter_observations_by_date(observations, init_date=None):
    """given a list of observations and a cut off date, return a sorted list of observation after that date"""
    return [observation
            for observation in sorted(observations, key=attrgetter('date'))
            if observation.date >= init_date]


def get_patient_status(start, end):
    if not isinstance(end, numbers.Real) or not isinstance(start, numbers.Real):
        return 'NA'
    elif end < start:
        return 'Improved'
    elif end == start:
        return 'Same'
    elif end > start:
        return 'Worsened'