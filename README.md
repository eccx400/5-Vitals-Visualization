# 5 Vitals Visualization
> Program used to report patient data using pandas in Python 

## Table of contents
* [General info](#general-info)
* [Technologies](#technologies)
* [Status](#status)

## General info
This program uses pandas and xlsxwriter in order to parse excel files in .xlsx
and csv and renders report of patient 5 vitals, GCS, and prescription information.
The files can be read individually or in batches based on subject data; the Sorter.py
file runs the patient data individually, while the iSorter.py file runs the patient
data in batches based on the file name provided. 

There are a total of 46520 patient ChartEvents files provided along with matching
prescription files that give information on the route and type of prescription
the subject is taking. All subjects have a registered subject_id. 

The project requires us to separate the patients batches by the 5 vitals: Heart Rate,
Blood Pressure, Temperature, Respiratory Rate, and O2 Saturation. The item ids that
are used to parse these data are listed below.

HEART RATE
- 220045, 211

BLOOD PRESSURE
- 220179, 455 -- "Systolic Blood Pressure"
- 220180, 8441 -- "Diastolic Blood Pressure"

TEMPERATURE
- 223762, 676 -- "Temperature Celsius"
- 223761, 678 -- "Temperature Fahrenheit"

RESPIRATORY RATE
- 220210, 618

O2 Saturation
- 220277, 646

The data are then sorted to form a dataframe with each of the vital's itemid and their 
corresponding chart time, and then merged into the visualization table which will be 
printed to the 'Visualization' sheet on the report. The data is then taken from the
Visualization sheet and charted under a new 'Report' sheet. In the report sheet, the
prescription and GCS table values are also listed.

The completed batch files will be stored in folders of their ICD9 disease values, which
are listed below. The subjects which match these values can be found in the folders
above.

ICD9 disease values:
- 460-466  Acute Respiratory Infections
- 470-478  Other Diseases Of Upper Respiratory Tract
- 480-488  Pneumonia And Influenza
- 490-496  Chronic Obstructive Pulmonary Disease And Allied Conditions
- 500-508  Pneumoconioses And Other Lung Diseases Due To External Agents
- 510-519  Other Diseases Of Respiratory System
- E8859    Accidental fall on same level from slipping tripping or stumbling

## Technologies
* Python 2.7 or 3.6

## Status
Project is: _in progress_

Currently working on testing and making sure that all batch reports have correct data and
that the charts and prescription tables are printed correctly. Formatting issues are also 
being resolved. The further implementation of the ADDS chart and categorization programs
is still continuing. Possible implementation of a dashboard or analysis algorithm is in 
consideration.
