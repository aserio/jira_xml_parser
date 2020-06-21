# Copyright (c) 2020 Adrian S. Lemoine
#
#   Distributed under the Boost Software License, Version 1.0. 
#   (See accompanying file LICENSE_1_0.txt or copy at 
#   http://www.boost.org/LICENSE_1_0.txt)

import xml.etree.ElementTree as ET
from openpyxl import Workbook, load_workbook
import os
import argparse
from datetime import date

def read_xml(file):
    # Read in xml file
    if not os.path.exists(file):
        sys.exit("No XML file found.")
    root = ET.parse(file).getroot()
    #print(root.tag)
    return root

def read_excel(file):
    # Set up Excel
    if os.path.exists(excel_file):
        wb = load_workbook(filename = excel_file)
    else:
       wb = Workbook()
       ws = wb.active
       # Add Heading
       ws.append(headings)
    return wb

def find_keys (root):
    # Finds all tickets in the XML file
    # The ticket's name is placed in the "key" XML tag
    key_ring = []
    for el in root.iter('key'):
        key_ring.append(el.text)
    return key_ring

def find_tag (key, tag):
    # Use the key to find a tags in the item that match the key
    # Returns a string
    tag_tablet = ''
    if tag is 'labels':
        tag_tablet = find_labels(key)
    elif tag is 'triage':
        tag_tablet = find_triage(key)
    elif tag is 'priority':
        tag_tablet = find_priority(key)
    else:
        text = "./channel/item/[key='"+key+"']"
        tag_tablet = xml_root.find(text).find(tag).text
    return tag_tablet

def find_labels(key):
    text = "./channel/item/[key='"+key+"']"
    labels = ""
    cntr = 0
    xelement = xml_root.find(text)
    for el in xelement.iter('label'):
       if cntr is 0:
           labels = el.text
       else:
           labels = labels + ', ' + el.text
       cntr = cntr + 1
    return labels

def find_triage(key):
    text = "./channel/item/customfields/..[key='"+key+"']/customfields/customfield[@id='customfield_14308']/customfieldvalues/"
    if ET.iselement(xml_root.find(text)):
        xelement = xml_root.find(text)
        return clean_string(xelement.text)
    else:
        return "None"

def find_priority(key):
    text = "./channel/item/[key='"+key+"']"
    prty = ""
    if ET.iselement(xml_root.find(text).find('priority')):
        prty = xml_root.find(text).find('priority').text
        return prty
    else:
        return "None"

def clean_string(str):
    # Remove returns and excess whitespace.
    str = str.replace('\n',"").replace("  ", "")
    return str

def add_issue(issue):
    ws.append(issue)

# Program Start
## Set up arguments
parser = argparse.ArgumentParser(description = 
    "jira_xml_parser is a program which takes an XML file "
    "of Jira issues and inserts the new issues into a provided Excel "
    "file.")
parser.add_argument('xml_file', help='Path to Jira XML file.')
parser.add_argument('excel_file', nargs = '?', help = 'Path to Excel file.'
                    , default = 'jira_issues_'+date.today().strftime("%Y.%m.%d")+'.xlsx')
parser.add_argument('-f', '--force_update', action = 'store_true', 
                    help = 'Force all issues to be updated regardless of when it was last updated')

args = parser.parse_args()
xml_file = args.xml_file
excel_file = args.excel_file

print("Reading in ", xml_file)
print("Writing to ", excel_file)

# URL Root for Jira tickets
jira_url_root = 'http://ontrack-internal.amd.com/browse/'

# List the tags you are interested in collcecting
dict_keys = ["type", "key" , "summary", "assignee"
                         , "reporter", "status", "created", "updated"
                         , "priority", "triage", "labels"]
                         #, "Target SW Release"]

# List the heading you would like to use
headings = ["Issue Type", "Key", "Summary", "Assignee"
                         , "Reporter", "Status", "Created", "Updated"
                         , "Priority", "Triage Assignment", "Lablels"]
                         #, "Target SW Release"]

xml_root = read_xml(xml_file)
wb = read_excel(excel_file)
ws = wb.active

column_b = ws['B'] # Need to update to find the colummn based on heading
column_h = ws['H']
excel_data = []
excel_keys = []
for el in range(1,len(column_b)):
    temp = []
    temp.append(el)
    temp.append(column_b[el].value)
    excel_keys.append(column_b[el].value)
    temp.append(column_h[el].value)
    excel_data.append(temp)

# Find all unique tickets in XML
keys = find_keys(xml_root)
# print(keys)

# Create a list of tickets and thier tags
tickets = []
for el in keys:
    temp = []
    for el2 in dict_keys:
        temp.append(find_tag(el, el2))
    tickets.append(temp)

# Write to Excel
## Search for copies
new_tickets = 0
updated_tickets = 0

for el in range(len(keys)):
    try:
        # Try to find each key in the excel keys.
        # Will throw an exception if it cannot find a match
       
        ## Return the index of the key
        result = excel_keys.index(keys[el])
        ## Determine if all values should updated vased on "Updated" value
        ### Compare Updated value (Column H)
        if tickets[el][7] != excel_data[result][2] or args.force_update:   #Need to enumerate headings
            updated_tickets += 1
            ## Update ticket with current infromation
            for el2 in range(len(dict_keys)):
                c = el2 + 1 # Columns start with 1
                r = result + 2 #Rows start with 1, add offset for heading
                ws.cell(row=r, column=c).value = tickets[el][el2]
    except ValueError:
        ## The keys was not found append the new ticket
        new_tickets += 1
        add_issue(tickets[el])

print("Updated Tickets: ", updated_tickets)
print("New Tickets", new_tickets)

# Add URLs to keys
for cell in ws.iter_rows(min_row=2, min_col=2, max_col=2):
    if cell[0].style is not "Hyperlink":
        iname = cell[0].value
        ws[cell[0].coordinate].hyperlink = jira_url_root + iname
        ws[cell[0].coordinate].style = "Hyperlink"

wb.save(excel_file)