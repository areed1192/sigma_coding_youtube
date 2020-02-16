import csv
import pprint
import pathlib
import collections
import xml.etree.ElementTree as ET

# Define Current Working Directory:
sec_directory = pathlib.Path.cwd().joinpath('xbrl-documents')

# Define the file path.
file_htm = sec_directory.joinpath('fb-09302019x10q.htm').resolve()
file_cal = sec_directory.joinpath('fb-20190930_cal.xml').resolve()
file_lab = sec_directory.joinpath('fb-20190930_lab.xml').resolve()
file_def = sec_directory.joinpath('fb-20190930_def.xml').resolve()

# Initalize storage units, one will be the master list, one will store all the values, and one will store all GAAP info.
storage_list = []
storage_values = {}
storage_gaap = {}

# Create a named tuple
FilingTuple = collections.namedtuple('FilingTuple',['file_path','namespace_root','namespace_label'])

# Initalize my list of named tuples, I plan to parse.
files_list = [
    FilingTuple(file_cal, r'{http://www.xbrl.org/2003/linkbase}calculationLink', 'calculation'), 
    FilingTuple(file_def, r'{http://www.xbrl.org/2003/linkbase}definitionLink','definition'), 
    FilingTuple(file_lab, r'{http://www.xbrl.org/2003/linkbase}labelLink','label')
    ]

# Labels come in two forms, those I want and those I don't want.
avoids = ['linkbase','roleRef']
parse = ['label','labelLink','labelArc','loc','definitionLink','definitionArc','calculationArc']

# part of the process is matching up keys, to do that we will store some keys as we parse them.
lab_list = set()
cal_list = set()

# loop through each file.
for file in files_list:

    # Parse the tree by passing through the file.
    tree = ET.parse(file.file_path)

    # Grab all the namespace elements we want.
    elements = tree.findall(file.namespace_root)

    # Loop throught each element that was found.
    for element in elements:

        # if the element has childrent we need to loop through those.
        for child_element in element.iter():

            # split the label to remove the namespace component, this will return a list.
            element_split_label = child_element.tag.split('}')

            # The first element is the namespace, and the second element is a label.
            namespace = element_split_label[0]
            label = element_split_label[1]

            # if it's a label we want then continue.
            if label in parse:

                # define the item type label
                element_type_label = file.namespace_label + '_' + label
                
                # initalize a smaller dictionary that will house all the content from that element.
                dict_storage = {}
                dict_storage['item_type'] = element_type_label

                # grab the attribute keys
                cal_keys = child_element.keys()

                # for each key.
                for key in cal_keys:

                    # parse if needed.
                    if '}' in key:

                        # add the new key to the dictionary and grab the old value.
                        new_key = key.split('}')[1]
                        dict_storage[new_key] = child_element.attrib[key]

                    else:
                        # grab the value.
                        dict_storage[key] = child_element.attrib[key]

                # At this stage I need to create my master list of IDs which is very important to program. I only want unique values.
                # I'm still experimenting with this one but I find `Label` XML file provides the best results.
                if element_type_label == 'label_label':

                    # Grab the Old Label ID for example, `lab_us-gaap_AllocatedShareBasedCompensationExpense_E5D37E400FB5193199CFCB477063C5EB`
                    key_store = dict_storage['label']

                    # Create the Master Key, now it's this: `us-gaap_AllocatedShareBasedCompensationExpense_E5D37E400FB5193199CFCB477063C5EB`
                    master_key = key_store.replace('lab_','')

                    # Split the Key, now it's this: ['us-gaap', 'AllocatedShareBasedCompensationExpense', 'E5D37E400FB5193199CFCB477063C5EB']
                    label_split = master_key.split('_')

                    # Create the GAAP ID, now it's this: 'us-gaap:AllocatedShareBasedCompensationExpense'
                    gaap_id = label_split[0] + ':' + label_split[1]

                    # One Dictionary contains only the values from the XML Files.
                    storage_values[master_key] = {} 
                    storage_values[master_key]['label_id'] = key_store
                    storage_values[master_key]['location_id'] = key_store.replace('lab_','loc_')
                    storage_values[master_key]['us_gaap_id'] = gaap_id
                    storage_values[master_key]['us_gaap_value'] = None
                    storage_values[master_key][element_type_label] = dict_storage 

                    # The other dicitonary will only contain the values related to GAAP Metrics.
                    storage_gaap[gaap_id] = {}
                    storage_gaap[gaap_id]['id'] = gaap_id
                    storage_gaap[gaap_id]['master_id'] = master_key

                # add to dictionary.
                storage_list.append([file.namespace_label, dict_storage])

'''
    PARSE THE HTML FILE.
'''

# Load the HTML file.
tree = ET.parse(file_htm)

# create a new dictionary to store context info.
context_dictionary = {}

# loop through all the elements in the HTML file.
for element in tree.iter():
    
    # for nonNumber the content is different.
    if 'nonNumeric' in element.tag:
        
        # Grab the attribute name and the master ID.
        attr_name = element.attrib['name']
        gaap_id = storage_gaap[attr_name]['master_id']

        storage_gaap[attr_name]['context_ref'] = element.attrib['contextRef']
        storage_gaap[attr_name]['context_id'] = element.attrib['id']
        storage_gaap[attr_name]['continued_at'] = element.attrib.get('continuedAt','null')
        storage_gaap[attr_name]['escape'] = element.attrib.get('escape','null')
        storage_gaap[attr_name]['format'] = element.attrib.get('format','null')

    # same for nonFraction tags.
    if 'nonFraction' in element.tag:
        
        # Grab the attribute name and the master ID.
        attr_name = element.attrib['name']
        gaap_id = storage_gaap[attr_name]['master_id']

        storage_gaap[attr_name]['context_ref'] = element.attrib['contextRef']
        storage_gaap[attr_name]['fraction_id'] = element.attrib['id']
        storage_gaap[attr_name]['unit_ref'] = element.attrib.get('unitRef','null')
        storage_gaap[attr_name]['decimals'] = element.attrib.get('decimals','null')
        storage_gaap[attr_name]['scale'] = element.attrib.get('scale','null')
        storage_gaap[attr_name]['format'] = element.attrib.get('format','null')
        storage_gaap[attr_name]['value'] = element.text.strip() if element.text else 'Null'

        # don't forget to store the actual value if it exist.
        if gaap_id in storage_values:
            storage_values[gaap_id]['us_gaap_value'] = storage_gaap[attr_name]  

    # context is very different.
    if 'context' in element.tag:
        context_dictionary[element.attrib['id']] = {}
        for cnx_item in element.iter():
            for att in cnx_item.attrib:
                if att:
                    context_dictionary[element.attrib['id']][att] = cnx_item.attrib[att]
                    if cnx_item.text.strip() != '':
                        context_dictionary[element.attrib['id']]['text'] = cnx_item.text.strip()
                    else:
                        context_dictionary[element.attrib['id']]['text'] = 'Null'
    

# pprint.pprint(list(storage_values.keys()))
# pprint.pprint(storage_list)
# # for key in storage_gaap:
# #     if storage_gaap[key]['master_id'] in storage_values:
# #         storage_values[storage_gaap[key]['master_id']]['us_gaap_value'] = storage_gaap[key]


# first write the xbrl_content.
file_name = 'data\sec_xbrl_scrape_content.csv'

# open the file.
with open(file_name, mode='w', newline='') as sec_file:

    # create the writer.
    writer = csv.writer(sec_file)

    # write the header.
    writer.writerow(['FILE','LABEL','VALUE'])

    # dump the dict to the csv file.
    for dict_cont in storage_list:
        for item in dict_cont[1].items():
            writer.writerow([dict_cont[0]] + list(item))


# second write the filing_values.
file_name = 'data\sec_xbrl_scrape_values.csv'

# open the file.
with open(file_name, mode='w', newline='') as sec_file:

    # create the writer.
    writer = csv.writer(sec_file)

    # write the header.
    writer.writerow(['ID','CATEGORY','LABEL','VALUE'])

    # start at level 1
    for storage_1 in storage_values:

        # level two is grab the items.
        for storage_2 in storage_values[storage_1].items():

            # if the value is a dictionary, we have one more possible level.
            if isinstance(storage_2[1], dict):

                # level three grab the items.
                for storage_3 in storage_2[1].items():

                    # Write the values to the csv.
                    writer.writerow([storage_1] + [storage_2[0]] + list(storage_3))

            # else just write it to the CSV.
            else:
                if storage_2[1] != None:
                    writer.writerow([storage_1] + list(storage_2) + ['None'])  