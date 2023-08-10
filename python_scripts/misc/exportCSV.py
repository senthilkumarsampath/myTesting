from io import StringIO
from turtle import pd
# define csv dir path
csv_dir = "/Users/surenderkumar/Documents/DST_Hub/Local-Testing/EP_Training_Data/docx_csv_01/"
# iterate over each file to convert as CSV
for file in os.listdir(data_directory):
    # print(data_directory+file)
    # get the file name
    get_filename = os.path.splitext(file)[0].split('/')[-1]
    # start parsing xml into json
    parse_json= parse_xml_to_json(data_directory+file)
    # read json as csv file
    data= pd.read_json(StringIO(parse_json))
    # add filename column into data
    data["file_name"] =  os.path.splitext(data_directory+file)[0].split('/')[-1]
    # saving data as csv file 
    data.to_csv(csv_dir+get_filename+'.csv')

import glob
# path for csv file
csv_dir = "/Users/surenderkumar/Documents/DST_Hub/Local-Testing/EP_Training_Data/docx_csv_01/"
combined_csv_path="/Users/surenderkumar/Documents/DST_Hub/Local-Testing/EP_Training_Data/ep_combined_csv/"
# iterate over csv files
for file in glob.glob(csv_dir+'*.csv'):
    # concat all csv file into 
    data = pd.concat(pd.read_csv, file)
    
    # save as combined csv
    data.to_csv(combined_csv_path+"blo_combined_data.csv")

""" Module to parse an XML file to various data formats """
import json
from lxml import etree
def parse_xml_to_json(xml_file):
    """Function to parse the incoming XMLs.
    Args:
        xml_files (xml): Incoming XML files from DST blob storage.
    Returns:
        json: parse into json format.
    """
    # open xml_file as input files
    with open(xml_file, 'rb') as input_file:
        # read file and store into string
        xml = input_file.read()
    # initialise the etree using loaded in string
    root = etree.fromstring(xml)
    # empty list to store xml elements
    elements = []
    # iterate over the context and extracting the tag elements
    for context in root.getchildren():
        # empty dict
        elem_contents = {}
        # traverse over each elemnt's and extract context's children
        for element in context.getchildren():
            # if the text are None
            if not element.text:
                text='None'
            else:
                text = element.text
            # append the dict
            elem_contents[element.tag] = text
        # append the list
        elements.append(elem_contents)
        # convert into json
        js_data = json.dumps(elements)
    return js_data