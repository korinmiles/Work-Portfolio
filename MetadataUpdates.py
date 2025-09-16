# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 15:50:16 2020

@author: kmmiles
"""

import arcpy
import os
import xml.etree.ElementTree as ET
import time
import pandas
import numpy
import csv
import sys


username = arcpy.GetParameterAsText(0)

# contact info to update
contact_person = username
contact_position = arcpy.GetParameterAsText(1)
contact_orginization = arcpy.GetParameterAsText(2)

contact_addrtype = arcpy.GetParameterAsText(3)
contact_address = arcpy.GetParameterAsText(4)
contact_city = arcpy.GetParameterAsText(5)
contact_state = arcpy.GetParameterAsText(6)
contact_postal = arcpy.GetParameterAsText(7)
contact_country = arcpy.GetParameterAsText(8)
contact_phone = arcpy.GetParameterAsText(9)
contact_email = arcpy.GetParameterAsText(10)

abstract_text = arcpy.GetParameterAsText(11)
process_text = arcpy.GetParameterAsText(12)
fullpath = arcpy.GetParameterAsText(13)


# metadata functions
# patterned off these pages:
#http://gis.stackexchange.com/questions/34729/creating-table-containing-all-filenames-and-possibly-metadata-in-file-geodatab
#https://stackoverflow.com/questions/20224501/updating-metadata-for-feature-classes-programatically-using-arcpy
#http://desktop.arcgis.com/en/arcmap/latest/manage-data/metadata/editing-metadata-for-many-arcgis-items.htm#ESRI_SECTION1_35F23B7429E8484890058C3C7A797D6E

def update_metadata(root, tag, update_text, update_type, sub_element):
    """
    update the metadata text for a tag under root
    update type
        replace - replace the text
        update  - append new text to the old
        add     - add a new subelement
            sub_element should be a dict of
            "name" : name
            "attributes" : attributes to add to the tag
            be sure to select the parent tag

    using a simple hack for html elements
    replace the trailing three divs with a new element and the divs

    converting to unicode and back to ascii for special characters found
    """

    num_elements = 0
    upd_elements = root.findall(tag)

    for element in upd_elements:
        if update_type == "update":
            # if default text is present, replace instead of update

            # encode/decode unicode text
            if element.text is not None:
                element.text = element.text.encode('utf-8')

                if element.text.startswith("REQUIRED:"):
                    element.text = update_text
                    num_elements += 1
                elif element.text.endswith("""&lt;/DIV&gt;&lt;/DIV&gt;&lt;/DIV&gt;"""):
                    html_insert = """&lt;P&gt;&lt;SPAN&gt;{0}&lt;/SPAN&gt;&lt;/P&gt;&lt;/DIV&gt;&lt;/DIV&gt;&lt;/DIV&gt;""".format(update_text)
                    element.text = element.text.replace("""&lt;/DIV&gt;&lt;/DIV&gt;&lt;/DIV&gt;""", html_insert)
                elif element.text.endswith("""</DIV></DIV></DIV>"""):
                    html_insert = """<P><SPAN>{0}</SPAN></P></DIV></DIV></DIV>""".format(update_text)
                    element.text = element.text.replace("""</DIV></DIV></DIV>""", html_insert)
                else:
                    if element.text.endswith("."):
                        element.text = "{0} {1}".format(element.text, update_text)
                    else:
                        element.text = "{0}. {1}".format(element.text, update_text)
                    num_elements += 1

                element.text = element.text.decode('utf-8')

        elif update_type == "replace":
            element.text = update_text
            num_elements += 1
        elif update_type == "add":
            # add a subelement
            # sub_element should be a dict of name, attributes
            newSubElem = ET.SubElement(element, sub_element["name"])

            if sub_element["attributes"] is not None:
                for k, v in sub_element["attributes"].iteritems():
                    newSubElem.set(k, v)

            newSubElem.text = update_text
            num_elements += 1

    return num_elements

def update_datasets(dataset_list, update_dict, abstract_dict):
    """
    for a list of datasets, update the metadata for the nodes
    using the input dict of { 'node': {'text': '', 'type': '', 'sub_element': ''}}

    added a simple hack to update html elements

    abstract seems to be missing sometimes, adding a new tag using the abstract_dict
    """

    # get the template xml file
    install_dir = arcpy.GetInstallInfo("desktop")["InstallDir"]
    copy_xslt = r"{0}".format(os.path.join(install_dir,"Metadata\Stylesheets\gpTools\exact copy of.xslt"))

    for dataset in dataset_list:
        # create a temp xml file from the dataset
        xmlfile = arcpy.CreateScratchName(".xml", workspace=arcpy.env.scratchFolder)
        arcpy.XSLTransform_conversion(dataset, copy_xslt, xmlfile, "")

        # use the ElementTree library to parse the file
        tree = ET.ElementTree()
        tree.parse(xmlfile)

        root = tree.getroot()

        # update the metadata using the update function
        for node, values in update_dict.items():
            update_metadata(root, node, values['text'], values['type'], values['sub_element'])

        # some files are missing the abstract tag
        # dataIdInfo is usually present, add a new tag there
        abstract_tag1 = root.find('dataIdInfo/idAbs')
        abstract_tag2 = root.find('idinfo/descript/abstract')
        abstract_tag3 = root.find('idPurp/idAbs')
##        abstract_tag4 = root.find('metadata')

        if abstract_tag1 is None and abstract_tag2 is None and abstract_tag3 is None:
            # missing abstract, add it to dataIdInfo if present
            if root.find('idinfo') is None:
                tag1 = ET.Element('idinfo')
                root.append(tag1)
                tag2 = ET.Element('descript')
                tag1.append(tag2)
                print('abstract tag is missing, adding a new tag')
                for k, v in abstract_dict.items():
                    update_metadata(root, k, v['text'], v['type'], v['sub_element'])

        # save and import the new metadata xml file
        tree.write(xmlfile)
        arcpy.MetadataImporter_conversion(xmlfile, dataset)

        # delete the temp xml file, change remove_temp_xml to false to debug
        # this is global, should remove at some point
        if remove_temp_xml:
            os.remove(xmlfile)

    return

# List feature classes stored in feature datasets in the gdb
def listFcsInGDB(fullpath):
    """
    List feature classes stored in feature datasets in the gdb
    """
    for fds in arcpy.ListDatasets('','feature') + ['']:
        for fc in arcpy.ListFeatureClasses('','',fds):
            yield os.path.join(arcpy.env.workspace, fds, fc)


#------------------------------ metadata updating------------------------------

# variables for updating metadata
mod_date = time.strftime("%Y%m%d", time.localtime())
friendly_date = time.strftime("%Y-%m-%d", time.localtime())


# change to false for debugging the xml files
# global variable in metadata function - should remove at some point
remove_temp_xml = True


# use a dictionary to hold subelement info so we can add attributes
# For adding the process step
subelement_process_dict = {
        "name" : "Process",
        "attributes" : {
                "Date" : mod_date,
                "Time" : time.strftime("%I%M%S", time.localtime()),
                "ToolSource" : "T:\FS\Reference\GeoTool\r06_csa4\Toolbox\NepaSOP.tbx"
                }
        }

# in case the abstract is missing in the data, need to add a subelement
# tag could also be dataIdInfo
abstract_dict = {
        'idinfo/descript':
            {"text" : abstract_text,
             "type" : "add",
             "sub_element" : {
                     "name" : "abstract",
                     "attributes" : None
                     }
            }
        }

# dictionary of tags to update
#   "tag" :
#     "text" to update
#     "type" of update - add, update, or replace
#     "sub_element" - if adding a new tag, a dictionary of name and attributes as above
#                   remember to select the parent tag
metadata_update_dict = {
        "idinfo/descript/abstract" :
            {"text" : abstract_text,
             "type" : "update",
             "sub_element" : ""
             },
        "Esri/DataProperties/lineage":
            {"text" : process_text,
             "type" : "add",
             "sub_element" : subelement_process_dict
             },
        "Esri/ModDate" :
            {"text" : mod_date,
             "type" : "replace",
             "sub_element" : ""
             },
        ".//cntinfo/cntorgp/cntper" :
            {"text" : contact_person,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntperp/cntorg" :
            {"text" : contact_orginization,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntorgp/cntorg" :
            {"text" : contact_orginization,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntpos" :
            {"text" : contact_position,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/addrtype" :
            {"text" : contact_addrtype,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/address" :
            {"text" : contact_address,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/city" :
            {"text" : contact_city,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/state" :
            {"text" : contact_state,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/postal" :
            {"text" : contact_postal,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntaddr/country" :
            {"text" : contact_country,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntvoice" :
            {"text" : contact_phone,
             "type" : "replace",
             "sub_element" : ""
            },
        ".//cntinfo/cntemail":
            {"text" : contact_email,
             "type" : "update",
             "sub_element" : ""
             },
        ".//dataIdInfo/idAbs" :
            {"text" : abstract_text,
             "type" : "update",
             "sub_element" : ""
             },
        ".//idPurp/idAbs" :
            {"text" : abstract_text,
             "type" : "update",
             "sub_element" : ""
             }
        }

# now we need to set the workspace, list datasets, and update metadata
arcpy.env.workspace = fullpath
dataset_list = listFcsInGDB(fullpath)
update_datasets(dataset_list, metadata_update_dict, abstract_dict)

#----------------------------end metadata updating-----------------------------
