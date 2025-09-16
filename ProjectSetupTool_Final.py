'''
#-------------------------------------------------------------------------------
# Name:        Project Setup Tool v3.0
# Purpose:
#
# Author:      kmmiles
#
# Created:     27/08/2019
# Copyright:   (c) kmmiles 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------
Last Updated 10/14/2020 by Korin Miles
Updated Dec 2019 Joel Thompson
Create Project folder structure and add all necessary data
    Clipping data to project boundary
    And/Or Select data that intersects boundary and copy to project
    Add general layer files to project
    Update metadata
    Create summary tables

Moved the majority of the data to process to a config file (csv)
    We can set one up for each region and allow the user to add the cofig file 
General layer files are saved in CSA specific folders
    The tool is set up to add all layer files saved in these folders

There are a few general functions with hardcoded data, outside of the config
    Roads
    Vegetation
    NHD Flowlines by subwatershed - subwatershed should be created by this tool
        before it's available for clipping, therefore can't be in the config

Optional extra config files parameter
    People can use these for forest specific data or else create their own
    For adding extra data to the project
    Originally there were values for forest specific data in CSA 4
    But now these can be added using the extra config file selectors



The config file is a csv with the following fields
    layer_source     - full path to surce data
    feature_dataset  - feature dataset to store output
    clip_output      - clip name or blank
    select_output    - select name or blank
    clip_symbology   - layer file for updating clip symbology or blank
    select_symbology - layer file for updating select symbology or blank


The select_output field is used when the data should be selected by features
    that intersect the project area boundary instead of clipping
    for when you need features in and out of the boundary

When setting up the config files, need to make sure they have unique names

Updated the config to contain the summary table info
    and a function to read it and create summary tables when info is present

'''

import arcpy
import os
import xml.etree.ElementTree as ET
import time
import pandas
import numpy
import csv
import sys
import time

#Set user set environments from the tool
#Need project folder location, name of project gdb, project area boundary
folder_loc = arcpy.GetParameterAsText(0)  #root project folder location
projectName = arcpy.GetParameterAsText(1)
project_area_bound = arcpy.GetParameterAsText(2)   #project area boundary feature class
project_folder = folder_loc

# these are the keys of the config_dict dictionary
region = arcpy.GetParameterAsText(3)

sr= arcpy.GetParameterAsText(4)
spatial_ref = sr

### forest specific config files - string with select multi

### extra config file, let the user add a custom one if needed
config_file_extra1 = arcpy.GetParameterAsText(5)
config_file_extra2 = arcpy.GetParameterAsText(6)

# whether to create veg and roads - boolean
write_veg = arcpy.GetParameterAsText(6)
write_roads = arcpy.GetParameterAsText(7)


config_dict = {
        "Region01" : {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region01_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region02" : {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region02_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region03" : {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region03_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region04" : {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region04_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region05" : {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region05_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region06": {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region06_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region08": {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region08_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region09": {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region09_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        "Region10": {"config" : r"{'ExamplePath'}\NEPA_SOP\config\Region10_edw_nepa_sop_config.csv",
                  "general" : r"{'ExamplePath'}\NEPA_SOP\general_layer_files"},
        }


# project geodatabase
gdb_name = projectName + ".gdb"
# project path
fullpath = os.path.join(folder_loc, "GIS", "Data", gdb_name)

# config csv file and layer file directory from the dictionary
config_file = config_dict[region]["config"]
general_lyrs = config_dict[region]["general"]


# Set workspace
arcpy.env.workspace = fullpath
arcpy.env.overwriteOutput= True

#Set Output Coordinate System
arcpy.env.outputCoordinateSystem = sr

#----------------- global variables not in the config -------------------------

# project area boundary symbology
pabSym = r"{'ExamplePath'}\NEPA_SOP\Project Boundary.lyr"
pab_symbology = pabSym

# acres/miles calculations
acExpress = """!Shape.area@acres!"""
miExpress = """!Shape.length@miles!"""


if region == "Region01" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R01.Hydrography\S_R01.NHDFlowline";
elif region == "Region02" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R02.Hydrography\S_R02.NHDFlowline";   
elif region == "Region03" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R03.Hydrography\S_R03.NHDFlowline";
elif region == "Region04" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R04.Hydrography\S_R04.NHDFlowline";
elif region == "Region05" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R05.Hydrography\S_R05.NHDFlowline";
elif region == "Region06" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R06.HydrographyPrj\S_R06.NHDFlowlinePrj";
elif region == "Region08" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_NHD.Hydro_08_NHD\S_NHD.Hydro_08_NHD_Flowline";
elif region == "Region09" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R09.Hydrography\S_R09.NHDFlowline";
elif region == "Region10" : nhdFlow = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_R10.Hydrography\S_R10.NHDFlowline"




# NHD flow lines for clipping to subwatershed after it's created
##nhdFlow = r"T:\FS\Reference\GeoTool\r06\DatabaseConnection\edw_sde_default_as_myself.sde\S_R06.HydrographyPrj\S_R06.NHDFlowlinePrj"
perennSym = r"{'ExamplePath'}\NEPA_SOP\Streams.lyr"

# roads data
rdConv = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.Transportation\S_USA.RoadCore_Converted"
rdDecom = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.Transportation\S_USA.RoadCore_Decommissioned"
rdExist = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.Transportation\S_USA.RoadCore_Existing"
rdPlan = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.Transportation\S_USA.RoadCore_Planned"
roads = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.Transportation\S_USA.Road"

# road symbology
roadSym = r"{'ExamplePath'}\NEPA_SOP\Roads.lyr"

# veg data
veg = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.NRIS_Veg\S_USA.NRIS_VegPoly"
vegTbl = r"{'ExamplePath'}\edw_sde_default_as_myself.sde\S_USA.NRIS_VegCharacterizations"


#-----------------------------Functions----------------------------------------

def pabClip(fc, dataset, outFC, project_area_bound, fullpath):
    """
    Clip features to PAB and save to feature dataset
    """

    arcpy.Clip_analysis(fc, project_area_bound, os.path.join(fullpath, dataset, outFC))



def pabSelect(fc, feature_dataset, outFC, project_area_bound, fullpath):
    """
    Select features that intersect the PAB and save to new file in feature dataset
    reworked to remove globals and use fc basename for temp layer file
    """
    temp_lyr = "{}_temp".format(os.path.basename(fc).split('.')[-1])
    output_file = os.path.join(fullpath, feature_dataset, outFC)

    arcpy.MakeFeatureLayer_management(project_area_bound, "PAB_Lyr")
    arcpy.MakeFeatureLayer_management(fc, temp_lyr)

    # input_layer, overlap_type, select_features
    intersect = arcpy.SelectLayerByLocation_management(temp_lyr, 'INTERSECT', "PAB_Lyr")
    arcpy.CopyFeatures_management(intersect, output_file)

    return output_file


def applyLayer(outDS, outFC, sym, fullpath, df, folder_loc):
    """
    apply symbology from sample and save output layer file
    mostly replaced with apply_symbology but
    this one is still used by the infra roads function
    """
    addLyr = arcpy.mapping.Layer(os.path.join(fullpath, outDS, outFC))

    arcpy.ApplySymbologyFromLayer_management(addLyr, sym)

    arcpy.mapping.AddLayer(df, addLyr, "AUTO_ARRANGE")
    arcpy.SaveToLayerFile_management(addLyr, os.path.join(folder_loc, "GIS", "Layerfile", outFC + ".lyr"))


def apply_symbology(feature_dataset, feature_class, symbology, gdb_path, mxd_df, project_folder, mxd): 
    """
    apply symbology from sample layer file and save output layer file
    added try/except to handle cases where ApplySymbology fails (NHDFlowline)
        UpdateLayer will be used if ApplySymbology fails
    """
    add_lyr = arcpy.mapping.Layer(os.path.join(gdb_path, feature_dataset, feature_class))

    try:
        arcpy.ApplySymbologyFromLayer_management(add_lyr, symbology)
        arcpy.mapping.AddLayer(mxd_df, add_lyr, "AUTO_ARRANGE")
        arcpy.SaveToLayerFile_management(add_lyr,
                                         os.path.join(project_folder, "GIS",
                                                      "Layerfile", feature_class + ".lyr"))

    except:
        try:
            arcpy.mapping.AddLayer(mxd_df, add_lyr, "AUTO_ARRANGE")
            map_lyr = arcpy.mapping.ListLayers(mxd, feature_class, mxd_df)[0]
            map_sym = arcpy.mapping.Layer(symbology)

            arcpy.mapping.UpdateLayer(mxd_df, map_lyr, map_sym, True)
            arcpy.SaveToLayerFile_management(add_lyr,
                                             os.path.join(project_folder, "GIS",
                                                          "Layerfile", feature_class + ".lyr"))

        except:
            arcpy.AddWarning("Error adding symbology")
            arcpy.AddWarning("error: {}".format(sys.exc_info()[0]))

    return


def updateLayer(outDS, outFC, sym, fullpath, df, folder_loc, mxd):
    """
    This function should no longer be necessary - use apply_symbology
    update symbology from sample layer file and save to new layer file
    globals - fullpath, df, mxd, folder_loc
    some layers fail on applySymbology, new function should fix
    """
    lyr = arcpy.mapping.Layer(os.path.join(fullpath, outDS, outFC))

    arcpy.mapping.AddLayer(df, lyr, "AUTO_ARRANGE")

    mapLyr = arcpy.mapping.ListLayers(mxd, outFC, df)[0]
    mapSym = arcpy.mapping.Layer(sym)

    arcpy.mapping.UpdateLayer(df, mapLyr, mapSym, True)
    arcpy.SaveToLayerFile_management(lyr, os.path.join(folder_loc, "GIS", "Layerfile", outFC + ".lyr"))




def get_new_extent(extent, current_spatial_ref, new_spatial_ref):
    """
    convert extent from current spatial reference
    to new spatial reference using extent 
    returns space separated string
      MinX MinY MaxX MaxY
    """
    extent_split = extent.split(' ')
    
    new_min_point = arcpy.PointGeometry(
            arcpy.Point(extent_split[0], extent_split[1]), 
                       current_spatial_ref).projectAs(new_spatial_ref)
    
    new_max_point = arcpy.PointGeometry(
            arcpy.Point(extent_split[2], extent_split[3]), 
                       current_spatial_ref).projectAs(new_spatial_ref)
    
    new_extent = "{0} {1} {2} {3}".format(new_min_point.firstPoint.X, 
                  new_min_point.firstPoint.Y, 
                  new_max_point.firstPoint.X, 
                  new_max_point.firstPoint.Y)
    
    return new_extent

def infraRds(con, decom, ex, plan, fullpath, pabUnion, roads, roadSym, mxd, folder_path):
    """
    create infra roads layers
    """
    road_merge = os.path.join(fullpath, "Transportation", "AllRoads")
    road_event_table = os.path.join(fullpath, "RoadLineEventsTable")
    road_locate = os.path.join(fullpath, "LocateFeaturesTable")
    road_overlay = os.path.join(fullpath, "RoadEvents_Overlay")
    
    road_infra_path = os.path.join(fullpath, "Transportation")
    road_infra_name = "INFRA_Roads"
    road_infra = os.path.join(road_infra_path, road_infra_name)
    
    road_two_four = os.path.join(fullpath, "Transportation", "Two_FourDigitRoads")
    
    # should set the extent variable to limit roads being merged
    # but need to check spatial references first, and manage extent later
    original_extent = arcpy.env.extent
    
    pab_union_sr = arcpy.Describe(pabUnion).spatialReference
    roads_sr = arcpy.Describe(con).spatialReference
    
    arcpy.env.outputCoordinateSystem = pab_union_sr
    
    pab_describe = arcpy.Describe(pabUnion)
    pab_extent = "{0} {1} {2} {3}".format(pab_describe.extent.XMin,
                      pab_describe.extent.YMin,
                      pab_describe.extent.XMax,
                      pab_describe.extent.YMax)
    
    if pab_union_sr.factoryCode != roads_sr.factoryCode:
        new_extent = get_new_extent(pab_extent, pab_union_sr, roads_sr)
        arcpy.env.extent = new_extent
    else:
        arcpy.env.extent = pab_extent    
    
    rdFC = arcpy.Merge_management([con, decom, ex, plan], road_merge)
    
    # reset extent to original, should only need it for merge
    arcpy.env.extent = original_extent    
    
    rdLnEventTable = arcpy.TableSelect_analysis(rdFC, road_event_table)
    rdFeat = arcpy.LocateFeaturesAlongRoutes_lr(pabUnion, roads, "RTE_CN", "0 Meters", 
                                                road_locate, "RTE_CN LINE FMEAS TMEAS",
                                                "FIRST", "DISTANCE", "NO_ZERO",
                                                "FIELDS", "M_DIRECTON")
    rdOverlay = arcpy.OverlayRouteEvents_lr(rdLnEventTable, "RTE_CN LINE BMP EMP",
                                            rdFeat, "RTE_CN LINE FMEAS TMEAS",
                                            "INTERSECT", road_overlay,
                                            "RTE_CN LINE BMP EMP", "NO_ZERO",
                                            "FIELDS", "INDEX")    
    
    exspre1 = "[EMP]- [BMP]"
    rdCalc = arcpy.CalculateField_management(rdOverlay, "SEG_LENGTH", exspre1, "VB")
    rdMakeRoute = arcpy.MakeRouteEventLayer_lr(roads, "RTE_CN", rdCalc, "RTE_CN LINE BMP EMP", "ProjectRoadEvent_InfraRoads")
    roadOut = arcpy.FeatureClassToFeatureClass_conversion(rdMakeRoute, road_infra_path, road_infra_name)
    
    create_layer_file(road_infra,
                      roadSym, 
                      os.path.join(folder_path, "INFRA_Roads.lyr"), 
                      mxd)
    
    expres2 = """ID LIKE '%%00000' OR ID LIKE '%%00000S' OR ID LIKE '%%%%000'"""
    rdLyr = arcpy.mapping.Layer(road_infra)
    arcpy.Select_analysis(rdLyr, road_two_four, expres2)    
   
    arcpy.Delete_management(road_merge)
    arcpy.Delete_management(road_event_table)
    arcpy.Delete_management(road_locate)
    arcpy.Delete_management(road_overlay)
    
    return    


def vegClip(fc, vegTable, project_area_bound, fullpath, df):
    """
    Clip vegetation to the project area boundary and join to the veg table
    """
    vegPAB = arcpy.Clip_analysis(fc, project_area_bound, os.path.join(fullpath, "Vegetation", "Veg"))
    veg_table= arcpy.TableSelect_analysis(vegTable, os.path.join(fullpath, "Veg_Table"))
    arcpy.JoinField_management(vegPAB, "SPATIAL_ID", vegTable, "SPATIAL_ID")
    addLyr = arcpy.mapping.Layer(os.path.join(fullpath, "Vegetation", "Veg"))
    arcpy.mapping.AddLayer(df, addLyr, "AUTO_ARRANGE")
##    arcpy.Delete_management(os.path.join(fullpath, "Veg_Table"))
    
def copy_pab(pab, project_name, gdb_fullpath, feature_dataset, folder_loc, pab_symbology, spatial_ref):
    """
    copy the project area boundary to the geodatabase
    apply symbology, add to current mxd dataframe, and create a layer file
    """
    pab_name = "{}_PAB".format(project_name)
    pab_output = os.path.join(gdb_fullpath, feature_dataset, pab_name)
    pab_lyr_output = os.path.join(folder_loc, "GIS", "Layerfile", "PAB.lyr")
    
    pab_sr = arcpy.Describe(pab).spatialReference
    if pab_sr == spatial_ref:
        arcpy.CopyFeatures_management(pab, pab_output)
    else:
        arcpy.Project_management(pab, pab_output, spatial_ref)
        
    # create the layer file
    pab_lyr = arcpy.mapping.Layer(pab_output)
    
    arcpy.ApplySymbologyFromLayer_management(pab_lyr, pab_symbology)
    arcpy.mapping.AddLayer(df, pab_lyr, "AUTO_ARRANGE")
    arcpy.SaveToLayerFile_management(pab_lyr, pab_lyr_output)
    
    return pab_output


def create_pab_union(gdb_path, pab, buffer_dist):
    """
    buffer the project area boundary, union, and calculate IN_AND_OUT
    """
    pab_buffer = os.path.join(gdb_path, "BufferPAB")
    pab_union = os.path.join(gdb_path, "RoadClipBuffer")
    
    arcpy.AddMessage("Buffering the project area boundary")
    arcpy.Buffer_analysis(pab, pab_buffer, buffer_dist, "OUTSIDE_ONLY", "ROUND", "ALL")
    
    arcpy.AddMessage("Creating union of the project area boundary and buffer")
    arcpy.Union_analysis([pab, pab_buffer], pab_union, "ALL", "", "NO_GAPS")
    
    # add and calculate IN_AND_OUT field
    arcpy.AddField_management(pab_union, "IN_AND_OUT", "TEXT")
    expression = "calc( !FID_BufferPAB!)"
    codeBlock = """def calc(val):
        if val == -1:
            return 'IN'
        if val >= 1:
            return 'OUT'"""
    arcpy.CalculateField_management(pab_union, "IN_AND_OUT", expression, "PYTHON", codeBlock)
    
    
    return pab_union   


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


def process_data(config_file, project_area_bound, fullpath, df, folder_loc, mxd):
    """
    read the config file and write out the project data
    clip data
    select by location
    apply symbology & save layer file
    runs the functions
        pabClip
        pabSelect
        update_symbology

    fields required in the config file
        layer_source     - full path to surce data
        feature_dataset  - feature dataset to store output
        clip_output      - clip name or blank
        select_output    - select name or blank
        clip_symbology   - layer file for updating clip symbology or blank
        select_symbology - layer file for updating select symbology or blank

    """
    with open(config_file) as csvfile:
        reader = csv.DictReader(csvfile)

        for row in reader:
            try:
                src_data = row['layer_source']
                feature_dataset = row['feature_dataset']

                arcpy.AddMessage("Processing {}".format(os.path.basename(src_data)))

                # create output files if names are present in the csv
                if row['clip_output']:
                    pabClip(src_data, feature_dataset, row['clip_output'],
                            project_area_bound, fullpath)
                    arcpy.AddMessage("Finished Clip ")

                if row['select_output']:
                    pabSelect(src_data, feature_dataset, row['select_output'],
                              project_area_bound, fullpath)
                    arcpy.AddMessage("Finished Select")

                # add layer files if source is present
                if row['clip_symbology']:
                    apply_symbology(feature_dataset, row['clip_output'],
                                    row['clip_symbology'], fullpath,
                                    df, folder_loc, mxd)
                    arcpy.AddMessage("Apply Clipped Symbology")

                if row['select_symbology']:
                    apply_symbology(feature_dataset, row['select_output'],
                                    row['select_symbology'], fullpath,
                                    df, folder_loc, mxd)
                    arcpy.AddMessage("Apply Select Symbology")

#                if row['add_map']: 
#                    lyr = arcpy.mapping.Layer(os.path.join(fullpath, feature_dataset, row['clip_output']))
#                    arcpy.mapping.AddLayer(df, lyr, "AUTO_ARRANGE")
#                    arcpy.AddMessage("Added to Map")


            except:
                arcpy.AddWarning("--- Unable to process dataset")
                arcpy.AddWarning("error: {}".format(sys.exc_info()[0]))

    return

def add_general_lyr(lyr_file_path, project_folder):
    """
    Add any general layer files from the specified directory
    might think about using either a folder path or a list of files

    lyr_file_path - the directory where the layer files are stored
    project_folder - the project folder to copy the layer files to
    """
    lyr_files = os.listdir(lyr_file_path)
    output = []

    for lyr in lyr_files:
        if lyr.endswith(".lyr"):
            out_lyr = os.path.join(project_folder, "GIS", "Layerfile", os.path.basename(lyr))
            arcpy.SaveToLayerFile_management(os.path.join(lyr_file_path, lyr), out_lyr)
            output.append(lyr)

    return output


def create_layer_file(feature_class, symbology, output_lyr_file, mxd):
    """    
    feature_class - full path to the feature class
    symbology - full path to layer file to use for symbology
    output_lyr_file - full path of output layer file    
    """
    add_lyr = arcpy.mapping.Layer(feature_class)
    mxd_df = arcpy.mapping.ListDataFrames(mxd)[0]
    
    
    try:
        arcpy.ApplySymbologyFromLayer_management(add_lyr, symbology)
        arcpy.mapping.AddLayer(mxd_df, add_lyr, "AUTO_ARRANGE")
        arcpy.SaveToLayerFile_management(add_lyr, output_lyr_file)
        
    except:
        
        try:          
            arcpy.mapping.AddLayer(mxd_df, add_lyr, "AUTO_ARRANGE")
            map_lyr = arcpy.mapping.ListLayers(mxd, add_lyr, mxd_df)[0]
            map_sym = arcpy.mapping.Layer(symbology)
            
            arcpy.mapping.UpdateLayer(mxd_df, map_lyr, map_sym, True)
            arcpy.SaveToLayerFile_management(add_lyr, output_lyr_file)
            
        except:
            arcpy.AddWarning("Error creating layer file")
            arcpy.AddWarning("error: {}".format(arcpy.GetMessages()))
            return None
        
    return output_lyr_file

def veg_clip(veg_poly, veg_table, project_area_bound, df):
    """
    Clip vegetation to the project area boundary and join to the veg table
    """    
    veg_PAB = os.path.join(fullpath, "Vegetation", "Veg")
    
    arcpy.Clip_analysis(veg_poly, project_area_bound, veg_PAB)
    arcpy.JoinField_management(veg_PAB, "SPATIAL_ID", veg_table, "SPATIAL_ID")
    
    addLyr = arcpy.mapping.Layer(veg_PAB)
    arcpy.mapping.AddLayer(df, addLyr, "AUTO_ARRANGE")
    
    return veg_PAB

def getcsvdata():
    with open(config_file_extra1) as csvfile:
        reader = csv.DictReader(csvfile)
        fdlist = []
        for row in reader:
            if row['feature_dataset'] not in fdlist:
                fdlist.append(row['feature_dataset'])
    if config_file_extra2:
        with open(config_file_extra2) as csvfile:
            reader = csv.DictReader(csvfile)
            fdlist = []
            for row in reader: 
                if row['feature_dataset'] not in fdlist:
                    fdlist.append(row['feature_dataset'])        
    return fdlist

        
    

#-------------------------------End functions----------------------------------
# now process the data
    
#Creating GDB and making datasets based off what is used in the csv files. 
    
arcpy.AddMessage("Creating project folder structure")

if not os.path.exists(os.path.join(folder_loc, "GIS")):
    os.makedirs(os.path.join(folder_loc, "GIS"))

if not os.path.exists(os.path.join(folder_loc, "GIS", "Data")):
    os.makedirs(os.path.join(folder_loc, "GIS", "Data"))

if not os.path.exists(os.path.join(folder_loc, "GIS", "Layerfile")):
    os.makedirs(os.path.join(folder_loc, "GIS", "Layerfile"))

if not os.path.exists(os.path.join(folder_loc, "GIS", "MXD")):
    os.makedirs(os.path.join(folder_loc, "GIS", "MXD"))

if not os.path.exists(os.path.join(folder_loc, "GIS", "Implementation")):
    os.makedirs(os.path.join(folder_loc, "GIS", "Implementation"))

if not os.path.exists(os.path.join(folder_loc, "GIS", "AnalysisProduct")):
    os.makedirs(os.path.join(folder_loc, "GIS", "AnalysisProduct"))
    
if not os.path.exists(os.path.join(folder_loc, "GIS", "AnalysisProduct", "Archive")):
    os.makedirs(os.path.join(folder_loc, "GIS", "AnalysisProduct", "Archive"))    


#makeGDB
   

arcpy.CreateFileGDB_management(os.path.join(folder_loc, "GIS", "Data"), gdb_name)
fdlist = getcsvdata()
for fd in fdlist:
    arcpy.CreateFeatureDataset_management(fullpath, fd, sr)
arcpy.CreateFeatureDataset_management(fullpath,"Project", sr)
arcpy.CreateFeatureDataset_management(fullpath,"Transportation", sr)

#end script if no PAB entered
if project_area_bound == "":
    exit()

#Create new mxd
mxd = arcpy.mapping.MapDocument("CURRENT")
df = arcpy.mapping.ListDataFrames(mxd)[0]



#------------- Copy PAB to the project gdb and buffer PAB-----------------
# we could set  arcpy.env.outputCoordinateSystem = sr
# and copy features will use this and reproject
arcpy.AddMessage("Copying project area boundary")

pabName = projectName + "_PAB"
pabSR = arcpy.Describe(project_area_bound).spatialReference
if pabSR == sr:
    pab_project = arcpy.CopyFeatures_management(project_area_bound, os.path.join(fullpath, "Project", pabName))
else:
    pab_project = arcpy.Project_management(project_area_bound, os.path.join(fullpath, "Project", pabName), sr)

pabLyr = arcpy.mapping.Layer(os.path.join(fullpath, "Project", pabName))

arcpy.ApplySymbologyFromLayer_management(pabLyr, pabSym)
arcpy.mapping.AddLayer(df, pabLyr, "AUTO_ARRANGE")
arcpy.SaveToLayerFile_management(pabLyr, os.path.join(folder_loc, "GIS", "Layerfile", "PAB.lyr"))

pab_union = create_pab_union(os.path.join(fullpath, "Transportation"),
                             project_area_bound, "1 Miles")

#--------------------- process the config files -------------------

# process the CSA config file
arcpy.AddMessage("Reading data from config file")
process_data(config_file_extra1, project_area_bound, fullpath, df, folder_loc, mxd)


# process other config files if selected
if config_file_extra2:
    process_data(config_file_extra2, project_area_bound, fullpath, df, folder_loc, mxd)


# extra steps not in the config
    
#------------------ create pab_buffer and copy pab -------------------
arcpy.AddMessage("Copying project area boundary to project geodatabase")
pab_output = copy_pab(project_area_bound, projectName, fullpath, "Project",
                      folder_loc, pab_symbology, spatial_ref)

pab_union = create_pab_union(os.path.join(fullpath, "Transportation"),
                             project_area_bound, "1 Miles")


#---------- clip nhdFlow to the subwatershed if it was created --------------
arcpy.AddMessage("Clipping the NHD Flowlines to the Subwatershed")

swsPAB = os.path.join(fullpath, "Hydrology", "Subwatershed")
if arcpy.Exists(swsPAB):
    nhdFlow_SWS = arcpy.Clip_analysis(nhdFlow, swsPAB, os.path.join(fullpath, "Hydrology", "NHDFlow_SWS"))
    apply_symbology("Hydrology", "NHDFlow_SWS", perennSym, fullpath, df, folder_loc, mxd)


#------------------ roads and veg functions --------------------
if write_roads in ("True", "true"):
    arcpy.AddMessage("Processing INFRA Roads")
    infraRds(rdConv, rdDecom, rdExist, rdPlan, 
             fullpath, pab_union, roads, roadSym, mxd,
             os.path.join(folder_loc, "GIS", "Layerfile"))

if write_veg in ("True", "true"):
    arcpy.AddMessage("Clipping Vegetation")
    veg_clip(veg, vegTbl, pab_project, df)

#------------ general layer files --------------------

# layer files added at the end, not applied to other datasets

# add CSA specific files to separate directories, then add all lyr
# files from that directory, path comes from the CSA selector dictionary
arcpy.AddMessage("Adding general layer files")

try:
    add_general_lyr(general_lyrs, folder_loc)
except:
    arcpy.AddWarning("unable to add general layer files")
    arcpy.AddWarning("error: {}".format(sys.exc_info()[0]))


#--------------- add acres/miles fields and calculate ---------------
arcpy.AddMessage("Calculating Acres and Miles")

fcs = listFcsInGDB(fullpath)

for fc in fcs:
    desc = arcpy.Describe(fc)
    if desc.shapeType == "Polygon":
        arcpy.AddField_management(fc, "Project_Acres", "FLOAT")
        arcpy.CalculateField_management(fc, "Project_Acres", acExpress, "PYTHON")
    elif desc.shapeType == "Polyline":
        arcpy.AddField_management(fc, "Project_Miles", "FLOAT")
        arcpy.CalculateField_management(fc, "Project_Miles", miExpress, "PYTHON")

    # delete empty datasets
    if int(arcpy.GetCount_management(fc).getOutput(0)) == 0:
        arcpy.Delete_management(fc)
        arcpy.AddMessage('{} has been deleted because it contains no features.'.format(fc))

#---------------- remove broken data sources ------------------------
list_src = arcpy.mapping.ListBrokenDataSources(mxd)
for lyr in list_src:
    arcpy.mapping.RemoveLayer(df, lyr)
arcpy.AddMessage("Geometry calculated and features with no data have been deleted. Stand by while the metadata is updated.")


#------------------------------ metadata updating------------------------------

# variables for updating metadata
mod_date = time.strftime("%Y%m%d", time.localtime())
friendly_date = time.strftime("%Y-%m-%d", time.localtime())
username = os.environ.get("USERNAME")
abstract_text = "This data was clipped to an analysis area for the {0} Project by {1} on {2} using the NEPA Project Setup Tool.".format(projectName, username, friendly_date)
process_text = "This data was clipped to an analysis area using the NEPA Project Setup Tool."

# change to false for debugging the xml files
# global variable in metadata function - should remove at some point
remove_temp_xml = True

# contact info to update
contact_person = username
contact_position = ""
contact_orginization = "Enterprise Tech-GIS Group"

contact_addrtype = ""
contact_address = ""
contact_city = ""
contact_state = ""
contact_postal = ""
contact_country = ""
contact_phone = ""
contact_email = ""

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

#---------------- Create basemap for the project -------------------
baseMap = mxd.saveACopy(os.path.join(folder_loc, "GIS", "MXD", "BaseMap.mxd"))
arcpy.AddMessage("Base map saved successfully")






















