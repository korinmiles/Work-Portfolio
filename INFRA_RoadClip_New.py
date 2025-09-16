

Copying project area boundary to project geodatabase
Traceback (most recent call last):
  File "<string>", line 361, in execute
  File "<string>", line 2056, in ProcessTheData
  File "<string>", line 1627, in copy_pab
AttributeError: module 'arcpy.mp' has no attribute 'Layer'


# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 14:51:16 2020

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



fullpath = arcpy.GetParameterAsText(0)  #root project folder location
project_area_bound = arcpy.GetParameterAsText(1)   #project area boundary feature class


# roads data
rdConv = r"{'ExamplePath'}\S_USA.Transportation\S_USA.RoadCore_Converted"
rdDecom = r"{'ExamplePath'}\S_USA.Transportation\S_USA.RoadCore_Decommissioned"
rdExist = r"{'ExamplePath'}\S_USA.Transportation\S_USA.RoadCore_Existing"
rdPlan = r"{'ExamplePath'}\S_USA.Transportation\S_USA.RoadCore_Planned"
roads = r"{'ExamplePath'}\S_USA.Transportation\S_USA.Road"

# road symbology
roadSym = r"{'ExamplePath'}Roads.lyr"


### Functions ###

def infraRds(con, decom, ex, plan, fullpath, pabUnion, roads, roadSym, mxd, folder_path):
    """
    create infra roads layers
    """
    road_merge = os.path.join(fullpath, "AllRoads")
    road_event_table = os.path.join(fullpath, "RoadLineEventsTable")
    road_locate = os.path.join(fullpath, "LocateFeaturesTable")
    road_overlay = os.path.join(fullpath, "RoadEvents_Overlay")
    
    road_infra_path = os.path.join(fullpath)
    road_infra_name = "INFRA_Roads"
    road_infra = os.path.join(road_infra_path, road_infra_name)
    
    road_two_four = os.path.join(fullpath, "Two_FourDigitRoads")
    
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
    arcpy.FeatureClassToFeatureClass_conversion(rdMakeRoute, road_infra_path, road_infra_name)
    
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

def create_pab_union(gdb_path, pab, buffer_dist):
    """
    buffer the project area boundary, union, and calculate IN_AND_OUT
    """
    pab_buffer = os.path.join(fullpath, "BufferPAB")
    pab_union = os.path.join(fullpath, "RoadClipBuffer")
    
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

### ------------ End Functions -----------  ###
    
#Create new mxd
mxd = arcpy.mapping.MapDocument("CURRENT")
df = arcpy.mapping.ListDataFrames(mxd)[0]

#------------- Copy PAB to the project gdb and buffer PAB-----------------
# we could set  arcpy.env.outputCoordinateSystem = sr
# and copy features will use this and reproject
arcpy.AddMessage("Copying project area boundary")

pabName = "PAB"

pabLyr = arcpy.mapping.Layer(project_area_bound)


pab_union = create_pab_union(fullpath,
                             project_area_bound, "1 Miles")

arcpy.AddMessage("Processing INFRA Roads")
infraRds(rdConv, rdDecom, rdExist, rdPlan, 
             fullpath, pab_union, roads, roadSym, mxd,
             fullpath)



