#
#  Update Field Aliases in an MXD
#
# based on these links:
#
#   https://tereshenkov.wordpress.com/2016/01/16/accessing-arcobjects-in-python/
#   https://tereshenkov.wordpress.com/2017/03/01/update-field-aliases-in-map-document-layers/
#   https://www.esri.com/arcgis-blog/products/3d-gis/3d-gis/arcgis-desktop-and-vba-moving-forward/
#   https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#Carto.htm
#
#  Run on GTS 10.6 server
#  Tested with Python 2.7.14
#
import os
import xlrd
import argparse
import sys

from collections import defaultdict
from comtypes.client import GetModule, CreateObject
from snippets106 import GetStandaloneModules, InitStandalone

#----------------------------------------------------------------------
def getAliasForField(fieldName,aliasDict,defaultAliasDict):
    try:
        return aliasDict[fieldName]
    except KeyError:
        try:
            return defaultAliasDict[fieldName]
        except KeyError:
            return fieldName.replace('_',' ').title()
def getVisibilityForField(fieldName,visibleDict,defaultVisibleDict):
    try:
        return visibleDict[fieldName]
    except KeyError:
        try:
            return defaultVisibleDict[fieldName]
        except KeyError:
            if fieldName[-6:].upper() == '_SYSID':
                return False
            else:
                return True

def updateMXDLayerFieldAliases(mxdPath,xlsxPath,outputMXDPath=None):
    esriCartoOLBPath = r"E:\sw_nt\arcgis\Desktop10.6\com\esriCarto.olb"
    defaultAliasDict = {'GEOMETRY':'Geometry',
                        'SHAPE':'Shape',
                        'OBJECTID':'ObjectId',
                        'SE_ANNO_CAD_DATA':'SE_Anno_CAD_Data'}
    defaultVisibleDict = {'GEOMETRY':False,
                        'SHAPE':False,
                        'OBJECTID':False,
                        'SE_ANNO_CAD_DATA':False,
                        'GEOMETRY.AREA':False,
                        'GEOMETRY.LEN':False,
                        'SHAPE.AREA':False,
                        'SHAPE.LEN':False,}
    visibleDict = {}
    aliasDict = {}
    try:
        workbook = xlrd.open_workbook(xlsxPath)
    except:
        print "Error opening "+xlsxPath+". Does it exist?"
        sys.exit(1)
    sheet = workbook.sheet_by_index(0)
    for i in range(sheet.nrows):
        key = str(sheet.cell(i,0).value)
        alias = str(sheet.cell(i,1).value)
#        print key+'|'+alias+'|'+str(sheet.cell(i,2).value)+'|'+str(sheet.cell(i,2).value)[0:1]+'|'
        if str(sheet.cell(i,2).value) is None:
            visible = False
        elif str(sheet.cell(i,2).value)[0:1] in ['N','F','n','f']:
            visible = False
        else:
            visible = True
        aliasDict[key] = alias
        visibleDict[key] = visible
#        print key+'|'+alias+'|'+str(visible)+'|'

    GetStandaloneModules()
    InitStandalone()
    esriCarto = GetModule(esriCartoOLBPath)
    mapDocument = CreateObject(esriCarto.MapDocument, interface=esriCarto.IMapDocument)
    try:
        mapDocument.Open(mxdPath)
    except:
        print "Error opening "+mxdPath+". Does it exist?"
        sys.exit(1)
    map = mapDocument.Map(0)
    enumLayer = map.Layers(None,True)
    layer = enumLayer.Next()
    while layer:
       try:
            fields = layer.QueryInterface(esriCarto.ILayerFields)
            print "Processing layer: ", layer.Name
       except:   # probably a group layer - skip
            layer = enumLayer.Next()
            continue

       for i in xrange(fields.FieldCount):
            fields.FieldInfo(i).Visible = getVisibilityForField(fields.Field(i).Name,visibleDict,defaultVisibleDict);
            fields.FieldInfo(i).alias = getAliasForField(fields.Field(i).Name,aliasDict,defaultAliasDict);
#            print fields.Field(i).Name,fields.FieldInfo(i).alias,fields.FieldInfo(i).Visible

       layer = enumLayer.Next()

    if not outputMXDPath:
        mapDocument.Save()
    else:
        mapDocument.SaveAs(outputMXDPath)
    return


argParser = argparse.ArgumentParser(description='Updates the field aliases in all layers in an MXD file')
argParser.add_argument('-s', dest='srcMXD', action='store', default=None, required=True, help='the source MXD')
argParser.add_argument('-x', dest='srcXLSX', action='store', default=None, required=True, help='the xlsx file containing aliases and visibility flags (can be empty)')
argParser.add_argument('-d', dest='destMXD', action='store', default=None, required=False, help='the target MXD - optional; if not specified then the source MXD is overwritten')

try:
  args = argParser.parse_args()
except argparse.ArgumentError as e:
  argParser.print_help()
  sys.exit(1)

mxd_path = args.srcMXD
xlsx_path = args.srcXLSX
if args.destMXD is None:
    updateMXDLayerFieldAliases(mxd_path,xlsx_path)
else:
    dest_mxd_path = args.destMXD
    updateMXDLayerFieldAliases(mxd_path,xlsx_path,dest_mxd_path)
sys.exit(0)
