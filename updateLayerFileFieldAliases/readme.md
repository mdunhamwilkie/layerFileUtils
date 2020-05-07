arcPy does not provide methods for modifying individual aliases in a layer file. However, this [blog article](https://tereshenkov.wordpress.com/2017/03/01/update-field-aliases-in-map-document-layers/) explains how to do this by calling the [ArcObjects .Net methods](https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#Carto.htm) from python using the [comtypes package](https://pypi.org/project/comtypes/).  

With ArcObjects you can do a lot of things that arcPy cannot do (e.g., update layer file aliases, create MXD files, etc.) 
For BC government GIS users these python 2 files can be run on a GTS 10.6 server (they rely on ArcObjects .Net having been installed, which is the case for GTS 10.6 servers)

There are two files:
1.	updateMXDLayerFieldAliases.py
2.	snippets106.py

Put these two files in the same directory.  The snippets file includes a bunch of helper functions, taken straight from the blog article.  The updateMXDLayerFieldAliases.py is what I adapted from the example in the article.

updateMXDLayerFieldAliases.py takes as input an MXD file with one or more layer files added to it. It also takes as input a (possibly empty) Excel file (the use of which is explained below).  An optional 3rd argument is the name of an MXD file to write the result to if you want to leave the input MXD unchanged. If this 3rd argument is not specified, the input MXD is overwritten.

The default behaviour of updateMXDLayerFieldAliases.py is to turn all fields in all of the layer files to visible, except:
                GEOMETRY
                SHAPE
                The *.len and *.area fields
                OBJECTID
                The _SYSID field
                SE_ANNO_CAD_DATA

The aliases generated are the column names, converted to title case, with all “_” changed to a space.

If you want to change this behaviour, supply an Excel file with:
                Column 1 set to a BCGW Column name referenced in the layer file
                Column 2 set to the desired alias
                Column 3 set to something starting with ‘N’,’n’,’F’,’f’ if you want to turn the visibility of that field off. Anything else (including a blank) will turn the visibility for that field on.
An Excel file has to be specified, but you can specify an empty one if you want to accept the default behaviour for all fields.

Example usages:

1.	Just accept the default behaviour and overwrite the input mxd:
python updateMXDLayerFieldAliases.py -s test_update_labels.mxd  -x empty.xlsx

2.	Override some of the default behaviour and write to a new mxd:
python updateMXDLayerFieldAliases.py -s test_update_labels.mxd  -x aliases_and_visibility_flags.xlsx -d test_update_labels_upd.mxd

