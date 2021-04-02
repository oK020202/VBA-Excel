import sys
from railroadtraffic import RailRoadsLinesWithTraffic

# Original shapefile
orgdata_path = 'C:\\Users\\oiax\\Documents\\RailLinesWithTraffic\\N02-19_GML\\'
path_to_section_layer_0 = orgdata_path + 'N02-19_RailroadSection.shp'
path_to_station_layer_0 = orgdata_path + 'N02-19_Station.shp'

# Work folder path.
workfolpath = 'C:\\Users\\oiax\\Documents\\RailLinesWithTraffic\\PyQGIS\\work\\'

# Middle shape files.
path_to_section_layer_1 = workfolpath + 'RailroadSection.shp'
path_to_station_layer_1 = workfolpath + 'RailroadStation.shp'

# create object.
rtrfcobj = RailRoadsLinesWithTraffic()

try:
    # create work layer.
    rtrfcobj.createlayer(path_to_section_layer_0, path_to_station_layer_0, path_to_section_layer_1, path_to_station_layer_1)
except LayerLoadError as instance:
    print(instance)
    sys.exit(1)
