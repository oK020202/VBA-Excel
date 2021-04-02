import sys
from railroadtraffic import RailRoadsLinesWithTraffic

# Work folder path.
workfolpath = 'C:\\Users\\oiax\\Documents\\RailLinesWithTraffic\\PyQGIS\\work\\'

# Excel book path.
exbokpath = workfolpath + '12_駅別発着・駅間通過人員表_首都圏.xlsx'

# Middle shapefiles.
shfshppath = workfolpath + 'RailroadSection_2.shp'
stashppath = workfolpath + 'RailroadStation_2.shp'

# Output shapefiles.
outscshppath = workfolpath + 'traffic.shp'
outstshppath = workfolpath + 'station.shp'

# create object.
rtrfcobj = RailRoadsLinesWithTraffic()

try:
    rtrfcobj.creategraph(exbokpath, outscshppath, outstshppath, shfshppath, stashppath)
except Exception as instance:
    print(instance)
    sys.exit(1)
