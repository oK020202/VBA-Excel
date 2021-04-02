import win32com.client
from qgis.core import *
from PyQt5.QtCore import *
import qgis_init
import datetime
from qgis.analysis import *
import sys

# Error Classes
class LayerLoadError(Exception):
    pass
    
class NoStationPointError(Exception):
    pass

class NotFoundStationError(Exception):
    pass

class NotFoundSectionError(Exception):
    pass

class FailCreateWriter(Exception):
    pass

# Class for a station.
class Station(object):
    # Constructor.
    # Argument splayer is object for the layer of PointXYs at the stations.
    # range is reference for excel cell in the worksheet of traffic census.
    # name is work name of station.
    # linelist is list storing the names of railroad lines at the original shapefile.
    def __init__(self, splayer, range, name, company, linelist):
        # List of PointXYs at the middles of sections on station.
        self.stplist = []
        
        # Fetch the points of station.
        stname = name
        spfeatures = splayer.getFeatures()
        
        for spfeature in spfeatures:
            if str(spfeature.attribute('Station')) == stname \
                and str(spfeature.attribute('Company')) == company:
                if len(linelist) > 0:
                    for linename in linelist:
                        if str(spfeature.attribute('Line')) == linename:
                            # Same point as already resisted one is not added to list.
                            self.addPoint(spfeature)
                else:
                    self.addPoint(spfeature)
        
        print('Get %u points of the station %s. ' % (len(self.stplist), stname))
        
        # When the station is not founded, raise error.
        if len(self.stplist) < 1:
            raise NoStationPointError('Points of station ' + stname + ' do not exist.')
        
        # Set the attributes of station.
        # Flag for switch back station.
        self.sbflg = False
        
        # Station name for work.
        self.workname = name
        
        # Station name.
        self.name = range.Text
        
        # Passenger number to take on outbound trains at the station.
        self.outboundon = range.GetOffset(RowOffset = 0, ColumnOffset = 13).Value
        
        # Passenger number to take off outbound trains at the station.
        self.outboundoff = range.GetOffset(RowOffset = 0, ColumnOffset = 14).Value
        
        # Passenger number to take outbound trains between next station.
        self.outboundpass = range.GetOffset(RowOffset = 0, ColumnOffset = 15).Value
        
        # Passenger number to take on inbound trains at the station.
        self.inboundon = range.GetOffset(RowOffset = 0, ColumnOffset = 16).Value
        
        # Passenger number to take off inbound trains at the station.
        self.inboundoff = range.GetOffset(RowOffset = 0, ColumnOffset = 17).Value
        
        # Passenger number to take inbound trains between next station.
        self.inboundpass = range.GetOffset(RowOffset = 0, ColumnOffset = 18).Value
        
        # Passenger number to take trains between next station.
        self.bothpass = self.outboundpass + self.inboundpass
        
        # Object of station for next station.
        self.next = None
        
        # Object of section of railroad for next station.
        self.sectionschain = None
        
        # Object of section of railroad for backstation.
        self.previoussection = None
        
    
    # Method for the PointXY from feature on station to list for PointXYs.
    def addPoint(self, spfeature):
        # Same point as already resisted one is not added to list.
        newflg = True
        astp = spfeature.geometry().asPoint()
        for stp in self.stplist:
            if stp.compare(astp) == True:
                newflg = False
                break
        
        if newflg == True:
            self.stplist.append(astp)


# Class to connect sections betwen two stations.
class RailroadSections(object):
    # Constructor.
    # Argument sslayer is layer for linestrings of the sections on railroads.
    # splayer is layer for Pointxys of the stations.
    # backsection is object of RailroadSections for backword station.
    # depth is position in the object chains of RailroadSections class.
    def __init__(self, sslayer, stptlist, backsection = None, depth = 1):
        self.stptlist = stptlist
        self.sslayer = sslayer
        self.scgeomlist = []
        self.next = None
        self.back = backsection
        self.endstptlist = []
        self.depth = depth
    
    # Method to connect section linestrings from points of start station.
    # Argument endstptlist is list for the PointXYs at the next station.
    # sbflg is flag for the switch back station of the start end.
    def createsectionlistforward(self, endstptlist, sbflg = False):
        edptlist = []
        for point in self.stptlist:
            ssfeatures = self.sslayer.getFeatures()
            for ssfeature in ssfeatures:
                workptxys = ssfeature.geometry().asMultiPolyline()[0]
                
                # Check the start point of the linstring.
                if workptxys[0].compare(point) == True:
                    
                    # Geometries in side of start station are not taken
                    # except switch back station.
                    backflg = False
                    if self.back != None and sbflg == False:
                        for backgeom in self.back.scgeomlist:
                            if backgeom.equals(ssfeature.geometry()):
                                backflg = True
                                break
                    if backflg == False:
                        # Geometry equal to already taken is not taken.
                        apsflg = False
                        for scgeom in self.scgeomlist:
                            if scgeom.equals(ssfeature.geometry()) == True:
                                apsflg = True
                                break
                        
                        # Take the end points.
                        # Points equal to already taken is not taken.
                        if apsflg == False:
                            self.scgeomlist.append(ssfeature.geometry())
                            appflg = False
                            for edpt in edptlist:
                                if edpt.compare(workptxys[len(workptxys) - 1]) == True:
                                    appflg = True
                                    break
                            if appflg == False:
                                edptlist.append(workptxys[len(workptxys) - 1])
                                for edstapt in endstptlist:
                                    if workptxys[len(workptxys) - 1].compare(edstapt) == True:
                                        self.endstptlist.append(edstapt)
                
                # Check the end point of the linstring.
                elif workptxys[len(workptxys) - 1].compare(point) == True:
                
                    # Geometries in side of start station are not taken.
                    backflg = False
                    if self.back != None and sbflg == False:
                        for backgeom in self.back.scgeomlist:
                            if backgeom.equals(ssfeature.geometry()):
                                backflg = True
                                break
                    if backflg == False:
                        # Geometry equal to already taken is not taken.
                        apsflg = False
                        for scgeom in self.scgeomlist:
                            if scgeom.equals(ssfeature.geometry()) == True:
                                apsflg = True
                                break
                            
                        # Take the end points.
                        # Points equal to already taken is not taken.
                        if apsflg == False:
                            self.scgeomlist.append(ssfeature.geometry())
                            appflg = False
                            for edpt in edptlist:
                                if edpt.compare(workptxys[0]) == True:
                                    appflg = True
                                    break
                            if appflg == False:
                                edptlist.append(workptxys[0])
                                for edstapt in endstptlist:
                                    if workptxys[0].compare(edstapt) == True:
                                        self.endstptlist.append(edstapt)
        
#        print(self.scgeomlist)
        
        # In arrival of the end station points, exit function.
        if len(self.endstptlist) > 0:
            return self
        
        # Not found the next step linestring.
        if len(self.scgeomlist) < 1:
            return None
        
        if self.depth + 1 > 99:
            raise NotFoundSectionError('Path can not be found between stations.')
        else:
            # Next linestrings search.
            self.next = RailroadSections(self.sslayer, edptlist, self, self.depth + 1)
            lastsection = self.next.createsectionlistforward(endstptlist)
            return lastsection
    
    # Method to select linestrings connect between two stations.
    # Argument endpointlist is list selected points by method createsectionlistforward().
    def selectsectionlistbackward(self, endpointlist):
        # Erase linestring geometries which do not have end point same as argument.
        stptlist = []
        geomidx = 0
#        print(endpointlist)
        while geomidx < len(self.scgeomlist):
            scgeom = self.scgeomlist[geomidx]
            onflg = False
            workptxys = scgeom.asMultiPolyline()[0]
            for edpt in endpointlist:
                print('%s %s ' % (workptxys[0], workptxys[len(workptxys) - 1]))
                if workptxys[0].compare(edpt) == True:
                    onflg = True
                    stptlist.append(workptxys[len(workptxys) - 1])
                elif workptxys[len(workptxys) - 1].compare(edpt) == True:
                    onflg = True
                    stptlist.append(workptxys[0])
            
            if onflg == False:
                print('Delete geometry')
                del self.scgeomlist[geomidx]
            else:
                geomidx += 1
        
        # Erase startpoints.
        stptidx = 0
#        print(self.stptlist)
        while stptidx < len(self.stptlist):
            inflg = False
            for stpt in stptlist:
                if stpt.compare(self.stptlist[stptidx]) == True:
                    inflg = True
                    break
            if inflg == False:
                print('Delete point')
                del self.stptlist[stptidx]
            else:
                stptidx += 1
        
        # Duplicate geometries are marged.
        geomidx1 = 0
        while geomidx1 < len(self.scgeomlist) -1:
            geomidx2 = len(self.scgeomlist) - 1
            while geomidx2 > geomidx1:
#                print('%d %d %d' %  (len(self.scgeomlist), geomidx1, geomidx2)  )
                if self.scgeomlist[geomidx1].equals(self.scgeomlist[geomidx2]):
                    del self.scgeomlist[geomidx2]
                geomidx2 -= 1
            geomidx1 += 1
        
        print(self.stptlist)
        print(self.scgeomlist)
        
        # To exit when calling object is the first of chain 
        if self.depth < 2:
            return self
        
        # To call selectsectionlistbackward() method of the object for the back section.
        else:
            firstsection = self.back.selectsectionlistbackward(self.stptlist)
            return firstsection
    

# Class for a railroad line.
class RailRoadsLine(object):
    # Constructor.
    def __init__(self):
        self.linename = None
        self.stobjlist = []
        self.sslayer = None
    
    # Function to take attributes of stations from excel sheet and geometries of stations from layer.
    # Argument of range is reference for the excel cell in the worksheet of railroad traffic.
    # sslayer is layer object for sections of railroad.
    # splayer is layer object for PointXYs of stations.
    def createstations(self, range, sslayer, splayer):
        rg = range
        company = None
        linelist = []
        backstlist = []
        
        # There are borders on the bottom of the last row cells.
        while rg.Borders(9).LineStyle == -4142:
            if len(rg.Text) > 0:
                # Get the attributes of line.
                # Strings in cells of railroad line name have under line.
                if rg.Font.Underline == 2:
                    self.linename = rg.Text
                    print(rg.Text)
                    # Create new layer sections along the line.
                    # At first, get the company name of rail line.
                    co = 19
                    company = rg.GetOffset(RowOffset = 0, ColumnOffset = co).Text
                    if len(company) < 1:
                        raise NotFoundStationError('Company name is not found on line of ' + rg.Text + '.')
                    
                    # When names of lines are also written, get them.
                    while True:
                        co += 1
                        linename = rg.GetOffset(RowOffset = 0, ColumnOffset = co).Text
                        if len(linename) > 0:
                            linelist.append(linename)
                        else:
                            break
                    
                    # Crate work layer.
                    uri = 'LineString?crs=epsg:6668'
                    self.sslayer = QgsVectorLayer(uri, "Railroad line layer",  "memory")
                    
                    # Set new layer's attributes.
                    sspr = self.sslayer.dataProvider()
                    sspr.addAttributes([QgsField('Company', QVariant.String, 'String', 254), 
                        QgsField('Line', QVariant.String, 'String', 254),
                        QgsField('Station', QVariant.String, 'String', 254)])
                        
                    self.sslayer.updateFields()
                    
                    # Take feature from railroad sections layer.
                    ssfeatures = sslayer.getFeatures()
                    ssflist = []
                    if len(linelist) > 0:
                        for ssfeature in ssfeatures:
                            for linename in linelist:
                                if ssfeature.attribute('Company') == company \
                                    and ssfeature.attribute('Line') == linename:
                                        ssflist.append(ssfeature)
                    else:
                        for ssfeature in ssfeatures:
                            if str(ssfeature.attribute('Company')).encode(encoding="utf-8", errors="strict").strip() == company.encode(encoding="utf-8", errors="strict"):
                                ssflist.append(ssfeature)
                    
                    sspr.addFeatures(ssflist)
                    self.sslayer.updateExtents()
                    
#                    shpfpath = workfolpath + self.linename + '.shp'
                    
#                    transform_context = QgsProject.instance().transformContext()
#                    save_options = QgsVectorFileWriter.SaveVectorOptions()
#                    save_options.driverName = 'ESRI Shapefile'
#                    save_options.fileEncoding = 'UTF-8'
#                    QgsVectorFileWriter.deleteShapeFile(shpfpath)
#                    QgsVectorFileWriter.writeAsVectorFormatV2(self.sslayer, shpfpath, transform_context, save_options)
                    
                # Row under the end station of line.
                elif rg.Text == '合計':
                    rg = rg.GetOffset(RowOffset = 1)
                    backstlist = []
                    break
                
                # Row on the stations.
                else:
                    try:
                        co = 19
                        srg = rg.GetOffset(RowOffset = 0, ColumnOffset = co)
                        while len(srg.Text) > 0:
                            # Create the object for the station.
                            st= Station(splayer, rg, srg.Text, company, linelist)
                            self.stobjlist.append(st)
                            
                            # Setting of member next of back station object.
                            # Member next must not be a work station.
                            if co < 20 and len(backstlist) > 0:
                                for backst in backstlist:
                                    backst.next = st
                                backstlist = []
                            backstlist.append(st)
                            co += 1
                            srg = rg.GetOffset(RowOffset = 0, ColumnOffset = co)
                            
                            # When flag of switch back station is specified.
                            if srg.Text == 'SB':
                                st.sbflg = True
                                co +=1
                                srg = rg.GetOffset(RowOffset = 0, ColumnOffset = co)
                            
                    except NoStationPointError:
                        raise
            rg = rg.GetOffset(RowOffset = 1)
        return rg
    
    # Function to search path between stations.
    # Argument trfields is object of Fields in the output layer of linestrings.
    # trwriter is VectorFileWriter object of output layer for linestrings..
    # stfields is object of Fields in output layer of Points at stations.
    # stwriter is VectorFileWriter object of output layer for Points at stations.
    def searchroute(self, trfields, trwriter, stfields, stwriter):
        # Search the shortest path between two stations.
        st2idx = 1
        st1 = self.stobjlist[st2idx - 1]
        st1ptlist = st1.stplist
        lastsection = None
        retryflg = False
        while True:
            firstsection = RailroadSections(self.sslayer, st1ptlist, lastsection)
            st2 = self.stobjlist[st2idx]
            print(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))
            print('Searching route between ' + self.linename + ' ' + st1.workname + ' and ' + st2.workname + '.')
            if retryflg == False:
                st2ptlist = st2.stplist
                st1.previoussection = lastsection
                
#            print(st2ptlist)
            # Search route from start station.
            try:
                lastsection = firstsection.createsectionlistforward(st2ptlist, st1.sbflg)
            except NotFoundSectionError as instance:
                raise
            
            # When the route is not found, take back one station and search path again without start station point.
            if lastsection == None:
                # When good path is not found, take back one station and search path again without start station point.
                if st2idx > 1:
                    st2ptlist = []
                    if len(st1.stplist) > 1:
                        pidx = 0
                        while True:
                            opt = st1.stplist[pidx]
                            delflg = False
                            for pt in st1ptlist:
                                if pt.compare(opt) == True:
                                    del st1.stplist[pidx]
                                    delflg = True
                            
                            if len(st1.stplist) < 1:
                                st2idx -= 1
                                st1 = self.stobjlist[st2idx - 1]
                                raise NotFoundSectionError('Path can not be found between stations of ' + st1.workname + ' and ' + st1.next.workname + ' again.')
                            
                            if delflg == False:
                               st2ptlist.append(opt)
                               pidx += 1
                               
                            if pidx > len(st1.stplist) - 1:
                                break
                        
                        st2idx -= 1
                        st1 = self.stobjlist[st2idx - 1]
                        st1ptlist = st1.sectionschain.stptlist
                        lastsection = st1.previoussection
                        retryflg = True
                        if len(st2ptlist) == 0:
                            raise NotFoundSectionError('Path can not be found between stations of ' + st1.workname + ' and ' + st1.next.workname + ' again.')
                            
                        print('Search the shortest route between ' + self.stobjlist[st2idx - 1].workname + ' and ' + self.stobjlist[st2idx].workname + ' again.')
                    else:
                        raise NotFoundSectionError('Path can not be found between stations of ' + st1.workname + ' and ' + st2.workname + '.')
                    continue
                
                # Path is not fount from the first station, process is aborted.
                else:
                    raise NotFoundSectionError('Path can not be found between stations of ' + st1.workname + ' and ' + st2.workname + '.')
            
            # Confirm the route from end station.
            retryflg = False 
#            print(lastsection.endstptlist)
            lastsection.selectsectionlistbackward(lastsection.endstptlist)
            
            # The shortest path is found, points of path is recorded to the start station object
            # and prceed one station.
            print('Shortest path is found.')
#            print(lastsection.endstptlist)
            st1.sectionschain = firstsection
            st1ptlist = lastsection.endstptlist
            st2idx += 1
            st1 = st2
                
            # When station is last station in the railline, exit this loop.
            if st2idx == len(self.stobjlist):
                break
        
        # Clear sections not on the shortest path.
        st2idx -= 1
        edptlist = lastsection.endstptlist
        while True:
            st2name = self.stobjlist[st2idx].workname
            st2idx -= 1
            st1 = self.stobjlist[st2idx]
            print('Clearing process between ' + st1.workname + ' and ' + st2name + '.')
            firstsection = lastsection.selectsectionlistbackward(edptlist)
            st1.stplist = firstsection.stptlist
            if st2idx < 1:
                break
            lastsection = firstsection.back
            edptlist = st1.stplist
        
        # Create features with attribute of transit and geometry of linestring between stations.
        sidx = 0
        sectionsobj = None
        st = None
        while sidx < len(self.stobjlist) - 1:
            st = self.stobjlist[sidx]
            if st.sectionschain == None:
                break
            if st.name == st.workname:
                print('Creating features at ' + st.name + ' station.')
                for point in st.stplist:
                    stfeature = QgsFeature()
                    stfeature.setGeometry(QgsGeometry.fromPointXY(point))
                    stfeature.setFields(stfields)
                    stfeature.setAttributes([self.linename, st.name, st.outboundon, st.outboundoff, \
                        st.inboundon, st.inboundoff])
                    stwriter.addFeature(stfeature)
                    
            sidx += 1
            print('Creating features between ' + st.workname + ' and ' + st.next.workname + ' on ' + self.linename + '.')
            
            # Take geometry from sections object.
            sectionsobj = st.sectionschain
            while True:
                for sectiongeom in sectionsobj.scgeomlist:
                    trfeature = QgsFeature()
                    trfeature.setGeometry(sectiongeom)
                    trfeature.setFields(trfields)
                    trfeature.setAttributes([self.linename, st.name, st.next.name, st.outboundon, st.next.outboundoff, \
                        st.outboundpass, st.next.inboundon, st.inboundoff, st.inboundpass, st.bothpass])
                    trwriter.addFeature(trfeature)
                if sectionsobj.next == None:
                    break
                else:
                   sectionsobj = sectionsobj.next
            
        # In stations on the first and the last, section linestrings of ends are added to features.
        if self.stobjlist[0].name != self.stobjlist[sidx].name:
            # In the last station.
            print('Creating features at ' + st.name + ' station.')
            st = self.stobjlist[sidx]
            for point in sectionsobj.endstptlist:
                stfeature = QgsFeature()
                stfeature.setGeometry(QgsGeometry.fromPointXY(point))
                stfeature.setFields(stfields)
                stfeature.setAttributes([self.linename, st.name, st.outboundon, st.outboundoff, \
                    st.inboundon, st.inboundoff])
                stwriter.addFeature(stfeature)
            
            print('Add feature on the end of line.')
            st = self.stobjlist[sidx - 1]
            self.addendfeature(st, sectionsobj.scgeomlist, sectionsobj.endstptlist, trfields, trwriter)
            
            # In the first station.
            print('Add feature on the start of line.')
            sidx = 0
            st = self.stobjlist[sidx]
            print(st.sectionschain.scgeomlist)
            self.addendfeature(st, st.sectionschain.scgeomlist, st.sectionschain.stptlist, trfields, trwriter)
        
        # Section linestrings of ends are added to features in swich back stations.
        sidx = 1
        while sidx < len(self.stobjlist) - 1:
            st = self.stobjlist[sidx]
            if st.sbflg == True:
                # For the outbound sections.
                print('Add feature on ' + st.workname + ' station.')
                self.addendfeature(st, st.sectionschain.scgeomlist, st.sectionschain.stptlist, trfields, trwriter)
                
                # For the inbound sections.
                geomlist = st.sectionschain.back.scgeomlist
                pointlist = st.sectionschain.stptlist
                st = self.stobjlist[sidx -1]
                self.addendfeature(st, geomlist, pointlist, trfields, trwriter)
                
            sidx += 1
    
    
    # Method to add features at the end stations of the railroad line.
    # Argument st is the object for the station.
    # addedgeomlist is list of linestrings to make shortest path.
    # trfields is the object of Fields for linestrings.
    # trwriter is VectorFileWriter object of linestrings.
    def addendfeature(self, st, addedgeomlist, pointlist, trfields, trwriter):
        for sp in pointlist:
            ssfeatures = self.sslayer.getFeatures()
            stflg = False
            
            print(sp)
            for ssfeature in ssfeatures:
                workptxys = ssfeature.geometry().asMultiPolyline()[0]
                addedflg = False
                # When either end point is contented to the linestrings, it is added to features without already added linestring.
                if workptxys[0].compare(sp) == True:
                    for addedgeom in addedgeomlist:
                        if addedgeom.equals(ssfeature.geometry()) == True:
                            addedflg = True
                            break
                    if addedflg == True:
                        continue
                    else:
                        trfeature = QgsFeature()
                        trfeature.setGeometry(ssfeature.geometry())
                        trfeature.setFields(trfields)
                        trfeature.setAttributes([self.linename, st.name, st.next.name, st.outboundon, st.next.outboundoff, \
                            st.outboundpass, st.next.inboundon, st.inboundoff, st.inboundpass, st.bothpass])
                        trwriter.addFeature(trfeature)
                        print(ssfeature.geometry())
                        stflg = True
                        break
                    
                if workptxys[len(workptxys) - 1].compare(sp) == True:
                    for addedgeom in addedgeomlist:
                        if addedgeom.equals(ssfeature.geometry()) == True:
                            addedflg = True
                            break
                    if addedflg == True:
                        continue
                    else:
                        trfeature = QgsFeature()
                        trfeature.setGeometry(ssfeature.geometry())
                        trfeature.setFields(trfields)
                        trfeature.setAttributes([self.linename, st.name, st.next.name, st.outboundon, st.next.outboundoff, \
                            st.outboundpass, st.next.inboundon, st.inboundoff, st.inboundpass, st.bothpass])
                        trwriter.addFeature(trfeature)
                        print(ssfeature.geometry())
                        stflg = True
                        break
                        
                if stflg == True:
                    break
            print(stflg)

# Class for the comprehensive reailroad lines.
class RailRoadsLinesWithTraffic(object):
    # Constructor.
    def __init__(self):
        qgis_init.setenvirons()
        
        # Supply the path to the qgis install location
        QgsApplication.setPrefixPath("C:\Program Files\QGIS 3.10", True)
        
        # Create a reference to the QgsApplication.
        # Setting the second argument to True enables the GUI.  We need
        # this since this is a custom application.
        self.qgs = QgsApplication([], True)
        
        self.crs = QgsCoordinateReferenceSystem("EPSG:6668")
        
        # load providers
        self.qgs.initQgis()
        
        self.transform_context = QgsProject.instance().transformContext()
        self.save_options = QgsVectorFileWriter.SaveVectorOptions()
        self.save_options.driverName = 'ESRI Shapefile'
        self.save_options.fileEncoding = 'UTF-8'
        
        self.sslayer = None
        self.splayer = None
        
        self.xl = None
        self.wb = None
        self.scwriter = None
        self.stwriter = None
    
    # Destructor.
    def __del__(self):
        # Close process of excel.
        if self.wb == None:
            pass
        else:
            self.wb.Close(SaveChanges=False)
            self.xl.quit()
        
        # delete the writer to flush features to disk
        if self.scwriter == None:
            pass
        else:
            del self.scwriter
        
        if self.stwriter == None:
            pass
        else:
            del self.stwriter
        
        self.qgs.exitQgis()
        
    
    # Method to output the layer file for the work.
    # Argument secshppath is file path of the shapefile for sections of railroad.
    # stashppath is file path of the shapefile for stations of railroad.
    # worksecshppath is output file path of the shapefile for new sections of railroad.
    # workstashppath is output file path of the shapefile for points of stations.
    def createlayer(self, secshppath, stashppath, worksecshppath = None, workstashppath = None):
        # Open original layer for sections of railrosds.
        sclayer = QgsVectorLayer(secshppath, "Section layer", "ogr")
        if not sclayer.isValid():
            raise LayerLoadError('Layer:' + secshppath + ' failed to load!')
        
        # Open original layer for sections of stations.
        stlayer = QgsVectorLayer(stashppath, "Station layer", "ogr")
        if not stlayer.isValid():
            raise LayerLoadError('Layer:' + stashppath + ' failed to load!')
        
        # Create new layer sections on station is divided as half length.
        uri = 'LineString?crs=epsg:6668'
        self.sslayer = QgsVectorLayer(uri, "Station Section devided Layer",  "memory")
        uri = 'Point?crs=epsg:6668'
        self.splayer = QgsVectorLayer(uri, "Station Point Layer",  "memory")
        
        # Set new layer's attributes.
        sspr = self.sslayer.dataProvider()
        sspr.addAttributes([QgsField('Company', QVariant.String, 'String', 254), 
            QgsField('Line', QVariant.String, 'String', 254),
            QgsField('Station', QVariant.String, 'String', 254)])
        
        self.sslayer.updateFields() # tell the vector layer to fetch changes from the provider
        
        sppr = self.splayer.dataProvider()
        sppr.addAttributes([QgsField('Company', QVariant.String, 'String', 254), 
            QgsField('Line', QVariant.String, 'String', 254),
            QgsField('Station', QVariant.String, 'String', 254)])
        
        self.splayer.updateFields() # tell the vector layer to fetch changes from the provider
        
        # Check all features on layer of railroads sections.
        scfeatures = sclayer.getFeatures()
        
        ssflist = []
        spflist = []
        
        for scfeature in scfeatures:
            scgeom = scfeature.geometry()
            stflg = False
            
            # Check features on layer of reilroads stations.
            stfeatures = stlayer.getFeatures()
            for stfeature in stfeatures:
                stgeom = stfeature.geometry()
                if scgeom.equals(stgeom):
                    if (scfeature.attribute('N02_003') == stfeature.attribute('N02_003')
                        and scfeature.attribute('N02_004') == stfeature.attribute('N02_004')):
                        # When the feature of section is same geometry and attributes as one of station, devide the section linestring to half.
                        print(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))
                        print('Create station linestrings of ' + stfeature.attribute('N02_005') + ' on ' + stfeature.attribute('N02_004') + ' ' + stfeature.attribute('N02_003'))
                        # Take vetice points along the section.
                        workpts = []
                        # Method of 'asMultistring' returns two dimensional list of points.
                        for workptxyslist in scgeom.asMultiPolyline():
                           for workptxy in workptxyslist:
                               cpflg = False
                               # same points is not added to list of workpts.
                               for workpt in workpts:
                                   if workptxy.compare(QgsPointXY(workpt)) == True:
                                       cpflg = True
                                       break
                               if cpflg == False:
                                   workpts.append(QgsPoint(workptxy))
                        # Check whether the half point is located on the vertex of linestring of section.
                        workls = QgsLineString(workpts)
                        worklshalf = workls.curveSubstring(0, workls.length() * 0.5)
                        workpts1 = []
                        workpts2 = []
                        hfflg = False
                        for workpt in workpts:
                            if hfflg == True:
                                workpts2.append(workpt)
                            else:
                                workpts1.append(workpt)
                            if QgsPointXY(workpt).compare(QgsPointXY(worklshalf.endPoint())) == True:
                                hfflg = True
                                workpts2.append(workpt)
                        ls1 = QgsLineString(workpts1)
                        ls2 = QgsLineString(workpts2)
                        
                        # When the half point is not located on the vertex of linestring of section, divede linestring to half.
                        if hfflg == False:
                            ls1 = workls.curveSubstring(0, workls.length() * 0.5)
                            ls2 = workls.curveSubstring(workls.length() * 0.5, workls.length())
                            
                        # Create the new feature and set the geometry and attributes to it.
                        sthgeom1 = QgsGeometry.fromPolyline(ls1.points())
                        st1edp = QgsPointXY(ls1.endPoint())
                        ssfeature = QgsFeature()
                        ssfeature.setGeometry(sthgeom1)
                        ssfeature.setAttributes([stfeature.attribute('N02_004'), stfeature.attribute('N02_003'), stfeature.attribute('N02_005')])
                        ssflist.append(ssfeature)
                        
                        sthgeom2 = QgsGeometry.fromPolyline(ls2.points())
                        ssfeature = QgsFeature()
                        ssfeature.setGeometry(sthgeom2)
                        ssfeature.setAttributes([stfeature.attribute('N02_004'), stfeature.attribute('N02_003'), stfeature.attribute('N02_005')])
                        ssflist.append(ssfeature)
                        
                        # The half point on the linestring of station is set to the new feature.
                        spgeom = QgsGeometry.fromPointXY(st1edp)
                        spfeature = QgsFeature()
                        spfeature.setGeometry(spgeom)
                        spfeature.setAttributes([stfeature.attribute('N02_004'), stfeature.attribute('N02_003'), stfeature.attribute('N02_005')])
                        spflist.append(spfeature)
                        
                        stflg = True
                        break
            
            # Geometry of sections not on the station are directly set to the new feature.
            if stflg == False:
                workptxys = scgeom.asMultiPolyline()[0]
                workpts = []
                for workptxy in workptxys:
                    workpts.append(QgsPoint(workptxy))
                ssfeature = QgsFeature()
                ssfeature.setGeometry(QgsLineString(workpts))
                ssfeature.setAttributes([scfeature.attribute('N02_004'), scfeature.attribute('N02_003'), None])
                ssflist.append(ssfeature)
        
        # update layer's extent when new features have been added
        # because change of extent in provider is not propagated to the layer
        sspr.addFeatures(ssflist)
        sppr.addFeatures(spflist)

        # update layer's extent when new features have been added
        # because change of extent in provider is not propagated to the layer
        self.sslayer.updateExtents()
        self.splayer.updateExtents()
        
        # Write the middle result about sections to file.
        if worksecshppath != None:
            QgsVectorFileWriter.deleteShapeFile(worksecshppath)
            res = QgsVectorFileWriter.writeAsVectorFormatV2(self.sslayer, worksecshppath, self.transform_context, self.save_options)
            if res[0] != 0:
                raise FailCreateWriter('Error when creating shapefile: ' + worksecshppath + ' with ' + res[1])
        
        # Write the middle result about sections to file.
        if worksecshppath != None:
            QgsVectorFileWriter.deleteShapeFile(workstashppath)
            res = QgsVectorFileWriter.writeAsVectorFormatV2(self.splayer, workstashppath, self.transform_context, self.save_options)
            if res[0] != 0:
                raise FailCreateWriter('Error when creating shapefile: ' + workstashppath + ' with ' + res[1])
    
    
    # Method to create layer with traffic between stations.
    # Argument exbokpath is the path of excel book for the traffic between stations.
    # outscshppath is the path of output shapefile of linestrings with traffic data.
    # outstshppath is the path of output shapefile of stations' points with traffic data.
    # worksecshppath is the path of linestrings shapefile for output of method createlayer().
    # workstashppath is the path of points shapefile for output of method createlayer().
    def creategraph(self, exbokpath, outscshppath, outstshppath, worksecshppath = None, workstashppath = None):
        # Open the layer of railroads sections.
        if self.sslayer == None:
            self.sslayer = QgsVectorLayer(worksecshppath, "New Section layer", "ogr")
            if not self.sslayer.isValid():
                raise LayerLoadError('Section Layer for work:' + worksecshppath + ' failed to load!')
        
        # Open the layer of station sections.
        if self.splayer == None:
            self.splayer = QgsVectorLayer(workstashppath, "New Station layer", "ogr")
            if not self.splayer.isValid():
                raise LayerLoadError('Section Layer for work:' + worksecshppath + ' failed to load!')
        
        # Generate excel application and open the workbook with transit data between stations.
        self.xl = win32com.client.Dispatch("Excel.Application")
        
        self.wb = self.xl.Workbooks.Open(Filename=exbokpath, ReadOnly=1)
        
        ws = self.wb.worksheets(1)
        
        # Create fields of output layer.
        trfields = QgsFields()
        trfields.append(QgsField('LineName', QVariant.String, 'String', 254), QgsFields.OriginProvider, 1)
        trfields.append(QgsField('Station1', QVariant.String, 'String', 254), QgsFields.OriginProvider, 2)
        trfields.append(QgsField('Station2', QVariant.String, 'String', 254), QgsFields.OriginProvider, 3)
        trfields.append(QgsField('OutOn', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 4)
        trfields.append(QgsField('OutOff', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 5)
        trfields.append(QgsField('OutPass', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 6)
        trfields.append(QgsField('InOn', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 7)
        trfields.append(QgsField('InOff', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 8)
        trfields.append(QgsField('InPass', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 9)
        trfields.append(QgsField('TotalPass', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 10)
        
        stfields = QgsFields()
        stfields.append(QgsField('LineName', QVariant.String, 'String', 254), QgsFields.OriginProvider, 1)
        stfields.append(QgsField('Station', QVariant.String, 'String', 254), QgsFields.OriginProvider, 2)
        stfields.append(QgsField('OutOn', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 4)
        stfields.append(QgsField('OutOff', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 5)
        stfields.append(QgsField('InOn', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 7)
        stfields.append(QgsField('InOff', QVariant.Int, 'Integer'), QgsFields.OriginProvider, 8)
        
        # Old same name layer is vanished.
        QgsVectorFileWriter.deleteShapeFile(outscshppath)
        QgsVectorFileWriter.deleteShapeFile(outstshppath)
        
        # Create writer for new layer.
        self.scwriter = QgsVectorFileWriter.create( \
            outscshppath, trfields, QgsWkbTypes.LineString, \
            self.crs, self.transform_context, self.save_options)
        
        if self.scwriter.hasError() != QgsVectorFileWriter.NoError:
            raise FailCreateWriter('Error when creating shapefile: ' + outscshppath + ' with ' + self.writer.errorMessage())
        
        self.stwriter = QgsVectorFileWriter.create( \
            outstshppath, stfields, QgsWkbTypes.Point, \
            self.crs, self.transform_context, self.save_options)
        
        if self.stwriter.hasError() != QgsVectorFileWriter.NoError:
            raise FailCreateWriter('Error when creating shapefile: ' + outstshppath + ' with ' + self.writer.errorMessage())
        
        rg = ws.Range("A6")
        
        while rg.Borders(9).LineStyle == -4142:
            # create railroad line object and read data from excel book and layer of station points.
            railline = RailRoadsLine()
            rg = railline.createstations(rg, self.sslayer, self.splayer)
            
            # Search the paths between stations.
            railline.searchroute(trfields, self.scwriter, stfields, self.stwriter)

