import os

def setenvirons():
    # GDAL
    os.environ['GDAL_DATA'] = 'C:\\Program Files\\QGIS 3.10\\share\\gdal'
    os.environ['GDAL_DRIVER_PATH'] = 'C:\\Program Files\\QGIS 3.10\\bin\\gdalplugins'
    os.environ['GDAL_FILENAME_IS_UTF8'] = 'YES'
    
    # OSGeo4W
    os.environ['O4W_QT_BINARIES'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\bin'
    os.environ['O4W_QT_HEADERS'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\include'
    os.environ['O4W_QT_LIBRARIES'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\lib'
    os.environ['O4W_QT_PLUGINS'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\plugins'
    os.environ['O4W_QT_PREFIX'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5'
    os.environ['O4W_QT_TRANSLATIONS'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\translations'
    os.environ['OSGEO4W_ROOT'] = 'C:\\Program Files\\QGIS 3.10'
    
    os.environ['PROJ_LIB'] = 'C:\\Program Files\\QGIS 3.10\\share\\proj'
    os.environ['QGIS_PREFIX_PATH'] = 'C:\\Program Files\\QGIS 3.10\\apps\\qgis-ltr'
    os.environ['QT_PLUGIN_PATH'] = 'C:\\Program Files\\QGIS 3.10\\apps\\qgis-ltr\\qtplugins;C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\plugins'
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = 'C:\\Program Files\\QGIS 3.10\\apps\\Qt5\\plugins'
    