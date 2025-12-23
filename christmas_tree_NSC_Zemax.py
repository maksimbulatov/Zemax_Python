#from enum import Enum
import clr, os, winreg
from itertools import islice

# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running a Python version compatible with PythonNET. Check the article "ZOS-API using Python.NET" or
#    "Getting started with Python" in our knowledge base for more details.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

# determine the Zemax working directory
aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
winreg.CloseKey(aKey)

# add the NetHelper DLL for locating the OpticStudio install folder
clr.AddReference(NetHelper)
import ZOSAPI_NetHelper

pathToInstall = ''
# uncomment the following line to use a specific instance of the ZOS-API assemblies
#pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

# connect to OpticStudio
success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);

zemaxDir = ''
if success:
    zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
    print('Found OpticStudio at:   %s' + zemaxDir);
else:
    raise Exception('Cannot find OpticStudio')

# load the ZOS-API assemblies
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
import ZOSAPI

TheConnection = ZOSAPI.ZOSAPI_Connection()
if TheConnection is None:
    raise Exception("Unable to intialize NET connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

def reshape(data, x, y, transpose = False):
    """Converts a System.Double[,] to a 2D list for plotting or post processing
    
    Parameters
    ----------
    data      : System.Double[,] data directly from ZOS-API 
    x         : x width of new 2D list [use var.GetLength(0) for dimension]
    y         : y width of new 2D list [use var.GetLength(1) for dimension]
    transpose : transposes data; needed for some multi-dimensional line series data
    
    Returns
    -------
    res       : 2D list; can be directly used with Matplotlib or converted to
                a numpy array using numpy.asarray(res)
    """
    if type(data) is not list:
        data = list(data)
    var_lst = [y] * x;
    it = iter(data)
    res = [list(islice(it, i)) for i in var_lst]
    if transpose:
        return self.transpose(res);
    return res
    
def transpose(data):
    """Transposes a 2D list (Python3.x or greater).  
    
    Useful for converting mutli-dimensional line series (i.e. FFT PSF)
    
    Parameters
    ----------
    data      : Python native list (if using System.Data[,] object reshape first)    
    
    Returns
    -------
    res       : transposed 2D list
    """
    if type(data) is not list:
        data = list(data)
    return list(map(list, zip(*data)))

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

# Insert Code Here


TheNCE = TheSystem.NCE

# Set system wavelength (visible light, 0.55 um)
TheSystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = 0.55

# Add the light source (Object 1: Source Point at bottom right corner of the trianlge)
source_row = TheNCE.InsertNewObjectAt(1)
source_1 = TheNCE.GetObjectAt(1)
oType_1 = source_1.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourcePoint)
source_1.ChangeType(oType_1)
source_1.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par1).IntegerValue = 1  # rays for visualization
source_1.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par2).IntegerValue = 0  # Number of visualization rays
source_1.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0  # Watt

# Add mirrors (flat planes with mirror material) to create a triangle "tree"

# Mirror 1
mirror1_row = TheNCE.InsertNewObjectAt(2)
mirror1 = TheNCE.GetObjectAt(2)
oType_2 = mirror1.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror1.ChangeType(oType_2)
mirror1.Material = 'MIRROR'
mirror1.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0  # Size
mirror1.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 5.0
mirror1.TiltAboutX = 30.0  # Tilt

# Mirror 2 
mirror2_row = TheNCE.InsertNewObjectAt(3)
mirror2 = TheNCE.GetObjectAt(3)
oType_3 = mirror2.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror2.ChangeType(oType_3)
mirror2.Material = 'MIRROR'
mirror2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = (5**2 - 2.5**2)**0.5
mirror2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 2.5
mirror2.TiltAboutX = 90.0 # Tilt

# Mirror 3
mirror3_row = TheNCE.InsertNewObjectAt(4)
mirror3 = TheNCE.GetObjectAt(4)
oType_4 = mirror3.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror3.ChangeType(oType_4)
mirror3.Material = 'ABSORB'
mirror3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror3.TiltAboutX = -30.0

# Add the second light source
source_row = TheNCE.InsertNewObjectAt(5)
source_2 = TheNCE.GetObjectAt(5)
oType_5 = source_2.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourcePoint)
source_2.ChangeType(oType_5)
source_2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = -2.5  # Top of tree
source_2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5
source_2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par1).IntegerValue = 1  # Rays for visualization
source_2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par2).IntegerValue = 0  # Number of analysis rays
source_2.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0  # Watt

# Mirror 4 
mirror4_row = TheNCE.InsertNewObjectAt(6)
mirror4 = TheNCE.GetObjectAt(6)
oType_5 = mirror4.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror4.ChangeType(oType_5)
mirror4.Material = 'MIRROR'
mirror4.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror4.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5
mirror4.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 7.5
mirror4.TiltAboutX = 30.0

# Mirror 5
mirror5_row = TheNCE.InsertNewObjectAt(7)
mirror5 = TheNCE.GetObjectAt(7)
oType_6 = mirror5.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror5.ChangeType(oType_6)
mirror5.Material = 'MIRROR'
mirror5.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror5.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = 0.0
mirror5.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 2.5
mirror5.TiltAboutX = 90.0

# Mirror 6
mirror6_row = TheNCE.InsertNewObjectAt(8)
mirror6 = TheNCE.GetObjectAt(8)
oType_7 = mirror6.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror6.ChangeType(oType_7)
mirror6.Material = 'ABSORB'
mirror6.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror6.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5
mirror6.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = -2.5
mirror6.TiltAboutX = -30.0

# Add the third light source
source_row = TheNCE.InsertNewObjectAt(9)
source_3 = TheNCE.GetObjectAt(9)
oType_8 = source_3.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourcePoint)
source_3.ChangeType(oType_8)
source_3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = -5.0  # Top of tree
source_3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5-(15**2 - 7.5**2)**0.5
source_3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par1).IntegerValue = 1  # Rays for visualization
source_3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par2).IntegerValue = 0  # number of analysis rays
source_3.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0  # Watt

# Mirror 7 
mirror7_row = TheNCE.InsertNewObjectAt(10)
mirror7 = TheNCE.GetObjectAt(10)
oType_9 = mirror7.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror7.ChangeType(oType_9)
mirror7.Material = 'MIRROR'
mirror7.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror7.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5-(15**2 - 7.5**2)**0.5
mirror7.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 10.0
mirror7.TiltAboutX = 30.0

# Mirror 8
mirror8_row = TheNCE.InsertNewObjectAt(11)
mirror8 = TheNCE.GetObjectAt(11)
oType_10 = mirror8.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror8.ChangeType(oType_10)
mirror8.Material = 'MIRROR'
mirror8.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror8.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5
mirror8.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = 2.5
mirror8.TiltAboutX = 90.0

# Mirror 9
mirror9_row = TheNCE.InsertNewObjectAt(12)
mirror9 = TheNCE.GetObjectAt(12)
oType_11 = mirror9.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.StandardSurface)
mirror9.ChangeType(oType_11)
mirror9.Material = 'ABSORB'
mirror9.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.Par3).DoubleValue = 1.0
mirror9.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.YPosition).DoubleValue = -(10**2 - 5**2)**0.5-(15**2 - 7.5**2)**0.5
mirror9.GetObjectCell(ZOSAPI.Editors.NCE.ObjectColumn.ZPosition).DoubleValue = -5.0
mirror9.TiltAboutX = -30.0