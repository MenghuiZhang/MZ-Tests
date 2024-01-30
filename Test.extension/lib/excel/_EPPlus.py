import sys
import clr
import os
sys.path.Add(os.path.dirname(__file__)+'\\net462')
clr.AddReference('EPPlus')
import OfficeOpenXml
from OfficeOpenXml import ExcelPackage,LicenseContext
from System.IO import FileStream,FileMode,FileAccess,FileShare