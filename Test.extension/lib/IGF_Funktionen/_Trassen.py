# coding: utf8
import System
from Autodesk.Revit.UI import TaskDialog
from System.IO import Path, File
from rpw import revit,DB,UI
from System.Collections.Generic import List
from color_conversion import rgb2hsl,hsl2rgb
from System import Byte
import clr
clr.AddReference('System.Drawing')
from System.Drawing import Color
import math

def ErstellenUebergang(doc,conn_0,conn_1):
    try:
        fi = doc.Create.NewTransitionFitting(conn_0, conn_1)
        doc.Regenerate()
        return fi
    except:
        try:
            fi = doc.Create.NewTransitionFitting(conn_1, conn_0)
            doc.Regenerate()
            return fi
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())

def ErstellenBogen(doc,conn_0,conn_1):
    doc.Regenerate()
    try:
        fi = doc.Create.NewElbowFitting(conn_0, conn_1)
        doc.Regenerate()
        return fi
    except:
        try:
            fi = doc.Create.NewElbowFitting(conn_1, conn_0)
            doc.Regenerate()
            return fi
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())
        
def IsHorizontalBogen(Bogen):
    if Bogen.MEPModel.PartType.ToString() == 'Elbow':
        conns = Bogen.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.CoordinateSystem.BasisZ.IsAlmostEqualTo(DB.XYZ(0,0,1)) or conn.CoordinateSystem.BasisZ.IsAlmostEqualTo(DB.XYZ(0,0,-1)):
                return True
        return False
    return None

def Volumenstrom_berechnen(tempo = 0,b = None,h = None,d = None):
    """
    tempo: Geschwindigkeit in m/s
    b: Breite in mm
    h: Höhe in mm
    d: Durchmesser in mm
    """
    if d:
        return round(3600 * float(tempo) * math.pi * d * d / 4000000, 2)
    else:
        return round(3600 * float(tempo) / 1000000 * b * h, 2)
    
def Tempo_berechnen(vol = 0,b = None,h = None,d = None):
    """
    vol: Volumenstrom in m³/s
    b: Breite in mm
    h: Höhe in mm
    d: Durchmesser in mm
    """
    if d:
        return round(float(vol) / (math.pi * d * d / 4000000) / 3600,2)
    else:
        return round(float(vol) / h / b * 1000000 / 3600,2)

def Dimension_berechnen(vol = None, tempo = None, b = None, h = None, d = None):
    """
    vol: Volumenstrom in m³/s
    tempo: Geschwindigkeit in m/s
    b: Breite in mm
    h: Höhe in mm
    d: Durchmesser in mm
    """
    if vol and tempo:
        if b:
            h = math.ceil(float(vol) / tempo / b / 3600 * 1000000.0)
            return h
        elif h:
            b = math.ceil(float(vol) / tempo / h / 3600 * 1000000.0 )
            return b
        else:
            d = math.ceil(math.sqrt(float(vol) / tempo / math.pi) * 100 / 3)
            return d

