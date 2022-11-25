# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from System import Guid 
from System.Collections.Generic import List
from IGF_log import getlog,getloglocal
import clr
from rpw import revit,DB,UI
from pyrevit import script, forms
# from eventhandler import Legend_Normal,ExternalEvent,Legend_Duct,Legend_Color,Legend_Keynote
from System.Collections.Generic import List
from System.Windows.Input import ModifierKeys,Keyboard,Key

__title__ = "ExtensibleStorage"
__doc__ = """

Legenden für ausgewählte Ansicht erstellen
[2022.05.31]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

from System import Guid 

uidoc = revit.uidoc
doc = revit.doc
logger = script.get_logger()

## Set Extensibel Storyge
t = DB.Transaction(doc, "Test")

t.Start()
schemaBuilder = DB.ExtensibleStorage.SchemaBuilder(Guid( "720080CB-DA99-40DC-9415-E53F280AA1F1" ))
schemaBuilder.SetSchemaName('TestSchema1')
schemaBuilder.SetReadAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
schemaBuilder.SetWriteAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
#fieldBuilder = schemaBuilder.AddSimpleField( "WireSpliceLocation", clr.GetClrType(DB.XYZ) )
fieldBuilder = schemaBuilder.AddSimpleField( "WireSpliceLocation", clr.GetClrType(str) )
# fieldBuilder.SetUnitType(DB.UnitType.UT_Length )
fieldBuilder.SetDocumentation( "A stored "
    + "location value representing a wiring "
    + "splice in a wall." )
schema = schemaBuilder.Finish()

entity = DB.ExtensibleStorage.Entity( schema )
fieldSpliceLocation = schema.GetField("WireSpliceLocation" )
#entity.Set[DB.XYZ]( fieldSpliceLocation, el.Location.Point,DB.DisplayUnitType.DUT_METERS )
entity.Set[str]( fieldSpliceLocation, 'Test' )
el.Symbol.SetEntity( entity )



t.Commit()

## Read entity
schame = DB.ExtensibleStorage.Entity(Guid( "720080CB-DA99-40DC-9415-E53F280AA1F1" )).Schema
Data = el.Symbol.GetEntity(schema).Get[str](schema.GetField( "WireSpliceLocation" ))

# 762a2314-1d1c-4087-a58f-bab902f57be5