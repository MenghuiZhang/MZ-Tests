# coding: utf8
# from IGF_log import getlog,getloglocal
# from pyrevit import forms
# from eventhandler import TransitionFitting,ExternalEvent,TransitionFitting1
# import os
# from System.Windows.Input import Key
from rpw import revit,UI,DB
import System
from System.Collections.Generic import List
import clr
import math

__title__ = "anschlussset"
__doc__ = """

[2022.08.01]
Version: 1.1
"""
__authors__ = "Menghui Zhang"



uidoc = revit.uidoc
doc = revit.doc


class Anschlussset:
    def __init__(self,elem):
        self.elem = elem
        self.conn_vl_0 = None
        self.conn_vl_1 = None
        self.conn_rl_0 = None
        self.conn_rl_1 = None

        self._conn_vl_0 = None
        self._conn_vl_1 = None
        self._conn_rl_0 = None
        self._conn_rl_1 = None

        self.__conn_vl_0 = None
        self.__conn_vl_1 = None
        self.__conn_rl_0 = None
        self.__conn_rl_1 = None

        self.Point_O_0 = None
        self.Point_O_1 = None
        self.Point_NO_0 = None
        self.Point_NO_1 = None
        self.VL = None
        self.RL = None
        self.level = doc.GetElement(self.elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId())
        self.get_anschlussset()
        self.get_Points_O()


    def get_anschlussset(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.PipeSystemType.ToString() == 'SupplyHydronic':
                if conn.Direction.ToString() == 'In':
                    self.conn_vl_0 = conn
                    for ref in conn.AllRefs:
                        owner = ref.Owner
                        if owner.Category.Name in ['HLS-Bauteile','Rohre','Rohrformteile']:
                            self.__conn_vl_0 = ref
                else:
                    self.conn_vl_1 = conn
                    for ref in conn.AllRefs:
                        owner = ref.Owner
                        if owner.Category.Name in ['HLS-Bauteile','Rohre','Rohrformteile']:
                            self.__conn_vl_1 = ref
            else:
                if conn.Direction.ToString() == 'In':
                    self.conn_rl_0 = conn
                    for ref in conn.AllRefs:
                        owner = ref.Owner
                        if owner.Category.Name in ['HLS-Bauteile','Rohre','Rohrformteile']:
                            self.__conn_rl_0 = ref

                else:
                    self.conn_rl_1 = conn
                    for ref in conn.AllRefs:
                        owner = ref.Owner
                        if owner.Category.Name in ['HLS-Bauteile','Rohre','Rohrformteile']:
                            self.__conn_rl_1 = ref
    
    def Point(self,conn0,conn1):
        if conn0 and conn1:
            line0 = DB.Line.CreateUnbound(conn0.CoordinateSystem.Origin,conn0.CoordinateSystem.BasisZ.Negate())
            line1 = DB.Line.CreateUnbound(conn1.CoordinateSystem.Origin,conn1.CoordinateSystem.BasisZ.Negate())
            
            array = clr.Reference[DB.IntersectionResultArray]()
            line0.Intersect(line1,array)

            line0.Dispose()
            line1.Dispose()
            return array.Item[0].XYZPoint
    
    def get_Points_O(self):
        self.Point_O_0 = self.Point(self.conn_vl_0, self.conn_vl_1)
        self.Point_O_1 = self.Point(self.conn_rl_0, self.conn_rl_1)
    
    # 7495661
    # 7502555
    def create(self):
        self.VL = doc.Create.NewFamilyInstance(self.Point_O_0, doc.GetElement(DB.ElementId(7495661)), self.level, DB.Structure.StructuralType.NonStructural)
        self.RL = doc.Create.NewFamilyInstance(self.Point_O_1, doc.GetElement(DB.ElementId(7502555)), self.level, DB.Structure.StructuralType.NonStructural)
        doc.Regenerate()
    
    def get_anschluss(self,elem):
        conns = elem.MEPModel.ConnectorManager.Connectors
        conn_0 = None
        conn_1 = None
        for conn in conns:
            if conn.Direction.ToString() == 'In':
                conn_0 = conn

            else:
                conn_1 = conn
        return conn_0,conn_1
    
    def get_Points_NO(self):
        conns = list(self.VL.MEPModel.ConnectorManager.Connectors)
        self.Point_NO_0 = self.Point(conns[0],conns[1])
        conns = list(self.RL.MEPModel.ConnectorManager.Connectors)
        self.Point_NO_1 = self.Point(conns[0],conns[1])

    def get_points_VL_RL(self):
        self._conn_vl_0,self._conn_vl_1 = self.get_anschluss(self.VL)
        self._conn_rl_0,self._conn_rl_1 = self.get_anschluss(self.RL)

    def verschieben(self):
        if not self.Point_O_0.IsAlmostEqualTo(self.Point_NO_0):
            self.VL.Location.Move(self.Point_O_0-self.Point_NO_0)
        if not self.Point_O_1.IsAlmostEqualTo(self.Point_NO_1):
            self.RL.Location.Move(self.Point_O_1-self.Point_NO_1)
    
    def umdrehen(self):
        doc.Regenerate()
        V_0 = self.conn_vl_0.CoordinateSystem.BasisZ
        _V_0 = self._conn_vl_0.CoordinateSystem.BasisZ
        if V_0.IsAlmostEqualTo(_V_0):
            Fa_0 = V_0
            An_0 = 0
        elif V_0.IsAlmostEqualTo(_V_0.Negate()):
            Fa_0 = self.conn_vl_1.CoordinateSystem.BasisZ
            An_0 = V_0.AngleTo(_V_0)
        else:
            Fa_0 = V_0.CrossProduct(_V_0)
            An_0 = V_0.AngleTo(_V_0)
        
        if An_0 > 0:
            DB.ElementTransformUtils.RotateElement(doc,self.VL.Id,DB.Line.CreateUnbound(self.Point_O_0,Fa_0),An_0)       
        doc.Regenerate()

        V_1 = self.conn_vl_1.CoordinateSystem.BasisZ
        _V_1 = self._conn_vl_1.CoordinateSystem.BasisZ
        if V_1.IsAlmostEqualTo(_V_1):
            Fa_1 = V_1
            An_1 = 0
        elif V_1.IsAlmostEqualTo(_V_1.Negate()):
            Fa_1 = self.conn_vl_0.CoordinateSystem.BasisZ
            An_1 = V_1.AngleTo(_V_1)
        else:
            Fa_1 = V_1.CrossProduct(_V_1)
            An_1 = V_1.AngleTo(_V_1)

        if An_1 > 0:
            DB.ElementTransformUtils.RotateElement(doc,self.VL.Id,DB.Line.CreateUnbound(self.Point_O_0,Fa_1),An_1)
        doc.Regenerate()

        R_0 = self.conn_rl_0.CoordinateSystem.BasisZ
        _R_0 = self._conn_rl_0.CoordinateSystem.BasisZ
        if R_0.IsAlmostEqualTo(_R_0):
            Fa_2 = R_0
            An_2 = 0
        elif R_0.IsAlmostEqualTo(_R_0.Negate()):
            Fa_2 = self.conn_rl_1.CoordinateSystem.BasisZ
            An_2 = R_0.AngleTo(_R_0)
        else:
            Fa_2 = R_0.CrossProduct(_R_0)
            An_2 = R_0.AngleTo(_R_0)
        
        if An_2 > 0:
            DB.ElementTransformUtils.RotateElement(doc,self.RL.Id,DB.Line.CreateUnbound(self.Point_O_1,Fa_2),An_2)
        doc.Regenerate()

        R_1 = self.conn_rl_1.CoordinateSystem.BasisZ
        _R_1 = self._conn_rl_1.CoordinateSystem.BasisZ
        if R_1.IsAlmostEqualTo(_R_1):
            Fa_3 = R_1
            An_3 = 0
        elif R_1.IsAlmostEqualTo(_R_1.Negate()):
            Fa_3 = self.conn_rl_0.CoordinateSystem.BasisZ
            An_3 = R_1.AngleTo(_R_1)
        else:
            Fa_3 = R_1.CrossProduct(_R_1)
            An_3 = R_1.AngleTo(_R_1)
        if An_3 > 0:
            DB.ElementTransformUtils.RotateElement(doc,self.RL.Id,DB.Line.CreateUnbound(self.Point_O_1,Fa_3),An_3)

    def systemtrennen(self):
        conn = DB.ConnectorSet()
        try:
            conn.Insert(self.__conn_vl_1)
            self.__conn_vl_1.MEPSystem.Remove(conn)
        except Exception as e:print(e)
        try:
            conn.Clear()
            conn.Insert(self.__conn_rl_0)
            self.__conn_rl_0.MEPSystem.Remove(conn)
        except Exception as e:print(e)
        conn.Dispose()
        doc.Regenerate()

    def systemverbinden(self):
        conn = DB.ConnectorSet()
        try:
            conn.Insert(self.__conn_vl_1)
            self.__conn_vl_0.MEPSystem.Add(conn)
        except Exception as e:print(e)
        try:
            conn.Clear()
            conn.Insert(self.__conn_rl_0)
            self.__conn_rl_1.MEPSystem.Add(conn)
        except Exception as e:print(e)
        conn.Dispose()
        doc.Regenerate()
    
    def verbinden_einzeln(self,conn0,conn1):
        try:
            conn1.Owner.Location.Move(conn0.Origin-conn1.Origin)
            doc.Regenerate()
            conn0.ConnectTo(conn1)
            doc.Regenerate()
        except:
            pass
    
    def verbinden(self):
        self.verbinden_einzeln(self.__conn_vl_1, self._conn_vl_1)
        self.verbinden_einzeln(self._conn_vl_0, self.__conn_vl_0)
        self.verbinden_einzeln(self.__conn_rl_0, self._conn_rl_0)
        self.verbinden_einzeln(self._conn_rl_1, self.__conn_rl_1)

    
    def bearbeiten(self):
        self.create()
        doc.Regenerate()
        self.get_points_VL_RL()
        self.get_Points_NO()
        doc.Regenerate()
        self.verschieben()
        doc.Regenerate()
        self.umdrehen()
        doc.Regenerate()
        doc.Delete(self.elem.Id)
        self.systemtrennen()
        self.systemverbinden()
        self.verbinden()
        
        


t = DB.Transaction(doc,'Ventile austauschen')
t.Start()
for elid in uidoc.Selection.GetElementIds():
    Anschlussset(doc.GetElement(elid)).bearbeiten() 
t.Commit()