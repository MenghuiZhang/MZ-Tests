# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from System.Windows.Input import Key
import os
from Autodesk.Revit.UI import TaskDialog,Selection
import Autodesk.Revit.DB as DB


__title__ = "Abstand"
__doc__ = """

Formteile und Zubehöre auf eingegebenen Abstand.

für Luftkanal- und Rohrteile.

[2022.08.01]
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

class AlleFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008010' or element.Category.Id.ToString() == '-2008016'\
            or element.Category.Id.ToString() == '-2008049' or element.Category.Id.ToString() == '-2008055':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class LuftFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008010' or element.Category.Id.ToString() == '-2008016':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class RohrFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008049' or element.Category.Id.ToString() == '-2008055':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

class GUI(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))
    
    def escape(self, sender, args):
        pass
        # if args.Key == self.Key.Escape:
        #     self.Abstand.IsEnabled = True
        #     self.a = False
        #     if self.verschieben.t != None:
        #         print(self.verschieben.t.GetStatus().ToString())
        #         if self.verschieben.t.GetStatus().ToString() == 'Started':
        #             self.verschieben.t.Commit()

    def aktualisieren(self, sender, args):
        self.Close()
        # self.a = True
        # self.verschiebenEvent.Raise()
    
    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        
        else:
            args.Handled = True

gui = GUI()
gui.ShowDialog()

# tg = DB.TransactionGroup(doc,'Abstand')
# tg.Start()

while (True):   
    try:
        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,AlleFilter(),'Wählt das fixierte Luftkanalformteil/-zubehör/Rohrformteil/-zubehör aus')
        el0 = doc.GetElement(el0_ref)

        if gui.Abstand.Text == None or gui.Abstand.Text == '':
            gui.Abstand.Text = '10'
        try:
            abstand = float(gui.Abstand.Text) / 304.8
        except:
            TaskDialog.Show('Fehler','ungültigen Abstand!')
            break

        if el0.Category.Id.ToString() in ['-2008010','-2008016']:
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,LuftFilter(),'Wählt das zu verschiebende Luftkanalformteil/-zubehör aus')
        else:
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,RohrFilter(),'Wählt das zu verschiebende Rohrformteil/-zubehör aus')

        el1 = doc.GetElement(el1_ref)

        distance = 100000000
        co0 = None
        co1 = None
        conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        conns1 = list(el1.MEPModel.ConnectorManager.Connectors)

        for con0 in conns0:
            for con1 in conns1:
                dis = con0.Origin.DistanceTo(con1.Origin)
                if dis < distance:
                    distance = dis
                    co0 = con0
                    co1 = con1
                    
        l0 = DB.Line.CreateBound(co0.Origin,co1.Origin)
        ln = l0.Direction.Normalize()
        neu = co0.Origin + ln * abstand
        t = DB.Transaction(doc,'Abstand')
        t.Start()
        pinned = el1.Pinned
        el1.Pinned = False
        doc.Regenerate()
        try:
            el1.Location.Move(neu - co1.Origin)
        except Exception as e:
            print(e)
        el1.Pinned = pinned
        doc.Regenerate()
        t.Commit()
        t.Dispose()
    except:
        break

# tg.Commit()
