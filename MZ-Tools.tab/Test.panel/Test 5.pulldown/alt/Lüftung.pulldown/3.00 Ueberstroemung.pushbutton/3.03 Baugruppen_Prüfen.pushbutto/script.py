# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import rpw
import time
import clr
from pyIGF_lib import get_value
from System.Collections.ObjectModel import *
from pyrevit.forms import WPFWindow

start = time.time()

__title__ = "3.01 Überströmung_aktualisieren_test"
__doc__ = """Raumnummer und Volumen von Überstrom in Baugruppen_Überströmung und in MEP-Räume schreiben
Überströmung:
System: 31_Überstromluft,

"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
auslass_config = script.get_config()
from pyIGF_logInfo import getlog
getlog(__title__)

doc = rpw.revit.doc

# Baugruppen
Baugruppen_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.AssemblyInstance)) \
    .WhereElementIsNotElementType()

Auslass_Collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal) \
    .WhereElementIsNotElementType()

Baugruppen = Baugruppen_collector.ToElementIds()
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} Baugruppen ausgewählt".format(len(Baugruppen)))

class Bauteile(object):
    def __init__(self):
        self.Element = ''

    @property
    def Element(self):
        return self._Element
    @Element.setter
    def Element(self, value):
        self._Element = value

Liste_Bauteile = []
Liste_zu = []
Liste_Aus = []
for el in Auslass_Collector:
    Name = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    if not Name in Liste_Bauteile:
        Liste_Bauteile.append(Name)
Liste_Bauteile.sort()

Liste_IN_Familie = ObservableCollection[Bauteile]()
Liste_AUS_Familie = ObservableCollection[Bauteile]()

def ObservableCollection2Liste(Coll):
    out_list = []

    if Coll.Count > 0:
        for el in Coll:
            out_list.append(el.Element)
    return out_list

def Liste2Class(input):
    out_class = Bauteile()
    out_class.Element = str(input)
    return out_class

def write_config():
    try:
        auslass_config.liste_In = ObservableCollection2Liste(Liste_IN_Familie)
    except:
        pass
    try:
        auslass_config.liste_Aus = ObservableCollection2Liste(Liste_AUS_Familie)
    except:
        pass
    script.save_config()

def read_config():
    try:
        if not len(auslass_config.liste_In) == 0:
            for el in auslass_config.liste_In:
                Liste_IN_Familie.Add(Liste2Class(el))
    except Exception as e:
        pass
    try:
        if not len(auslass_config.liste_Aus) == 0:
            for el1 in auslass_config.liste_Aus:
                Liste_AUS_Familie.Add(Liste2Class(el1))
    except Exception as e:
        pass

read_config()


class FamilienBearbeiten(WPFWindow):
    def __init__(self, xaml_file_name,liste_in,liste_aus):
        self.ListeIN = liste_in
        self.ListeAUS = liste_aus
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.dataGrid.ItemsSource = liste_in
            self.dataGrid.Columns[0].ItemsSource = Liste_Bauteile
        except:
            pass

    def cancel(self,sender,args):
        self.Close()
        self.ListeIN.Clear()
        self.ListeAUS.Clear()
        read_config()


    def OK(self,sender,args):
        self.Close()

    def IN(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeIN
        self.dataGrid.Columns[0].ItemsSource = Liste_Bauteile


    def AUS(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeAUS
        self.dataGrid.Columns[0].ItemsSource = Liste_Bauteile

    def Add(self,sender,args):
        temp_class = Bauteile()
        self.dataGrid.ItemsSource.Add(temp_class)
        self.dataGrid.Items.Refresh()

if forms.alert('Überströmung_Familien ergänzen?', ok=False, yes=True, no=True):
    FamilienDialog = FamilienBearbeiten('window.xaml', Liste_IN_Familie, Liste_AUS_Familie)
    FamilienDialog.ShowDialog()
    write_config()

Liste_zu = ObservableCollection2Liste(Liste_IN_Familie)
Liste_Aus = ObservableCollection2Liste(Liste_AUS_Familie)

phase = list(doc.Phases)[-1]

# Baugruppen
def Ueberstrom_BG(Familie,Familie_Id):
    Kanal_Collector = []
    Auslass_Collector = []
    BG_Liste = []
    Title = '{value}/{max_value} Baugruppen'
    with forms.ProgressBar(title=Title, cancellable=True, step=2) as pb:
        n = 0

        for item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            ElementIds = item.GetMemberIds()
            id_BG = item.Id
            Volumen = get_value(item.LookupParameter('IGF_RLT_Überströmung'))
            BG_Liste.append([id_BG,Volumen])
            for id in ElementIds:
                Element = doc.GetElement(id)
                Category_Name = Element.Category.Name
                if Category_Name == 'Luftkanäle':
                    Kanal_Collector.append(Element)
                elif Category_Name == 'Luftdurchlässe':
                    Auslass_Collector.append(Element)

    logger.info("{} Kanäle ausgewählt".format(len(Kanal_Collector)))
    logger.info("{} Luftauslässe ausgewählt".format(len(Auslass_Collector)))

    return Kanal_Collector,Auslass_Collector,BG_Liste

# Raumdaten ermitteln
def Raum_ermitteln(Familie,BG_List):
    Raum = []
    Title = '{value}/{max_value} Raumnummer Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=10) as pb:
        n = 0

        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie))

            Raumnummer = None
            if Item.Space[phase]:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                if System_Typ == '31_Überstromluft':
                    Raumnummer = Item.Space[phase].Number
                    System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                    Familie_Typ = Item.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                    BG_ID = Item.AssemblyInstanceId
                    Vol = 0
                    for Bg in BG_List:
                        if BG_ID == Bg[0]:
                            Vol = Bg[1]
                    Raum.append([System_Name,Familie_Typ,Raumnummer,Vol,BG_ID])

                    if Familie_Typ in Liste_zu:
                        logger.info('Raum: {}, Überstrom_IN {}'.format(Raumnummer,Vol))
                    elif Familie_Typ in Liste_Aus:
                        logger.info('Raum: {}, Überstrom_AUS {}'.format(Raumnummer,Vol))
                else:
                    logger.error(System_Typ)

    return Raum

def Liste(input_list_Raum):
    out_liste = []
    out_liste1 = []
    sys_liste = []
    for item in input_list_Raum:
        sys_liste.append(item[0])
    sys_liste = set(sys_liste)
    sys_liste = list(sys_liste)
    sys_liste.sort()
    for item in sys_liste:
        out_liste.append([item,'','',0])
        out_liste1.append([item,'','',0,''])
    for item in input_list_Raum:
        for sys in out_liste:
            if item[0] == sys[0]:
                if item[1] in Liste_Aus:
                    sys[1] = item[2]
                    if item[3] > 0:
                        sys[3] = item[3]
                elif item[1] in Liste_zu:
                    sys[2] = item[2]
                    if item[3] > 0:
                        sys[3] = item[3]
        for sys in out_liste1:
            if item[0] == sys[0]:
                if item[1] in Liste_Aus:
                    sys[1] = item[2]
                    sys[4] = item[4]
                    if item[3] > 0:
                        sys[3] = item[3]
                elif item[1] in Liste_zu:
                    sys[2] = item[2]
                    sys[4] = item[4]
                    if item[3] > 0:
                        sys[3] = item[3]

    output.print_table(
        table_data=out_liste,
        title="System",
        columns=['Systemname', 'Eingangsraum', 'Ausgangsraum','Volumenstrom']
    )
    return out_liste,out_liste1

# def Pruefen(Sys_Liste):
#     duct = DB.FilteredElementCollector(doc)\
#            .OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystem))\
#            .WhereElementIsNotElementType()
#     System = []
#     for sys in duct:
#         type = sys.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
#         Name = sys.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsValueString()
#         if type == '31_Überstromluft':
#             System.append(Name)
#     if i == len(Sys_Liste):
#         pass
#     else:
#         logger.error('Baugruppe: {}, Überstrom Systeme: {}'.format(Sys_Liste,i))
#         system = []
#         for syst in Sys_Liste:
#             system.append(syst[0])
#         c = set(System) - set(system)
#         logger.error('nur in Systeme: {}'.format(c))

def Datenschreiben_Kanal(input_list,Familie):
    Title = '{value}/{max_value} Daten in Kanal schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Ein-/Ausgangsraum und Überstrom in Kanal schreiben")
        t.Start()

        for Item in Familie:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie))
            id_BG = Item.Id
            for item in input_list:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                if System_Name == item[0] and System_Typ == '31_Überstromluft':
                    Item.LookupParameter('IGF_RLT_Überströmung_Eingang').Set(str(item[1]))
                    Item.LookupParameter('IGF_RLT_Überströmung_Ausgang').Set(str(item[2]))
                    Item.LookupParameter('IGF_RLT_Überströmung').SetValueString(str(item[3]))

        t.Commit()

def Datenschreiben_Auslass(input_list,Familie):
    Title = '{value}/{max_value} Daten in Auslässe schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Überstrom in Auslass schreiben")
        t.Start()
        Error = ''

        for Item in Familie:

            if pb.cancelled:
                t.RollBack()
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie))

            for item in input_list:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                if System_Name == item[0] and System_Typ == '31_Überstromluft':
                    Item.LookupParameter('IGF_RLT_Überströmung').SetValueString(str(item[3]))
                    try:
                        raumnummer = Item.Space[phase].Number
                        if raumnummer == item[2]:
                            Item.LookupParameter('IGF_RLT_ÜberströmungAusRaum').Set(str(item[1]))
                        elif raumnummer == item[1]:
                            Item.LookupParameter('IGF_RLT_ÜberströmungInRaum').Set(str(item[2]))
                    except Exception as e:
                        Error = Error + '\n'+ System_Name+': Kein Ein/Ausgangsraum gefunden.'
                        logger.error(System_Name+': Kein Ein-/Ausgangsraum gefunden.')
        if Error:
            TaskDialog.Show('Error',Error)
            t.RollBack()
            script.exit()

        t.Commit()

def Datenschreiben_BG(input_list,Familie):
    Title = '{value}/{max_value} Daten in Baugruppen schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Ein-/Ausgangsraum und Überstrom in Baugruppen schreiben")
        t.Start()

        for item in input_list:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(input_list))

            for Item in Familie:
                id_BG = Item.Id
                if id_BG == item[4]:
                    Item.LookupParameter('IGF_RLT_Überströmung_Eingang').Set(str(item[1]))
                    Item.LookupParameter('IGF_RLT_Überströmung_Ausgang').Set(str(item[2]))

        t.Commit()

Kanal_Collector,Auslass_Collector,BG_Liste = Ueberstrom_BG(Baugruppen_collector,Baugruppen)
Raum_Liste = Raum_ermitteln(Auslass_Collector,BG_Liste)
Sys_Liste,Sys_Liste_MitBGId = Liste(Raum_Liste)

if forms.alert('Daten in Baugruppe, Kanäle und Auslässe schreiben?', ok=False, yes=True, no=True):
    Datenschreiben_Kanal(Sys_Liste, Kanal_Collector)
    Datenschreiben_Auslass(Sys_Liste,Auslass_Collector)
    Datenschreiben_BG(Sys_Liste_MitBGId,Baugruppen_collector)
def Liste_Bearbeiten(input_list):
    Eingang = []
    Ausgang = []
    Ein_Raum = []
    Aus_Raum = []
    Raum = []
    for item in input_list:
        Eingang.append([item[0],item[1],item[2]])
        Ausgang.append([item[1],item[0],item[2]])
        Ein_Raum.append(item[0])
        Aus_Raum.append(item[1])
        Raum.append(item[0])
        Raum.append(item[1])
    Ein_Raum = set(Ein_Raum)
    Ein_Raum = list(Ein_Raum)
    Aus_Raum = set(Aus_Raum)
    Aus_Raum = list(Aus_Raum)
    Raum = set(Raum)
    Raum = list(Raum)
    return Eingang,Ausgang,Ein_Raum,Aus_Raum,Raum

def Raum_Daten(Eingang,Ausgang,Ein_Raum,Aus_Raum,Raum):
    Eingangsdaten = []
    Ausgangsdaten = []
    Raumdaten = []
    Raum_Text = []
    for item in Ein_Raum:
        raum = [item]
        for daten in Eingang:
            if daten[0] == item:
                raum.append([daten[1],daten[2]])
        Eingangsdaten.append(raum)
    for item in Aus_Raum:
        raum = [item]
        for daten in Ausgang:
            if daten[0] == item:
                raum.append([daten[1],daten[2]])
        Ausgangsdaten.append(raum)
    for item in Raum:
        raum_ein = []
        raum_aus = []
        for ein in Eingangsdaten:
            if ein[0] == item:
                raum_ein = ein[1:]

        for aus in Ausgangsdaten:
            if aus[0] == item:
                raum_aus = aus[1:]
        Raumdaten.append([item,raum_ein,raum_aus])

    for item in Raumdaten:
        if len(item[1]) > 0:
            output.print_table(
                table_data=item[1],
                title=item[0],
                columns=['Aus Raum', 'Volumen']
            )
            if len(item[2]) > 0:
                output.print_table(
                    table_data=item[2],
                    columns=['In Raum', 'Volumen'])
        else:
            if len(item[2]) > 0:
                output.print_table(
                    table_data=item[2],
                    title=item[0],
                    columns=['In Raum', 'Volumen'])

    for item in Raumdaten:
        text_raum = ''
        text_ein = ''
        text_aus = ''
        if len(item[1]) > 0:
            text_ein = 'Aus: [m3/h]--'
            for i in item[1]:
                text_ein = text_ein + i[0] + ': ' + str(int(i[1])) + ', '
        if len(item[2]) > 0:
            text_aus = 'In: [m3/h]--'
            for i in item[2]:
                text_aus = text_aus + i[0] + ': ' + str(int(i[1])) + ', '
        text_raum = text_ein + text_aus
        Raum_Text.append([item[0],text_raum])

    return Eingangsdaten, Ausgangsdaten, Raumdaten, Raum_Text

def Luftsumme(Raum_Daten):
    out = []
    for item in Raum_Daten:
        ein = 0
        aus = 0
        ges = 0
        if len(item[1]) > 0:
            for i in item[1]:
                ein = ein + i[1]
        if len(item[2]) > 0:
            for i in item[2]:
                aus = aus + i[1]
        ges = ein - aus
        out.append([item[0],ein,aus,ges])

    return out

def gesamt_list(luftsumme,luftverteilung):
    out = []
    for item in luftsumme:
        for item2 in luftverteilung:
            if str(item[0]) == str(item2[0]):
                out.append([item[0],item[1],item[2],item[3],item2[1]])
    output.print_table(
        table_data=out,
        title='Überstrom',
        columns=['Raum', 'Überstrom_In','Überstrom_Aus','ÜberstromSumme','Luftverteilung'])

    return out

def Datenschreiben(Raumlist,Familie):
    Title = '{value}/{max_value} Daten schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Daten schreiben")
        t.Start()

        for item in Raumlist:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Raumlist))

            if not item[0] in Familie.keys():
                continue
            Item = Familie[item[0]]

            Item.LookupParameter('IGF_RLT_ÜberströmungSummeIn').SetValueString(str(item[1]))
            Item.LookupParameter('IGF_RLT_ÜberströmungSummeAus').SetValueString(str(item[2]))
            Item.LookupParameter('IGF_RLT_ÜberströmungRaum').SetValueString(str(item[3]))
            Item.LookupParameter('IGF_RLT_ÜberstromVerteilung').Set(str(item[4]))

        t.Commit()


Kanaldaten = [[el[2],el[1],el[3]] for el in Sys_Liste]
Eingang_liste,Ausgang_liste,Ein_Raum_liste,Aus_Raum_liste,Raum_liste = Liste_Bearbeiten(Kanaldaten)
Eingangsdaten, Ausgangsdaten, Raumdaten, Raum_Text = Raum_Daten(Eingang_liste,Ausgang_liste,Ein_Raum_liste,Aus_Raum_liste,Raum_liste)
Luftsumme_liste = Luftsumme(Raumdaten)
gesamt_list = gesamt_list(Luftsumme_liste,Raum_Text)
Raumdic = {}
for el in spaces_collector:
    Raumnummer = get_value(el.LookupParameter('Nummer'))
    Raumdic[Raumnummer] = el

if forms.alert('Daten in MEP-Räume schreiben?', ok=False, yes=True, no=True):
    Datenschreiben(gesamt_list, Raumdic)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
