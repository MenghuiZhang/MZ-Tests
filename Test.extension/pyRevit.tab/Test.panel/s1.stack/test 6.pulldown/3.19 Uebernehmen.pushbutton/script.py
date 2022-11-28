
# coding: utf8
from pyrevit import script, forms
from rpw import revit,DB
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal


__title__ = "Schema"
__doc__ = """Luftmengenberechnung für Projekt MFC"""
__author__ = "Menghui Zhang"


uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

Parameter_Dict = {}
exapplication = ex.ApplicationClass()

def Element_TypFilter(Fam_name = None, Typname = None, ActiveviewOn = None):

    param_equality=DB.FilterStringContains()

    param_equality1=DB.FilterStringEquals()

    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)

    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)

    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)

    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)


    Fam_typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)

    Fam_typ_prov = DB.ParameterValueProvider(Fam_typ_id)

    Fam_typ_value_rule = DB.FilterStringRule(Fam_typ_prov, param_equality, Typname, True)

    Fam_typ_filter = DB.ElementParameterFilter(Fam_typ_value_rule)


    coll = DB.FilteredElementCollector(doc,active_view.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType().ToElements()

    return coll


logger = script.get_logger()
output = script.get_output()

def Funktion00():
    from IGF_lib import get_value
    import xlsxwriter
    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    # Liste = [['Nummer','Name','DeS-Heizleistung','Heizlast','DeS-Kühlleistung','Kühllast']]
    Liste = [['Nummer','Name','Kühlleistung','Kühllast']]
    Dict = {}
    for el in Bauteile:
        mep = el.Space[doc.GetElement(el.CreatedPhaseId)].Id.ToString()
        if mep not in Dict.keys():
            Dict[mep] = []
        Dict[mep].append(el)
    for Id in Dict.keys():
        mep = doc.GetElement(DB.ElementId(int(Id)))
        nummer = mep.Number
        name = mep.LookupParameter('Name').AsString()
        # Heizlast = get_value(mep.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD'))
        Kuehllast = get_value(mep.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD'))
        hl = 0
        kl = 0
        for des in Dict[Id]:
            # hl += get_value(des.LookupParameter('Power'))
            kl += get_value(des.LookupParameter('MC Cooling Power'))
        # Liste.append([nummer,name,int(hl),int(Heizlast),int(kl),int(Kuehllast)])
        Liste.append([nummer,name,int(kl),int(Kuehllast)])

    workbook = xlsxwriter.Workbook(r'C:\Users\Zhang\Desktop\Schema\ULK.xlsx')
    worksheet = workbook.add_worksheet()

    for col in range(len(Liste[0])):
        for row in range(len(Liste)):
             
            wert = Liste[row][col]
            worksheet.write(row, col, wert)
            

    worksheet.autofilter(0, 0, int(len(Liste))-1, int(len(Liste[0])-1))
    workbook.close()


def Funktion0():
    """Info in MEP Räume in Schema schreiben"""
    excelPath = r'C:\Users\Zhang\Desktop\Schema\MEP-Raumliste ULK.xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1): 
            nummer = sheet.Cells[row, 1].Value2
            if nummer == '' or  nummer == None:
                continue
            else:
                name = sheet.Cells[row, 2].Value2
                temp1 = sheet.Cells[row, 3].Value2
                temp2 = sheet.Cells[row, 4].Value2
                temp3 = sheet.Cells[row, 5].Value2
                temp4 = sheet.Cells[row, 6].Value2
                Parameter_Dict[nummer] = [name,temp1,temp2,temp3,temp4]

    book1.Save()
    book1.Close()
    mep = DB.FilteredElementCollector(doc,active_view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElements()
    t =DB.Transaction(doc,'MEP-Räume')
    t.Start()
    for el in mep:
        if el.Number in Parameter_Dict.keys():
            # if el.Number in ['100.148','110.148','120.148','130.139']:
            el.LookupParameter('Name').Set(Parameter_Dict[el.Number][0])
            el.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').SetValueString(str(Parameter_Dict[el.Number][1]))
            el.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').SetValueString(str(Parameter_Dict[el.Number][2]))
            el.LookupParameter('LIN_BA_DESIGN_COOLING_TEMPERATURE').SetValueString(str(Parameter_Dict[el.Number][3]))
            el.LookupParameter('LIN_BA_DESIGN_HEATING_TEMPERATURE').SetValueString(str(Parameter_Dict[el.Number][4]))
        else:print(el.Number)
    t.Commit()

def Funktion1():
    """RV ID von ULK in RV in Schema"""

    def Element_TypFilter(Fam_name = None, Typname = None, ActiveviewOn = None):
        param_equality=DB.FilterStringContains()
        param_equality1=DB.FilterStringEquals()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality1,Fam_name,True)
        Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)
        Fam_typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        Fam_typ_prov = DB.ParameterValueProvider(Fam_typ_id)
        Fam_typ_value_rule = DB.FilterStringRule(Fam_typ_prov, param_equality1, Typname, True)
        Fam_typ_filter = DB.ElementParameterFilter(Fam_typ_value_rule)
        coll = DB.FilteredElementCollector(doc,active_view.Id).WherePasses(Fam_name_filter).WherePasses(Fam_typ_filter).WhereElementIsNotElementType().ToElements()
        # coll = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()

        return coll

    class ULK:
        def __init__(self,elem):
            self.elem = elem
            self.vsr = ''
            self.liste = []
            if self.Muster_Pruefen():return
            self.get_vsr(self.elem)
            
            self.elemID = self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString()
            self.vsrID = 'RVI_'+self.elemID[4:]

        def Muster_Pruefen(self):
            '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
            try:
                bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
                if bb == 'KG4xx_Musterbereich':return True
                else:return False
            except:return False
        def get_vsr(self,elem):
            if self.vsr:return
            elemid = elem.Id.ToString()
            self.liste.append(elemid)
            cate = elem.Category.Name

            if not cate in ['Rohr Systeme','Rohrdämmung']:
                conns = None
                try:conns = elem.MEPModel.ConnectorManager.Connectors
                except:
                    try:conns = elem.ConnectorManager.Connectors
                    except:pass
                if conns:

                    if conns.Size > 2 and cate != 'HLS-Bauteile':return
                    for conn in conns:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste:
                                if owner.Category.Name == 'Rohrzubehör':
                                    faminame = owner.Symbol.FamilyName
                                    if faminame.upper().find('CONTROL') != -1:
                                        self.vsr = owner
                                        return
                                                                            
                                self.get_vsr(owner)
        def wert_schreiben(self):
            self.vsr.LookupParameter('ZS_MARK').Set(self.vsrID)
            
    Liste = []
    Bauteile = Element_TypFilter('SBI_DET_ULK','SBI_ULK_OOD')  
    # Bauteile = Element_TypFilter('MC_K_IGF_FG_Flex Geko','SBI_ULK_OOD')   
    print(Bauteile.Count) 
    for el in Bauteile:
        ulk = ULK(el)

        if ulk.Muster_Pruefen():continue
        Liste.append(ULK(el))

    t = DB.Transaction(doc,'VSR Id')
    t.Start()
    for el in Liste:
        if not el.vsr:
            print(el.elem.Id)
            continue
        try:
            el.wert_schreiben()
        except Exception as e:print(e)
    t.Commit()

def Funktion2():
    """RV ID in Haupt Modell schreiben"""
    class RV:
        def __init__(self,elemid,raumnr):
            self.elemid = elemid
            self.raumnr = raumnr
            self.vrid = self.raumnr
            self.elem = doc.GetElement(DB.ElementId(int(self.elemid)))
            

        def wert_schreiben(self):
            try:self.elem.LookupParameter('SBI_Bauteilnummerierung').Set(self.vrid)
            except:print(self.elemid)

    excelPath = r'C:\Users\Zhang\Desktop\Rohrzubehör-Bauteilliste1.xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    liste = []

    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1 ): 
            ID = sheet.Cells[row, 1].Value2
            temp1 = sheet.Cells[row, 3].Value2
            
            if ID == '' or ID == None:continue
            liste.append(RV(ID,temp1))
            
           
    book1.Save()
    book1.Close()

    t = DB.Transaction(doc,'Test')
    t.Start()
    for el in liste:
        if el.elem == None:
            print(el.vrid)
        else:
            el.wert_schreiben()
    t.Commit()

def Funktion3():
    """RV ID von ULK in RV in Schema"""

    def Element_TypFilter(Fam_name = None, Typname = None, ActiveviewOn = None):
        param_equality=DB.FilterStringContains()
        param_equality1=DB.FilterStringEquals()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)
        Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)
        Fam_typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        Fam_typ_prov = DB.ParameterValueProvider(Fam_typ_id)
        Fam_typ_value_rule = DB.FilterStringRule(Fam_typ_prov, param_equality1, Typname, True)
        Fam_typ_filter = DB.ElementParameterFilter(Fam_typ_value_rule)
        coll = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
        # coll = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()

        return coll

    class RV:
        def __init__(self,elem):
            self.elem = elem           
            self.elemID = self.elem.LookupParameter('ZS_MARK').AsString()
            self.vsrID = 'RVD_'+self.elemID[3:]

        
        def wert_schreiben(self):
            self.elem.LookupParameter('ZS_MARK').Set(self.vsrID)
            
    Liste = []
    Bauteile = Element_TypFilter('SBI_DET_W_VPIC_Pressure_Independent_Control_Valve','SBI_ULK_OOD')  
    # Bauteile = Element_TypFilter('MC_K_IGF_FG_Flex Geko','SBI_ULK_OOD')   
    print(Bauteile.Count) 
    for el in Bauteile:
        rv = RV(el)

        Liste.append(rv)

    t = DB.Transaction(doc,'VSR Id')
    t.Start()
    for el in Liste:
        try:
            el.wert_schreiben()
        except Exception as e:print(e)
    t.Commit()

def Funktion4():
    """HK Nummer von MFC in Schema anhand von Raumnummer"""
    excelPath = r'C:\Users\Zhang\Desktop\HK_MFC_Final.xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1): 
            ID = sheet.Cells[row, 1].Value2
            temp1 = sheet.Cells[row, 5].Value2
            if ID == '' or ID == None:continue

            if ID not in Parameter_Dict.keys():

                Parameter_Dict[ID] = []
            Parameter_Dict[ID].append(temp1)
            

    book1.Save()
    book1.Close()

    param_equality=DB.FilterStringContains()
    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,'SBI_DET_WH_BTL_HZK',True)
    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType().ToElements()

    dict_neu = {}

    for el in Bauteile:
        space = el.Space[doc.GetElement(el.CreatedPhaseId)]
        if not space:
            continue
        else:
            if space.Number not in dict_neu.keys():
                dict_neu[space.Number] = []
            dict_neu[space.Number].append(el)
    # for el in Parameter_Dict.keys():
    #     if el not in dict_neu.keys():
    #         print(el,'MEP nicht gezeichnet')
    t =DB.Transaction(doc,'HK')
    t.Start()
    for el in dict_neu.keys():
        if el in Parameter_Dict.keys():
            if len(Parameter_Dict[el]) == len(dict_neu[el]):
                if len(Parameter_Dict[el]) == 1:
                    dict_neu[el][0].LookupParameter('ZS_MARK').Set(Parameter_Dict[el][0])
                else:
                    print(el,'Manuell')
            else:
                print(el,'Anzahl nicht passt')
        else:
            print(el,'MEP falsch gezeichnet')
    t.Commit()

def Funktion5():
    """HK Leistung von MFC in Schema mit hilfe von Bauteilid"""
    excelPath = r'C:\Users\Zhang\Desktop\HK_MFC.xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1): 
            ID = sheet.Cells[row, 5].Value2
            hl = sheet.Cells[row, 10].Value2
            height = sheet.Cells[row, 11].Value2
            width = sheet.Cells[row, 12].Value2
            tiefe = sheet.Cells[row, 13].Value2
            if ID == '' or ID == None:continue
            Parameter_Dict[ID]= [hl,str(int(height))+'x'+str(int(width))+'x'+str(int(tiefe))]

    book1.Save()
    book1.Close()

    param_equality=DB.FilterStringContains()
    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,'SBI_DET_WH_BTL_HZK',True)
    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType().ToElements()
    
    t = DB.Transaction(doc,'HK')
    t.Start()
    for el in Bauteile:
        Id = el.LookupParameter('ZS_MARK').AsString()
        if Id in Parameter_Dict.keys():
            el.LookupParameter('HHA_POWER').SetValueString(str(Parameter_Dict[Id][0]))
            el.LookupParameter('ZS_DESCR').Set(str(Parameter_Dict[Id][1]))
        else:
            print(Id)
    t.Commit()

def Funktion6():
    param_equality=DB.FilterStringEquals()
    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,'_H_IGF_422_MC_Charleston-2-1200-3000',True)
    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    Typ_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    Typ_name_prov=DB.ParameterValueProvider(Typ_name_id)
    Typ_name_value_rule=DB.FilterStringRule(Typ_name_prov,param_equality,'Charleston-2200-552_15',True)
    Typ_name_filter = DB.ElementParameterFilter(Typ_name_value_rule)

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    t = DB.Transaction(doc,'Test')
    t.Start()
    for el in Bauteile:
        p = el.LookupParameter('Power').AsDouble()
        if p < 0.1:
            el.LookupParameter('Power').SetValueString('300')

        
    t.Commit()
    
def Funktion7():
    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    class HK:
        def __init__(self,elem):
            self.elem = elem
            self.rv = ''
            self.liste = []  
            self.get_des(self.elem)
            
        def get_des(self,elem):
            if self.rv:return
            if len(self.liste) > 5:return
            elemid = elem.Id.ToString()
            self.liste.append(elemid)
            cate = elem.Category.Name

            if cate not in ['Rohr Systeme','Rohrdämmung']:
                conns = None
                try:conns = elem.MEPModel.ConnectorManager.Connectors
                except:
                    try:conns = elem.ConnectorManager.Connectors
                    except:pass
                if not conns:return
                for conn in conns:
                    if elemid == self.elem.Id.ToString():
                        if conn.Direction.ToString() != 'In':
                            continue
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        
                        if owner.Id.ToString() not in self.liste:
                            if owner.Category.Name in ['Rohrzubehör']:
                                faminame = owner.Symbol.FamilyName
                                self.rv = owner
                                return                                 
                            self.get_des(owner)
        def printinfo(self):
            
            if self.rv:pass

                #print(self.rv.LookupParameter('SBI_Bauteilnummerierung').AsString(),self.rv.Symbol.FamilyName)
            else:
                print(self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString())
                print('kein rv')
                print('-----------')
        def wert_schreiben(self):
            self.rv.LookupParameter('SBI_Bauteilnummerierung').Set('THV_'+self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString()[3:])

    
    # for el in Bauteile:HK(el).printinfo()
        
    t = DB.Transaction(doc,'Test')
    t.Start()
    for el in Bauteile:
        HK(el).wert_schreiben()
    t.Commit()

def Funktion8():
    """HK Typ ändern"""
    excelPath = r'C:\Users\Zhang\Desktop\Schema\HK_MFC_Final.xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1): 
            ID = sheet.Cells[row, 5].Value2
            temp1 = sheet.Cells[row, 6].Value2
            temp2 = sheet.Cells[row, 7].Value2
            if ID == '' or ID == None:continue
            if temp2 == 'Seitlich':continue
            Parameter_Dict[ID] = temp1
    book1.Save()
    book1.Close()

    param_equality=DB.FilterStringContains()
    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,'SBI_DET_WH_BTL_HZK',True)
    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType().ToElements()

    t = DB.Transaction(doc,'HK')
    t.Start()
    for el in Bauteile:
        iD = el.LookupParameter('ZS_MARK').AsString()
        if iD not in Parameter_Dict.keys():
            print(iD,el.Symbol.FamilyName)
            continue
        if Parameter_Dict[iD] == 'Charleston':
            if el.Symbol.Id.ToString() != '36282581':el.ChangeTypeId(DB.ElementId(36282581))
        elif Parameter_Dict[iD] == 'Radiapanel':
            if el.Symbol.Id.ToString() != '36161708':el.ChangeTypeId(DB.ElementId(36161708))
        else:
            print(Parameter_Dict[iD])
    t.Commit()

def Funktion9():
    param_equality=DB.FilterStringEquals()
    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,'MC_ConnectionNode_Hydronic',True)
    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    Typ_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    Typ_name_prov=DB.ParameterValueProvider(Typ_name_id)
    Typ_name_value_rule=DB.FilterStringRule(Typ_name_prov,param_equality,'Segel-Heizen',True)
    Typ_name_filter = DB.ElementParameterFilter(Typ_name_value_rule)

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    Dict_ = {}
    for el in Bauteile:
        nr = el.Space[doc.GetElement(el.CreatedPhaseId)].Number
        if nr == '120.134c':nr = '120.135c'
        if nr == '120.133c':nr = '120.135c'
        if nr == '130.129':nr = '130.131'
        if nr not in Dict_.keys():
            Dict_[nr] = 0
        hl = el.LookupParameter('Power').AsDouble()
        Dict_[nr] += hl 

    Bauteile1 = DB.FilteredElementCollector(doc).WherePasses(Fam_name_filter).WherePasses(Typ_name_filter).WhereElementIsNotElementType().ToElements()
    liste = []
    t = DB.Transaction(doc,'Segel')
    t.Start()
    for el in Bauteile1:
        nr = el.LookupParameter('SBI_Bauteilnummerierung').AsString()[4:]
        liste.append(nr)
        if nr in ['120.347','110.347','100.305','130.149','120.148','110.148','100.148']:
            continue
        if nr not in Dict_.keys():
            print(nr)
        else:
            hl = Dict_[nr]
            el.LookupParameter('Power').Set(hl)
    print('-------------------------')
    for el in Dict_.keys():
        if el not in liste:
            print(el)
    t.Commit()

def Funktion10():
    """HK Leistung von MFC in Schema mit hilfe von Bauteilid"""
    excelPath = r'C:\Users\Zhang\Desktop\Schema\IGF_Regelventile ULK - Kopie (2).xlsx'
    book1 = exapplication.Workbooks.Open(excelPath)
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1): 
            ID = sheet.Cells[row, 1].Value2
            hl = sheet.Cells[row, 2].Value2

            Parameter_Dict[ID]= hl

    book1.Save()
    book1.Close()

    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    
    t = DB.Transaction(doc,'ULK')
    t.Start()
    for n,el in enumerate(Bauteile):
        print(n)
        Id = el.LookupParameter('ZS_MARK').AsString()
        if Id in Parameter_Dict.keys():
            el.ChangeTypeId(DB.ElementId(int(Parameter_Dict[Id])))
        else:
            print(Id)
    t.Commit()

def Funktion11():
    cl = uidoc.Selection.GetElementIds()
    
    t = DB.Transaction(doc,'Segel')
    t.Start()
    for elid in cl:
        el =doc.GetElement(elid)
        if el.Symbol.FamilyName.find('_K_IGF_435_Deckensegelverbinder') != -1:
            hl = el.LookupParameter('Leistung_Heizen_Hersteller').AsDouble()
            kl = el.LookupParameter('Leistung_Kühlen_Hersteller').AsDouble()
            
            el.LookupParameter('Leistung_Heizen_Hersteller').Set(hl*2)
            el.LookupParameter('Leistung_Kühlen_Hersteller').Set(kl*2)
        elif el.Symbol.FamilyName.find('_K_IGF_435_Deckensegel') != -1:
            el.LookupParameter('Leistung_Heizen_Hersteller').SetValueString('0')
            el.LookupParameter('Leistung_Kühlen_Hersteller').SetValueString('0')
            
        
    t.Commit()

def Funktion12():
    cl = uidoc.Selection.GetElementIds()
    
    dict_hls = {}
    dict_zu = {}
    for elid in cl:
        el =doc.GetElement(elid)
        if el.Category.Name == 'HLS-Bauteile':
            if el.Symbol.FamilyName == 'MC_ConnectionNode_Hydronic':
                dict_hls[el] = el.Location.Point
        elif el.Category.Name == 'Rohrzubehör':
            if el.Symbol.FamilyName.find('6-Wege'):
                Id = el.LookupParameter('SBI_Bauteilnummerierung').AsString()
                if not Id:
                    continue
                else:
                    dict_zu[el.Location.Point] = Id
    if len(dict_hls) != len(dict_zu):
        print('Anzahl passt nicht')
        return
    neu = {}
    for el0 in dict_hls.keys():
        dis = 100
        for el1 in dict_zu.keys():
            di = el1.DistanceTo(dict_hls[el0])
            if di < dis:
                dis = di
                neu[el0] = dict_zu[el1]
    
    test0 = neu.values()[:]
    test1 =set(test0)
    test2 = list(test1)
    if len(test1) != len(test2):
        print('elem bisitzen gleiche Id')
        return
    t = DB.Transaction(doc,'Segel')
    t.Start()
    for el in neu.keys():
        el.LookupParameter('SBI_Bauteilnummerierung').Set('DBH_'+neu[el][4:])
        el.ChangeTypeId(DB.ElementId(37932325))  
    t.Commit()
    print('erledigt!',len(neu))

def Funktion13():
    Bauteile = DB.FilteredElementCollector(doc,active_view.Id).WhereElementIsNotElementType().ToElements()
    
    class RV:
        def __init__(self,elem):
            self.elem = elem
            self.des = ''
            self.liste = []  
            print(self.elem.Id)
            self.get_des(self.elem)
            
        def get_des(self,elem):
            if self.des:return
            if len(self.liste) > 500:return
            elemid = elem.Id.ToString()
            self.liste.append(elemid)
            cate = elem.Category.Name

            if cate not in ['Rohr Systeme','Rohrdämmung']:
                conns = None
                try:conns = elem.MEPModel.ConnectorManager.Connectors
                except:
                    try:conns = elem.ConnectorManager.Connectors
                    except:pass
                if not conns:return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        
                        if owner.Id.ToString() not in self.liste:
                            if owner.Category.Name in ['Rohrzubehör','HLS-Bauteile']:
                                faminame = owner.Symbol.FamilyName
                                if faminame.find('Segel') != -1 or faminame.find('Geko') != -1 or faminame.find('ULK') != -1:
                                    return
                                elif faminame.find('6-Wege') != -1 or faminame.find('6 Wege') != -1:
                                    self.des = owner
                                    return                                 
                            self.get_des(owner)

        def wert_schreiben(self):
            
            self.des.LookupParameter('SBI_Bauteilnummerierung').Set('6WV_'+self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString()[4:])

    t = DB.Transaction(doc,'Test')
    t.Start()
    for el in Bauteile:
        RV(el).wert_schreiben()
    t.Commit()


    
Funktion10()






# Bauteile = Element_TypFilter('SBI_DET_W_VPIC_Pressure_Independent_Control_Valve','')
# liste = []
# for el in Bauteile:
#     if el.Symbol.FamilyName[-1] != 'V':
#         print(el.Symbol.FamilyName)
#         liste.append(el)
# print(len(liste))

# t = DB.Transaction(doc,'Test')
# t.Start()
# for el in liste:
#     try:
#         el.LookupParameter('Raumnummer').Set(el.LookupParameter('Nummer_temp').AsString())
#     except:pass
# t.Commit()

# VSR ID
# class RV:
#     def __init__(self,elem):
#         self.elem = elem
#         self.des = ''
#         self.liste = []
#         if self.Muster_Pruefen():return       
#         self.raumnummer = ''
#         self.get_raumnr()
        

#     def Muster_Pruefen(self):
#         '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
#         try:
#             bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
#             if bb == 'KG4xx_Musterbereich':return True
#             else:return False
#         except:return False
#     def get_des(self,elem):
#         if self.des:return
#         if len(self.liste) > 500:return
#         elemid = elem.Id.ToString()
#         self.liste.append(elemid)
#         cate = elem.Category.Name

#         if cate not in ['Rohr Systeme','Rohrdämmung']:
#             conns = None
#             try:conns = elem.MEPModel.ConnectorManager.Connectors
#             except:
#                 try:conns = elem.ConnectorManager.Connectors
#                 except:pass
#             if not conns:return
#             for conn in conns:
#                 refs = conn.AllRefs
#                 for ref in refs:
#                     owner = ref.Owner
                    
#                     if owner.Id.ToString() not in self.liste:
#                         if owner.Category.Name in ['Rohrzubehör','HLS-Bauteile']:
#                             faminame = owner.Symbol.FamilyName
#                             if faminame.find('Segel') != -1 or faminame.find('Geko') != -1 or faminame.find('ULK') != -1:
#                                 self.des = owner
#                                 return
#                             elif faminame.find('_W_V3W') != -1:
#                                 return                                 
#                         self.get_des(owner)
#     def get_raumnr(self):
#         if self.elem.LookupParameter('Systemtyp').AsValueString().find('KLD') == -1:
#             return
#         self.get_des(self.elem)
#         try:
#             self.raumnummer = self.des.Space[doc.GetElement(self.des.CreatedPhaseId)].Number
#         except:
#             print(self.elem.Id,'Nummer nicht ermittelt')

#     def wert_schreiben(self):
#         if self.elem.LookupParameter('Raumnummer').AsString() == '' or self.elem.LookupParameter('Raumnummer').AsString() == None:
#             self.elem.LookupParameter('Raumnummer').Set(self.raumnummer)

# class RV:
#     def __init__(self,elem):
#         self.elem = elem

#     def wert_schreiben(self):
#         try:self.elem.LookupParameter('ZS_MARK').Set('RV_' + self.elem.LookupParameter('Raumnummer').AsString())
#         except:print(self.elem.Symbol.FamilyName,self.elem.Id)



# 

# excelPath = r'C:\Users\Zhang\Desktop\Bauteilliste ULK Schema.xlsx'


# book1 = exapplication.Workbooks.Open(excelPath)


# sheetscount = book1.Worksheets.Count
# n = 0
# for sheet in book1.Worksheets:
#     rows = sheet.UsedRange.Rows.Count
#     for row in range(2, rows + 1): 
#         ID = sheet.Cells[row, 2].Value2
#         temp1 = sheet.Cells[row, 3].Value2
#         if ID == '':continue
#         Parameter_Dict[ID] = temp1
#         # nummer = sheet.Cells[row, 1].Value2
#         # if nummer == '':
#         #     continue
#         # else:
#         #     name = sheet.Cells[row, 2].Value2
#         #     temp1 = sheet.Cells[row, 3].Value2
#         #     if nummer not in Parameter_Dict.keys():
#         #         Parameter_Dict[nummer] = []

#         #     Parameter_Dict[nummer].append([name,temp1])

# book1.Save()
# book1.Close()

# dict_neu = {}
# for el in Bauteile:
#     id = el.LookupParameter('ZS_MARK').AsString()
#     dict_neu[id] = el
# for el in Bauteile:
#     space = el.Space[doc.GetElement(el.CreatedPhaseId)]
#     if not space:
#         continue
#     else:
#         if space.Number not in dict_neu.keys():
#             dict_neu[space.Number] = []
#         dict_neu[space.Number].append(el)

# for el in Parameter_Dict.keys():
#     if el in dict_neu.keys():
#         pass
#     else:print(el,'MEP nicht gezeichnet')
# ULK
# for el in Parameter_Dict.keys():
#     if el not in dict_neu.keys():print(el)
# t =DB.Transaction(doc,'ULK')
# t.Start()
# for el in dict_neu.keys():
#     if el in Parameter_Dict.keys():
#         dict_neu[el].LookupParameter('HCA_POWER').SetValueString(str(Parameter_Dict[el]))
#     else:
#         print(el)
#     #     if len(Parameter_Dict[el]) == len(dict_neu[el]):
#     #         if len(Parameter_Dict[el]) == 1:
#     #             dict_neu[el][0].LookupParameter('HCA_POWER').SetValueString(str(Parameter_Dict[el][0][1]))
#     #             dict_neu[el][0].LookupParameter('ZS_MARK').Set(Parameter_Dict[el][0][0])
#     #         else:
#     #             print(el,'Manuell')
#     #     else:
#     #         print(el,'Anzahl nicht passt')
#     # else:
#     #     print(el,'MEP nicht passt')
# t.Commit()

##VSR ID
# class ULK:
#     def __init__(self,elem):
#         self.elem = elem
#         self.vsr = ''
#         self.liste = []
#         if self.Muster_Pruefen():return
#         self.get_vsr(self.elem)
        
#         self.elemID = self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString()
#         self.vsrID = 'RV_'+self.elemID[4:]

#     def Muster_Pruefen(self):
#         '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
#         try:
#             bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
#             if bb == 'KG4xx_Musterbereich':return True
#             else:return False
#         except:return False
#     def get_vsr(self,elem):
#         if self.vsr:return
#         elemid = elem.Id.ToString()
#         self.liste.append(elemid)
#         cate = elem.Category.Name

#         if not cate in ['Rohr Systeme','Rohrdämmung']:
#             conns = None
#             try:conns = elem.MEPModel.ConnectorManager.Connectors
#             except:
#                 try:conns = elem.ConnectorManager.Connectors
#                 except:pass
#             if conns:

#                 if conns.Size > 2 and cate != 'HLS-Bauteile':return
#                 for conn in conns:
#                     refs = conn.AllRefs
#                     for ref in refs:
#                         owner = ref.Owner
#                         if owner.Id.ToString() not in self.liste:
#                             if owner.Category.Name == 'Rohrzubehör':
#                                 faminame = owner.Symbol.FamilyName
#                                 if faminame.upper().find('CONTROL') != -1:
#                                     self.vsr = owner
#                                     return
                                                                        
#                             self.get_vsr(owner)
#     def wert_schreiben(self):
#         self.vsr.LookupParameter('SBI_Bauteilnummerierung').Set(self.vsrID)
        
# Liste = []
# for el in Bauteile:
#     ulk = ULK(el)

#     if ulk.Muster_Pruefen():continue
#     Liste.append(ULK(el))

# t = DB.Transaction(doc,'VSR Id')
# t.Start()
# for el in Liste:
#     if not el.vsr:
#         print(el.elem.Id)
#         continue
#     try:
        
#         el.wert_schreiben()
#     except Exception as e:print(e)
# t.Commit()
# # name_liste = [] 
# # for el in mepraum:print(el.Number)
# #     if el.Number in name_liste:
# #         print(el.Number)
# #     else:
# #         name_liste.append(el.Number)

# excelPath = r'C:\Users\Zhang\Desktop\Rohrzubehör-Bauteilliste1.xlsx'



# book1 = exapplication.Workbooks.Open(excelPath)

# liste = []
# liste1 = []
# sheetscount = book1.Worksheets.Count
# n = 0
# for sheet in book1.Worksheets:
#     rows = sheet.UsedRange.Rows.Count
#     for row in range(2, rows + 1): 
#         temp1 = sheet.Cells[row, 2].Value2
#         temp2 = sheet.Cells[row, 3].Value2
#         if temp1 != None or temp1 != '': liste.append(temp1)
#         if temp2 != None or temp2 != '': liste1.append(temp2)
        
        # nummer = sheet.Cells[row, 1].Value2
        # if nummer == '':
        #     continue
        # else:
        #     name = sheet.Cells[row, 2].Value2
        #     temp1 = sheet.Cells[row, 3].Value2
        #     if nummer not in Parameter_Dict.keys():
        #         Parameter_Dict[nummer] = []

        #     Parameter_Dict[nummer].append([name,temp1])

# book1.Save()
# book1.Close()

# for el in liste:
#     if el not in liste1:
#         print(el)
# print(30*'-')
# for el in liste1:
#     if el not in liste:
#         print(el)
