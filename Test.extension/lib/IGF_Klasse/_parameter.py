# coding: utf8
from IGF_Klasse import DB
doc = __revit__.ActiveUIDocument.Document

class SharedParameter(object):
    def __init__(self, name = '',guid = '',type = '',disziplin = '',userinfo = '', group = '', externaldefinition = ''):
        self.name = name
        self.guid = guid
        self.type = type
        self.disziplin = disziplin
        self.userinfo = userinfo
        self.externaldefinition = externaldefinition
        self.group = group
        self.set_up_shared()
    
    def get_grundinfo(self):
        if self.externaldefinition:
            self.name = self.externaldefinition.Name
            self.userinfo = self.externaldefinition.Description
            self.guid = self.externaldefinition.GUID.ToString()
            self.group = self.externaldefinition.OwnerGroup.Name
    
    def get_type(self):
        if self.externaldefinition:
            try:
                self.type = DB.LabelUtils.GetLabelFor(self.externaldefinition.ParameterType)
            except:
                self.type = DB.LabelUtils.GetLabelForSpec(self.externaldefinition.GetDataType())
    
    def get_discipline(self):
        if self.externaldefinition:
            try:
                diszipline = DB.UnitUtils.GetUnitGroup(self.externaldefinition.UnitType)
                if diszipline.ToString() == 'Common':self.disziplin = 'Allgemein'
                elif diszipline.ToString() == 'Energy':self.disziplin = 'Energie'
                elif diszipline.ToString() == 'Structural':self.disziplin = 'Tragwerk'
                elif diszipline.ToString() == 'HVAC':self.disziplin = 'Lüftung'
                elif diszipline.ToString() == 'Piping':self.disziplin = 'Rohre'
                elif diszipline.ToString() == 'Electrical':self.disziplin = 'Elektro'

            except:
                self.disziplin = DB.LabelUtils.GetLabelForDiscipline(DB.UnitUtils.GetDiscipline(self.externaldefinition.GetDataType()))
    
    def set_up_shared(self):
        self.get_grundinfo()
        self.get_type()
        self.get_discipline()

class ProjektParameter(SharedParameter):
    def __init__(self, bindingmap = None):
        SharedParameter.__init__(self)
        self.bindingmap = bindingmap
        self.typOrex = ''
        self.externaldefinition = ''
        self.sharedparam = ''
        self.binding = ''
        self.internaldefinition = ''
        self.cates = ''
        self.set_up_projekt()
        self.set_up_shared()
    
    def get_binding(self):
        if self.bindingmap:
            self.binding = self.bindingmap.Current
            self.internaldefinition = self.bindingmap.Key
    
    def get_paramtyp(self):
        if self.binding:
            if self.binding.GetType().ToString() == 'Autodesk.Revit.DB.InstanceBinding':
                self.typOrex = 'Exemplar'
            else:
                self.typOrex = 'Type'

    def get_sharedparameter(self):
        if self.internaldefinition:
            self.sharedparam = self.doc.GetElement(self.internaldefinition.Id)
            self.guid = self.sharedparam.GuidValue.ToString()
    
    def get_externaldefinition(self):
        if self.guid:
            file = __revit__.Application.OpenSharedParameterFile()
            for g in file.Groups:
                if g:
                    for d in g.Definitions:
                        if d.GUID.ToString() == self.guid:
                            self.externaldefinition = d
                            return

    def get_paramgroup(self):
        if self.internaldefinition:
            try:
                self.paramgroup = DB.LabelUtils.GetLabelFor(self.internaldefinition.ParameterGroup)
            except:
                self.paramgroup = DB.LabelUtils.GetLabelForGroup(self.internaldefinition.GetGroupTypeId())

    def get_cates(self):
        if self.binding:
            cates = self.binding.Categories
            cateName = ''
            for cate in cates:
                cateName = cateName + cate.Name + ','
            self.cates = cateName[:-1]
    
    def set_up_projekt(self):
        self.get_binding()
        self.get_paramtyp()
        self.get_sharedparameter()
        self.get_externaldefinition()
        self.get_cates()
        self.get_paramgroup()
        