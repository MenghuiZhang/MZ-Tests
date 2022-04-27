# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import clr
import time
from pyrevit.forms import WPFWindow, SelectFromList
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager

start = time.time()


__title__ = "1.30 Raum-/Schachtbeschriftung  "
__doc__ = """Beschriftung für Raum und Schacht"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
my_config = script.get_config()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
views = forms.select_views()

if not views:
    logger.error('Keine Anschits ausgewählt!')
    script.exit()


out = DB.FilteredElementCollector(doc).OfClass(DB.Family)


Beschriftungstype_dic = {}  

for item in out:
    category = item.FamilyCategory.Name
    if category == "MEP-Raumbeschriftungen":
        Tpy = item.GetFamilySymbolIds()
        for i in Tpy:
            famliyUtype = doc.GetElement(i).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
            Beschriftungstype_dic[famliyUtype] = i

class beschriftungen(WPFWindow):
    def __init__(self, xaml_file_name,beschriftungDict,views):
        self.Views = views
        self.Beschriftungstype_dic = beschriftungDict
        WPFWindow.__init__(self, xaml_file_name)
        self._set_comboboxes()
        self.read_config()

    def read_config(self):
        # check are last parameters available
        try:
            if my_config.raumBeschriftung not in self.Beschriftungstype_dic.keys():
                my_config.raumBeschriftung = ""
        except:
            pass
        try:
            if my_config.schachtBeschriftung not in self.Beschriftungstype_dic.keys():
                my_config.schachtBeschriftung = ""
        except:
            pass

        try:
            self.Raumbeschriftung.Text = str(my_config.raumBeschriftung)
        except:
            self.Raumbeschriftung.Text = my_config.raumBeschriftung = ""

        try:
            self.Schachtbeschriftung.Text = str(my_config.schachtBeschriftung)
        except:
            self.Schachtbeschriftung.Text = my_config.schachtBeschriftung = ""

    def write_config(self):
        my_config.raumBeschriftung = self.Raumbeschriftung.Text.encode('utf-8')
        my_config.schachtBeschriftung = self.Schachtbeschriftung.Text.encode('utf-8')
        script.save_config()

    def _set_comboboxes(self):
        self.Raumbeschriftung.ItemsSource = sorted(self.Beschriftungstype_dic.keys())
        self.Schachtbeschriftung.ItemsSource = sorted(self.Beschriftungstype_dic.keys())

    @property
    def raumBeschriftungID(self):
        p = self.Beschriftungstype_dic[self.Raumbeschriftung.Text]
        return p

    @property
    def schachtBeschriftungID(self):
        p = self.Beschriftungstype_dic[self.Schachtbeschriftung.Text]
        return p


    def run(self, sender, args):
        self.write_config()
        trans = DB.Transaction(doc,"MEP-Raum-Beschriftung")
        trans.Start()
        for view in self.Views:
            MEPRaum = DB.FilteredElementCollector(doc, view.Id) \
                .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
                .WhereElementIsNotElementType()
            spaces = MEPRaum.ToElementIds()
            logger.info("{} MEP Räume ausgewählt in Ansicht {}".format(len(spaces),view.Id))
            if not spaces:
                logger.error("Keine MEP Räume in Ansicht {} gefunden".format(view.Id))
                continue

            Tag = DB.FilteredElementCollector(doc, view.Id) \
                .OfCategory(DB.BuiltInCategory.OST_MEPSpaceTags)\
                .WhereElementIsNotElementType()
            spaceTags = Tag.ToElementIds()
            logger.info("{} MEP-Raum-Beschriftungen ausgewählt in Ansicht {}".format(len(spaceTags),view.Id))

            try:
                vorhandenTag = []
                tagUspace = []
                for tag in Tag:
                    mep = tag.Space.Id
                    vorhandenTag.append(mep)
                    tagUspace.append([tag,mep])

                for raum in MEPRaum:
                    isSchacht = raum.LookupParameter("TGA_RLT_InstallationsSchacht").AsInteger()
                    if any(vorhandenTag):
                        if raum.Id in vorhandenTag:
                            for item in tagUspace:
                                if item[1] == raum.Id:                     
                                    if isSchacht == 1:
                                        item[0].ChangeTypeId(self.schachtBeschriftungID)
                                    else:
                                        item[0].ChangeTypeId(self.raumBeschriftungID)
                                        break
                        else:
                            uv = DB.UV(raum.Location.Point.X,raum.Location.Point.Y)
                            newtag = doc.Create.NewSpaceTag(raum,uv,view)
                            if isSchacht == 1:
                                newtag.ChangeTypeId(self.schachtBeschriftungID)
                            else:
                                newtag.ChangeTypeId(self.raumBeschriftungID)
                    else:
                        uv = DB.UV(raum.Location.Point.X,raum.Location.Point.Y)
                        newtag = doc.Create.NewSpaceTag(raum,uv,view)
                        if isSchacht == 1:
                            newtag.ChangeTypeId(self.schachtBeschriftungID)
                        else:
                            newtag.ChangeTypeId(self.raumBeschriftungID)
            except Exception as exc:
                logger.error(exc)  
        trans.Commit()
               
        self.Close()


beschriftungen('window.xaml',Beschriftungstype_dic,views).ShowDialog()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
