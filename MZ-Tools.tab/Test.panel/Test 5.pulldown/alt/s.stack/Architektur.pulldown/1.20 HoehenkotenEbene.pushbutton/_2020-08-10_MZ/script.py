# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr
import operator

start = time.time()


__title__ = "9.Ebene"
__doc__ = """Nullpunkt"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# Ebene aus aktueller Ansicht
Levels_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Levels)
levels = Levels_collector.ToElementIds()

logger.info("{} Ebenen ausgewählt".format(len(levels)))


if not levels:
    logger.error("Keine Ebene in aktueller Projekt gefunden")
    script.exit()


class ebenen:
    def __init__(self, element_id):
        """
        Definiert Ebene Klasse mit allen object properties.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['hoehe', 'Ansicht'],
            ['Geschoss', 'Gebäudegeschoss']
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.name, self.hoehe))


    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)

    def __get_value_in_project_units(self, param):
        """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""

        value = revit.query.get_param_value(param)

        try:
            unit = param.DisplayUnitType

            value = DB.UnitUtils.ConvertFromInternalUnits(
                value,
                unit)

        except Exception as e:
            pass

        return value


    def table_row(self):
        """ Gibt eine Datenreihe für den Ebenen aus. Für die tabellarische Übersicht."""
        return [
            self.name,
            self.Geschoss,
            self.hoehe,

        ]

    def __repr__(self):
        return "Ebene({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.name, self.hoehe)


table_data = []
Ebene_liste = []
with forms.ProgressBar(title='{value}/{max_value} Ebenen',
                       cancellable=True, step=10) as pb:

    for n, level_id in enumerate(levels):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(levels))
        ebene = ebenen(level_id)

        Ebene_liste.append(ebene)
        table_data.append(ebene.table_row())

output.print_table(
    table_data=table_data,
    title="Ebenen",
    columns=['Name', 'Geschoss', 'Höhe', ]
)
logger.info("{}".format(table_data))

table_Date = []

NullPunkt = 0
ProjektNullPunkt = 0
for EB in table_data:
    if EB[0] == 'NN':
        NullPunkt = EB[2]
    if EB[0] == 'EG OKFB':
        ProjektNullPunkt = EB[2]

for date in table_data:
    if date[0] and date[1]:
        Ueber_PN = date[2] - ProjektNullPunkt
        Ueber_NN = date[2] - NullPunkt
        date.append(Ueber_PN)
        date.append(Ueber_NN)
        del(date[1])
        table_Date.append(date)

table_Date.sort(key = operator.itemgetter(2))


output.print_table(
    table_data=table_Date,
    title="Ebenen",
    columns=['Name', 'Höhe', 'Über_PN', 'Über_NN']
)

# Parameters erstellen, Werte zuückschreiben + Abfrage
if forms.alert('Globale Parameter erstellen?', ok=False, yes=True, no=True):
    t = Transaction(doc, "Erstellen")
    t.Start()

    for Ebe in table_Date:
        if Ebe[3] == 0:
            new_Para = DB.GlobalParameter.Create(doc,'IGF_NN',DB.ParameterType.Number)
            nn = Ebe[2]
            new_Para.SetFormula(str(nn))
            new_Para.GetDefinition().set_ParameterGroup(BuiltInParameterGroup.PG_CONSTRAINTS)
            logger.info("{},  {}".format(new_Para.Name, Ebe[2]))
        if Ebe[2] != 0 and Ebe[3] != 0:
            name_PN = 'IGF_' + str(Ebe[0]) + '_ÜPN'
            new_Para_PN = DB.GlobalParameter.Create(doc,name_PN,DB.ParameterType.Number)
            new_Para_PN.GetDefinition().set_ParameterGroup(BuiltInParameterGroup.PG_CONSTRAINTS)
            new_Para_PN.SetFormula(str(Ebe[2]))
            name_NN = 'IGF_' + str(Ebe[0]) + '_ÜNN'
            new_Para_NN = DB.GlobalParameter.Create(doc,name_NN,DB.ParameterType.Number)
            new_Para_NN.GetDefinition().set_ParameterGroup(BuiltInParameterGroup.PG_CONSTRAINTS)
            new_Para_NN.SetFormula(str(Ebe[3]))

            logger.info("{},  {}".format(new_Para_PN.Name, Ebe[2]))
            logger.info("{},  {}".format(new_Para_NN.Name, Ebe[3]))

    t.Commit()


gl_Para_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.GlobalParameter))

gl_Para = gl_Para_collector.ToElementIds()
logger.info("{} globale Parameters erstellt".format(len(gl_Para)))

for g in gl_Para_collector:
    print(g.GroupId)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
