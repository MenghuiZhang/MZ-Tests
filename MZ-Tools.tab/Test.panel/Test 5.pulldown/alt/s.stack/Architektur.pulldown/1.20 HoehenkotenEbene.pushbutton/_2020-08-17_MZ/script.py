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

# Ebene aus aktueller Projekt
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
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
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
    title="Alle Ebenen",
    columns=['Name', 'Geschoss', 'Höhe', ]
)


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

# Sorteiren nach Höhe
table_Date.sort(key = operator.itemgetter(2))


output.print_table(
    table_data=table_Date,
    title="Ebenen mit Gebäudegeschoss",
    columns=['Name', 'Höhe', 'Über_PN', 'Über_NN']
)


# Globale Parameters aus aktureller Projekt
GlobaleParameter_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.GlobalParameter))

GlParameter = GlobaleParameter_collector.ToElementIds()
logger.info("{} globale Parameters in Projekt".format(len(GlParameter)))

Name = []
for g in GlobaleParameter_collector:
    Name.append(g.Name)

# Parameters erstellen, Werte zuückschreiben + Abfrage
if forms.alert('Globale Parameter erstellen?', ok=False, yes=True, no=True):
    t = Transaction(doc, "Erstellen")
    t.Start()

    i = 0
    j = 0

    with forms.ProgressBar(title='globale Parameter erstellen',
                           cancellable=True, step=1) as pb2:


        n_1 = 0
        for Ebe in table_Date:

            if pb2.cancelled:
                script.exit()
            n_1 += 1
            pb2.update_progress(n_1, len(table_Date))

            # IGF_NN erstellen
            if Ebe[3] == 0:
                # Document, Name und Datentyp
                pruefung = []
                for na in Name:
                    pruefung.append(na == 'IGF_Ebene_NN')
                if any(pruefung):
                    logger.info('Parameter ({}) bereits vorhanden'.format('IGF_Ebene_NN'))
                    if forms.alert('Neuen Wert schreiben?', ok=False, yes=True, no=True):
                        for Gl in GlobaleParameter_collector:
                            if Gl.Name == 'IGF_Ebene_NN':
                                Gl.SetFormula(str(Ebe[2]))
                                logger.info("{},  {}".format(Gl.Name, Ebe[2]))
                else:
                    new_Para = DB.GlobalParameter.Create(doc,'IGF_Ebene_NN',DB.ParameterType.Number)
                    # Formel und Wert
                    new_Para.SetFormula(str(Ebe[2]))
                    # Gruppe
                    new_Para.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS
                    logger.info("{},  {}".format(new_Para.Name, Ebe[2]))

                    i += 1


            # IGF_ÜNN und IGF_ÜPN erstellen
            if Ebe[2] != 0 and Ebe[3] != 0:
                name_PN = 'IGF_' + str(Ebe[0]) + '_ÜPN'
                name_NN = 'IGF_' + str(Ebe[0]) + '_ÜNN'
                pruefung1 = []
                pruefung2 = []
                for nam in Name:
                    pruefung1.append(nam == name_PN)
                    pruefung2.append(nam == name_NN)
                if any(pruefung1):
                    logger.info('Parameter ({}) bereits vorhanden'.format(name_PN))
                    if forms.alert('Neuen Wert schreiben?', ok=False, yes=True, no=True):
                        for Gl in GlobaleParameter_collector:
                            if Gl.Name == name_PN:
                                Gl.SetFormula(str(Ebe[2]))
                                logger.info("{},  {}".format(Gl.Name, Ebe[2]))
                if any(pruefung2):
                    logger.info('Parameter ({}) bereits vorhanden'.format(name_NN))
                    if forms.alert('Neuen Wert schreiben?', ok=False, yes=True, no=True):
                        for Gl in GlobaleParameter_collector:
                            if Gl.Name == name_NN:
                                Gl.SetFormula(str(Ebe[3]))
                                logger.info("{},  {}".format(Gl.Name, Ebe[3]))
                if not any(pruefung1) and not any(pruefung2):
                    new_Para_PN = DB.GlobalParameter.Create(doc,name_PN,DB.ParameterType.Number)
                    new_Para_PN.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS
                    new_Para_PN.SetFormula(str(Ebe[2]))
                    new_Para_NN = DB.GlobalParameter.Create(doc,name_NN,DB.ParameterType.Number)
                    new_Para_NN.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS
                    new_Para_NN.SetFormula(str(Ebe[3]))

                    logger.info("{},  {}".format(new_Para_PN.Name, Ebe[2]))
                    logger.info("{},  {}".format(new_Para_NN.Name, Ebe[3]))

                    j += 2


            else:
                pass



    t.Commit()
    logger.info("{} globale Parameters erstellt".format(i+j))

# Globale Parameters aus aktureller Projekt
GlobaleParameter_Collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.GlobalParameter))

GLParameter = GlobaleParameter_Collector.ToElementIds()
logger.info("{} globale Parameters ausgewählt".format(len(GLParameter)))



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
