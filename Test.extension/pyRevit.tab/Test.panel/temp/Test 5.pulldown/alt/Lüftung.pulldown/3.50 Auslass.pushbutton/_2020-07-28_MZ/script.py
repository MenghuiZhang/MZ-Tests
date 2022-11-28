# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time


start = time.time()


__title__ = "8.Luftauslass"
__doc__ = """Informationen der Luftauslaesse zeigen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal)
ducts = spaces_collector.ToElementIds()

phase = list(doc.Phases)[-1]

logger.info("{} Luftauslässe ausgewählt".format(len(ducts)))

if not ducts:
    logger.error("Keine Luftauslässe in aktueller Ansicht gefunden")
    script.exit()


class DUCT_Terminal:
    Auslass_Liste = []
    Auslass_Tag = []
    Raumnummer = []

    def __init__(self, element_id):
        """
        Definiert Luftauslässe Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr_Auslass = [
            ['S_ABZ', 'Systemabkürzung'],
            ['S_Name', 'Systemname'],
            ['S_Grupp', 'Systemklassifizierung'],
            ['KZ', 'Kennzeichen'],
            ['Vmin', 'IGF_RLT_AuslassVolumenstromMin'],
            ['Vmax', 'IGF_RLT_AuslassVolumenstromMax'],
            ['V_Nacht', 'IGF_RLT_AuslassVolumenstromNacht'],
#            ['V_TNacht', 'TGA_RLT_AbluftSumme24h'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['zuluft_min', 'TGA_RLT_ZuluftminRaum'],
            ['abluft_min', 'TGA_RLT_AbluftminRaum'],
            ['abluft_labor', 'TGA_RLT_AbluftSummeLabor'],
            ['abluft_24h', 'TGA_RLT_AbluftSumme24h'],
            ['TS_Zuluft', 'Tatsächlicher Zuluftstrom'],
            ['TS_Abluft', 'Tatsächlicher Abluftstrom'],
        ]

        for a in attr_Auslass:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr_Auslass(revit_name))

        for b in attr_Raum:
            python_name, revit_name = b
            setattr(self, python_name, self.__get_element_attr_Raum(revit_name))

        logger.info(50 * "=")
        logger.info("{}\t{}\t{}\t{}".format(self.nummer, self.name, self.S_Name, self.S_ABZ))

        self.Raumnummer.append(self.nummer)

        self.Auslass_Tag = [self.nummer, self.name, self.zuluft_min, self.abluft_min,
                      self.TS_Zuluft, self.TS_Abluft,self.abluft_labor,
                      self.abluft_24h, self.S_ABZ, self.S_Name, self.S_Grupp,
                      self.KZ, self.Vmin, self.Vmax, self.V_Nacht]

        self.Auslass_Liste.append(self.Auslass_Tag)


    def __get_element_attr_Auslass(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)

    def __get_element_attr_Raum(self, param_name):
        if self.element.Space[phase]:
            param = self.element.Space[phase].LookupParameter(param_name)

        else:
            param = None

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
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.zuluft_min,
            self.abluft_min,
            self.abluft_labor,
            self.abluft_24h,
            self.S_Name,
            self.S_ABZ,
            self.S_Grupp,
            self.KZ,
            self.Vmin,
            self.Vmax,
            self.V_Nacht,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


table_data = []
Luftauslass_Liste = []

with forms.ProgressBar(title='{value}/{max_value} Luftmengenberechnung',
                       cancellable=True, step=10) as pb:

    for n, duct_id in enumerate(ducts):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(ducts))
        Auslass = DUCT_Terminal(duct_id)

        Luftauslass_Liste.append(Auslass)
        table_data.append(Auslass.table_row())

Raum_Nr = set(DUCT_Terminal.Raumnummer)
Raumnummer_Liste = list(Raum_Nr)

if not all(Raumnummer_Liste):
    Raumnummer_Liste.remove(None)

Raumnummer_Liste.sort()

AL_Liste = []
for RN in Raumnummer_Liste:
    for Al in DUCT_Terminal.Auslass_Liste:
        if Al[0] == RN:
            AL_Liste.append(Al)




output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name',
             'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', ]
)


output.print_table(
    table_data=AL_Liste,
    title="Luftmengen",
    columns=['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
             'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', ]
)

logger.info("{} Luftauslässe in Luftauslass_Liste".format(len(Luftauslass_Liste)))
logger.info("{} MEPRaum in Luftauslass_Liste".format(DUCT_Terminal.Raumnummer))
logger.info("{} MEPRaum in Luftauslass_Liste".format(set(DUCT_Terminal.Raumnummer)))
logger.info("{} MEPRaum in Luftauslass_Liste".format(Raumnummer_Liste))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
