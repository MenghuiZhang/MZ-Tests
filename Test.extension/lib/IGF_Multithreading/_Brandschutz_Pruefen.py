# coding: utf8
import Autodesk.Revit.DB as DB
from IGF_Multithreading import threading, Environment

class BrandschutzPruefen:
    def __init__(self, Liste_ZuPruefen,Liste_Brandschutz,doc):
        self.Liste_ZuPruefen = Liste_ZuPruefen
        self.Liste_Brandschutz = Liste_Brandschutz
        self.doc = doc
        self._Dict = {}
    
    def reset(self,Liste_ZuPruefen,Liste_Brandschutz,doc):
        self.Liste_ZuPruefen = Liste_ZuPruefen
        self.Liste_Brandschutz = Liste_Brandschutz
        self.doc = doc
        self._Dict = {}

    def process_data_in_parallel(self):
        self._Dict = {}
        total_threads = Environment.ProcessorCount
        threads = []
        chunk_size = -(-len(self.Liste_ZuPruefen) // total_threads)
        nichterfolgreich = []
        def process_data_range(start, end):
            gesamt = None
            temp = None
            for i in range(start, end):
                elem = self.Liste_ZuPruefen[i]
                Filter = DB.ElementIntersectsElementFilter(elem)
                coll = DB.FilteredElementCollector(self.doc,self.Liste_Brandschutz).WherePasses(Filter).ToElementIds()
                Filter.Dispose()
                if coll.Count > 0:
                    self._Dict[elem.Id.ToString()] = coll

        for i in range(total_threads):
            start_index = i * chunk_size
            end_index = min(start_index + chunk_size, len(self.Liste_ZuPruefen))
            thread = threading.Thread(target=process_data_range, args=(start_index, end_index))
            thread.start()
            threads.append(thread)

        for thread in threads:
            thread.join()
    
    def process_data_einzehln(self):
        self._Dict = {}
        for i in range(len(self.Liste_ZuPruefen)):
            elem = self.Liste_ZuPruefen[i]
            Filter = DB.ElementIntersectsElementFilter(elem)
            coll = DB.FilteredElementCollector(self.doc,self.Liste_Brandschutz).WherePasses(Filter).ToElementIds()
            Filter.Dispose()
            if coll.Count > 0:
                self._Dict[elem.Id.ToString()] = coll

        