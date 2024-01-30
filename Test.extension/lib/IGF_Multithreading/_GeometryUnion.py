# coding: utf8
import Autodesk.Revit.DB as DB
from IGF_Multithreading import threading, Environment

class Geometry:
    def __init__(self, Liste):
        self.data = Liste
        self.Gesamt = []
        self.Export = None
    
    def reset(self,Liste):
        self.data = Liste
        self.Gesamt = []
        self.Export = None

    def process_data_in_parallel(self):
        self.Gesamt = []
        total_threads = Environment.ProcessorCount
        threads = []
        chunk_size = -(-len(self.data) // total_threads)
        nichterfolgreich = []
        def process_data_range(start, end):
            gesamt = None
            temp = None
            for i in range(start, end):
                if i == start:
                    gesamt = DB.BooleanOperationsUtils.ExecuteBooleanOperation(self.data[i],self.data[i],DB.BooleanOperationsType.Union)
                else:
                    try:
                        temp = gesamt
                        gesamt = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt,self.data[i],DB.BooleanOperationsType.Union)
                        temp.Dispose()
                    except:nichterfolgreich.append(self.data[i])
            self.Gesamt.append(gesamt)   

        for i in range(total_threads):
            start_index = i * chunk_size
            end_index = min(start_index + chunk_size, len(self.data))
            thread = threading.Thread(target=process_data_range, args=(start_index, end_index))
            thread.start()
            threads.append(thread)

        for thread in threads:
            thread.join()
        
        gesamt = None
        n = 0
        liste0  = self.Gesamt[:]
        liste = nichterfolgreich
        liste.extend(liste0)
        while(len(liste)> 0):
            n+=1
            if n > 10:
                break
            templiste = liste[:]
            liste = []
            for el in templiste:
                if not gesamt:
                    gesamt = el
                else:
                    try:
                        gesamt = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt,el,DB.BooleanOperationsType.Union)
                    except Exception as e:
                        liste.append(el)
        print(n)
        self.Export = gesamt


