# coding: utf8

import threading
import clr
clr.AddReference("System")
from System import Environment

class ItemTemplate:
    def __init__(self, data, method):
        self.data = data
        self.method = method

    def process_data_in_parallel(self):
        total_threads = Environment.ProcessorCount
        threads = []
        chunk_size = -(-len(self.data) // total_threads)
        def process_data_range(start, end):
            for i in range(start, end):
                self.method(self.data[i])

        for i in range(total_threads):
            start_index = i * chunk_size
            end_index = min(start_index + chunk_size, len(self.data))
            thread = threading.Thread(target=process_data_range, args=(start_index, end_index))
            thread.start()
            threads.append(thread)

        for thread in threads:
            thread.join()