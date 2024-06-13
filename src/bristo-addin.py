import xlwings as xw
import pandas as pd

__memory = {}
__callers = {}
__memory_pointer = 2**32-1
_gc_count = 0

__cell_infos = {}
__cell_infos_memory_address = {}

class Cell_Info:
    def __init__(self, caller) -> None:
        self.address = caller.address
        self.sheet_name = caller.sheet.name
        self.book_name = caller.sheet.book.name

    def __key(self):
        return (self.book_name, self.sheet_name, self.address)

    def __hash__(self) -> int:
        return hash(self.__key())
    
    def __eq__(self, value: object) -> bool:
        if isinstance(value, Cell_Info):
            return self.__key() == value.__key()
        return NotImplemented


def gc(caller):
    global _gc_count
    _gc_count+=1
    if _gc_count%10 == 0:
        for key in __memory.keys:
            tokens = key.split('|')
            book = caller.sheet.book
            if book.name == tokens[0]:
                if book.sheets[tokens[1]].range(tokens[2]) == "":
                    del __memory[__callers[key]]


@xw.func
def create_df(caller, range):
    # define key
    key = f'{caller.sheet.book.name}|{caller.sheet.name}|{caller.address}'
    print(key)
    data = pd.DataFrame(columns=range[0], data=range[1:])

    # check if key is already present an delete the old reference 
    if key in __callers:
        old_data = __memory[__callers[key]]
        if data.equals(old_data):
            return __callers[key]
        del __memory[__callers[key]]
    
    global __memory_pointer
    new_memory = hex(__memory_pointer)
    __memory_pointer-=1
    __callers[key] = new_memory
    __memory[new_memory] = data
    return new_memory

@xw.func
def reveal_df(memory):
    if memory not in __memory:
        raise Exception('object not found in memory!')
    return __memory[memory]

@xw.func
def get_caller_info(caller):
    app = caller.sheet.book.app 
    key = f'[{caller.sheet.book.name}]{caller.sheet.name}!{caller.address};ExcelVersion={app.version}'
    return key


@xw.func
def create_new_df(caller, range):
    # define key
    cell_info = Cell_Info(caller=caller)
    print(cell_info)
    data = pd.DataFrame(columns=range[0], data=range[1:])

    # check if key is already present an delete the old reference 
    if cell_info in __cell_infos:
        old_data = __cell_infos_memory_address[__cell_infos[cell_info]]
        if data.equals(old_data):
            return __cell_infos[cell_info]
        del __cell_infos_memory_address[__cell_infos[cell_info]]
    
    global __memory_pointer
    new_memory = hex(__memory_pointer)
    __memory_pointer-=1
    __cell_infos[cell_info] = new_memory
    __cell_infos_memory_address[new_memory] = data
    return new_memory

@xw.func
def reveal_new_df(memory):
    if memory not in __cell_infos_memory_address:
        raise Exception('object not found in memory!')
    return __cell_infos_memory_address[memory]

@xw.func
def debug_cache(caller, purge):
    tt = []
    leaks = []
    book = caller.sheet.book
    if purge:
        for key in __cell_infos:
            if key.book_name == book.name:
                value = book.sheets[key.sheet_name][key.address].value
                if value not in __cell_infos_memory_address:
                    tt.append(f'[{key.sheet_name}]!{key.address}={book.sheets[key.sheet_name][key.address].value}')
                    leaks.append(key)
        for leak in leaks:
            memory = __cell_infos[leak]
            del __cell_infos_memory_address[memory]
            del __cell_infos[leak]
    else:
        for key in __cell_infos:
            if key.book_name == book.name:
                tt.append(f'[{key.sheet_name}]!{key.address}={book.sheets[key.sheet_name][key.address].value}')
    return ';'.join(tt)