import xlwings as xw
import pandas as pd

__memory = {}
__callers = {}
__memory_pointer = 2**24-1

@xw.func
def create_df(caller, range):
    # define key
    key = f'[{caller.sheet.book.name}]{caller.sheet.name}!{caller.address}'
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