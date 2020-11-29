from glob import glob
from itertools import takewhile, repeat
from tqdm import tqdm
import openpyxl
import os

forbidden_char_filename = {
    # src https://stackoverflow.com/questions/1976007/what-characters-are-forbidden-in-windows-and-linux-directory-names
    '/' : chr(0x2215),
    '\\' : chr(0x29F5),
    ':' : chr(0xA789),
    '*' : chr(0x2217),
    '?' : chr(0xFE56),
    '"' : chr(0x201C),
    '<' : chr(0xFE64),
    '>' : chr(0xFE65),
    '|' : chr(0x01C0)
}

class Parser:
    IDLE = 0
    INITIALIZE = 1
    COLUMN = 2
    READING = 3
    FINISHED = 4
    
    def __init__(self):
        self.init()
        self.map = {
            Parser.INITIALIZE : self.__set_table_name,
            Parser.COLUMN : self.__set_columns,
            Parser.READING : self.__process_data
        }
    
    def init(self):
        self.current_state = Parser.IDLE
        self.column_called = False #Some tables don't have proper column
    
    def process_line(self, line:str):
        if(line.startswith('+-')):
            self.current_state += 1
            return
        if(not line.startswith(';')):
            self.init()
            return
        func = self.map.get(self.current_state)
        if(func):
            func(line)
        
    
    def __set_table_name(self,line:str):
        self.table_name = line.strip().strip(';').strip()
    
    def __set_columns(self, line:str):
        if(self.column_called):
            self.__no_proper_column(line)
            return
        __columns = line.strip().strip(';').split(';')
        _columns = dict()
        length = []
        for i in __columns:
            c_name = i.strip()
            _columns[c_name]= []
            length.append(len(c_name))
        
        self.columns = _columns
        self.length = length
        self.column_called = True
    
    def __no_proper_column(self,line):
        self.current_state = Parser.READING
        j = 1
        __columns = dict()
        for i in self.columns:
            __columns['_' * j] = [i]
            j += 1
        self.columns = __columns
        self.__process_data(line)
    
    def __process_data(self, line:str):
        data = line.strip().strip(';').split(';')
        i = 0
        for c,d in zip(self.columns,data):
            value = d.strip()
            self.columns[c].append(value)
            self.length[i] = max(self.length[i],len(value))
            i += 1

    
    def get_current_state(self):
        return self.current_state
    
    def get_json(self):
        ret = dict()
        ret['table_name'] = self.table_name
        ret['data'] = self.columns
        ret['length'] = self.length
        
        self.init()
        return ret

def replace_forbidden(name:str):
    for i in forbidden_char_filename:
        name = name.replace(i,forbidden_char_filename[i])
    return name

def save_xl(json, save_path):
    file_name = replace_forbidden(json['table_name']) + '.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    for i, column in enumerate(json['data'], start=1):
        ws_column = sheet.cell(row = 1, column = i).column_letter
        sheet.column_dimensions[ws_column].width = (json['length'][i-1] + 4) * 1.2
        if(column.replace('_','')):
            cell = sheet.cell(row = 1, column = i)
            cell.font = openpyxl.styles.Font(name='Courier New',bold=True)
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            cell.value = column
            col_start = 2
        else:
            col_start = 1
        
        for j, value in enumerate(json['data'][column],start=col_start):
            cell = sheet.cell(row = j, column = i)
            cell.font = openpyxl.styles.Font(name='Courier New')
            cell.value = value
    
    file_path = os.path.join(save_path,file_name)
    i=1
    while(os.path.exists(file_path)):
        file_path = os.path.join(save_path,file_name.split('.')[-2] + '-' + str(i)+'.xlsx')
        i += 1
    
    workbook.save(file_path)
    workbook.close()

def get_line_count(path):
    # Copied from: https://stackoverflow.com/questions/845058/how-to-get-line-count-of-a-large-file-cheaply-in-python
    # Answer by: Michael Bacon ( https://stackoverflow.com/users/4367773/michael-bacon )

    with open(path, 'rb') as f:
        bufgen = takewhile(lambda x: x, (f.raw.read(1024*1024) for _ in repeat(None)))
        return sum(buf.count(b'\n') for buf in bufgen)

def main():
    reports = [i for i in glob('output_files/*.rpt')]
    if(not reports):
        print('No reports found. Make sure you have compiled using Quartus Prime and you are on correct diretory.')
        return

    if(not os.path.exists('ExcellReport')):
        os.mkdir('ExcellReport')

    parser = Parser()

    for r in reports:
        parser.init()
        total_line = get_line_count(r)

        file = open(r, 'r', encoding='utf-8',errors='backslashreplace')
        try:
            report_name = file.readline().split(' report ')[0]
        except:
            report_name = os.path.basename(r)
        print('Genarating Excell Report for', report_name)

        save_path = os.path.join('ExcellReport',report_name)
        if(not os.path.exists(save_path)):
            os.mkdir(save_path)
        
        for line in tqdm(file,total=total_line-1, unit='line'):
            parser.process_line(line)
            if(parser.get_current_state()==Parser.FINISHED):
                json = parser.get_json()
                save_xl(json, save_path)
                
        file.close()


if __name__ == '__main__':
    main()
