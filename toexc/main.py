import argparse
import xlsxwriter as xw

parser = argparse.ArgumentParser()

parser.add_argument("-f", "--file")
parser.add_argument("-x", "--excel")

args = parser.parse_args()

def get_data_from_db(fname):
    with open(fname) as f:
        return f.readlines()

def sort_data(data):
    md = {}
    for i in range(len(data)):
        d = data[i].strip().split(';')
        tmp = {}
        tmp['name'] = d[0]
        tmp['surname'] = d[1]
        tmp['age'] = d[2]
        tmp['profession'] = d[3]   
        md[i] = tmp     
    return md

def add_to_excel(data, fname):
    workbook = xw.Workbook(fname)
    ws = workbook.add_worksheet(name='result') 
    format1 = workbook.add_format({'bg_color':'#FAF603', 'border':1})
    ws.write(0, 0, 'Name', format1)
    ws.write(0, 1, 'Surname', format1)
    ws.write(0, 2, 'Age', format1)
    ws.write(0, 3, 'Profession', format1)
    for i in range(len(data)):
        format2 = workbook.add_format({'bg_color':'#CAC3C1', 'border':1})
        if data[i]['profession'] == "Programmer":
            format2 = workbook.add_format({'bg_color':'#D92A08', 'border':1})
        ws.write(i+1, 0, data[i]['name'], format2)
        ws.write(i+1, 1, data[i]['surname'], format2)
        ws.write(i+1, 2, data[i]['age'], format2)
        ws.write(i+1, 3, data[i]['profession'], format2)
        
    workbook.close()
    
def main():
    db = args.file;
    exc = args.excel;
    data = get_data_from_db(db)
    sorted_data = sort_data(data)
    add_to_excel(sorted_data, exc)
    
    
if __name__ == "__main__":
    main()