import openpyxl
import openpyxl.utils

def get_data_rows(file_name):
    excel_file = openpyxl.load_workbook(file_name)
    sheet = excel_file.active
    return sheet.rows

def process_data(all_data):
    town_data = []
    for row in all_data:
        labor_force = row[1].value
        if type(labor_force) is int:
            town_name= row[0].value
            data_tuple = (town_name,labor_force)
            town_data.append(data_tuple)
    return town_data

def get_key(town_tuple):
    return town_tuple[1]

def main():
    all_data = get_data_rows("MAEmplyomentData.xlsx")
    town_data = process_data(all_data)
    town_data.sort(key=get_key)
    how_many = len(town_data)
    middle = how_many//2
    median_town = town_data[middle]
    print(f"The town with the median labor force is {median_town[0]} with a labor force of {median_town[1]} people")

if __name__ == '__main__':
    main()
