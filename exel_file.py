import xlsxwriter

def parse_line(line):
    parts = line.strip().split()
    if len(parts) >= 4:
        name = ' '.join(parts[:-3])  
        profession = parts[-3]
        try:
            age = int(parts[-2])  
            gender = parts[-1]  
            return name, profession, age, gender
        except ValueError:
            return None
    return None

def get_profession_color(profession):
    if 'программист' in profession.lower():
        return "green"
    elif 'певец' in profession.lower() or 'музыкант' in profession.lower():
        return "yellow"
    elif 'писатель' in profession.lower():
        return "red"
    elif 'шахтер' in profession.lower():
        return "brown"
    elif 'дантист' in profession.lower():
        return "blue"
    elif 'официантка' in profession.lower():
        return "pink"
    elif 'продовщица' in profession.lower():
        return "tomato"
    else:
        return "gray"  

def process_file(input_filename, output_filename):
    people = []

    with open(input_filename, 'r', encoding='utf-8') as f:
        for line in f:
            parsed_data = parse_line(line.strip())
            if parsed_data:
                people.append(parsed_data)

    
    people.sort(key=lambda x: x[1])

    
    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet()

   
    color_formats = {
        "green": workbook.add_format({'bg_color': 'green', 'font_color': 'white'}),
        "yellow": workbook.add_format({'bg_color': 'yellow', 'font_color': 'black'}),
        "red": workbook.add_format({'bg_color': 'red', 'font_color': 'white'}),
        "brown": workbook.add_format({'bg_color': 'brown', 'font_color': 'white'}),
        "blue": workbook.add_format({'bg_color': 'blue', 'font_color': 'white'}),
        "pink": workbook.add_format({'bg_color': 'pink', 'font_color': 'black'}),
        "tomato": workbook.add_format({'bg_color': 'tomato', 'font_color': 'white'}),
        "gray": workbook.add_format({'bg_color': 'gray', 'font_color': 'white'}),
    }

    worksheet.write_row(0, 0, ["ФИО", "Профессия", "Возраст", "Пол", "Цвет"])

    row = 1
    for person in people:
        name, profession, age, gender = person
        color = get_profession_color(profession)
        worksheet.write(row, 0, name)
        worksheet.write(row, 1, profession)
        worksheet.write(row, 2, age)
        worksheet.write(row, 3, gender)
        worksheet.write(row, 4, color, color_formats[color])
        row += 1
    workbook.close()
    print(f"Данные успешно сохранены в файл: {output_filename}")
    
input_filename = '/home/usernamezero00/Desktop/people.txt'  
output_filename = '/home/usernamezero00/Desktop/people_output.xlsx'
process_file(input_filename, output_filename)

