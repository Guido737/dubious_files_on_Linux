import xlsxwriter

def parse_line(line):
    try:
        parts = line.strip().split()
        if len(parts) >= 4:
            name = ' '.join(parts[:-3])
            profession = parts[-3]  
            age = int(parts[-2]) 
            gender = parts[-1]  
            return name, profession, age, gender
    except IndexError as e: 
        print(f"Error parsing line (IndexError): {line}. Error: {e}")
    except ValueError as e: 
        print(f"Error parsing line (ValueError): {line}. Error: {e}")
    return None

def get_profession_color(profession):
    profession_colors = {
        'программист': "green",
        'певец': "yellow",
        'музыкант': "yellow",
        'писатель': "red",
        'горняк': "brown",
        'стоматолог': "blue",
        'официантка': "pink",
        'продавец': "tomato"
    }
    
    try:
        profession = profession.lower()
        for key, color in profession_colors.items():
            if key in profession:
                return color
    except AttributeError as e:
        print(f"Error processing profession (AttributeError): {profession}. Error: {e}")
    return "gray"

def read_data(input_filename):
    people = []
    try:
        with open(input_filename, 'r', encoding='utf-8') as f:
            for line in f:
                parsed_data = parse_line(line.strip())
                if parsed_data:
                    people.append(parsed_data)
    except FileNotFoundError as e: 
        print(f"Error: The file '{input_filename}' was not found. Error: {e}")
    except IOError as e:
        print(f"Error reading file '{input_filename}'. Error: {e}")
    return people

def create_workbook(output_filename):
    try:
        workbook = xlsxwriter.Workbook(output_filename)
        worksheet = workbook.add_worksheet()
        color_formats = {
            "green": workbook.add_format({'bg_color': 'green', 'font_color': 'white'}),
            "yellow": workbook.add_format({'bg_color': 'yellow', 'font_color': 'white'}),
            "red": workbook.add_format({'bg_color': 'red', 'font_color': 'white'}),
            "brown": workbook.add_format({'bg_color': 'brown', 'font_color': 'white'}),
            "blue": workbook.add_format({'bg_color': 'blue', 'font_color': 'white'}),
            "pink": workbook.add_format({'bg_color': 'pink', 'font_color': 'white'}),
            "tomato": workbook.add_format({'bg_color': 'tomato', 'font_color': 'white'}),
            "gray": workbook.add_format({'bg_color': 'gray', 'font_color': 'white'}),
        }
        return workbook, worksheet, color_formats
    except IOError as e: 
        print(f"Error creating workbook '{output_filename}'. Error: {e}")
        return None, None, None

def write_data_to_excel(worksheet, color_formats, people):
    try:
        worksheet.write_row(0, 0, ["Full Name", "Profession", "Age", "Gender"])  
        row = 1
        for person in people:
            name, profession, age, gender = person
            color = get_profession_color(profession)
            format_for_row = color_formats.get(color, color_formats["gray"]) 
            worksheet.write(row, 0, name, format_for_row)  
            worksheet.write(row, 1, profession, format_for_row) 
            worksheet.write(row, 2, age, format_for_row)  
            worksheet.write(row, 3, gender, format_for_row)  
            row += 1
            
    except ValueError as e:  
        print(f"Error writing data to Excel (ValueError). Error: {e}")
    except IOError as e: 
        print(f"Error writing to Excel file. Error: {e}")
    except Exception as e: 
        print(f"Unexpected error while writing to Excel. Error: {e}")

def process_file(input_filename, output_filename):
    people = read_data(input_filename)  
    if not people:
        print("No valid data to process.")
        return
    try:
        people.sort(key=lambda x: x[1]) 
    except TypeError as e: 
        print(f"Error sorting the data (TypeError). Error: {e}")
    except Exception as e:
        print(f"Error sorting the data. Error: {e}")
    
    workbook, worksheet, color_formats = create_workbook(output_filename)  
    if workbook is None:
        print("Workbook creation failed.")
        return
    
    write_data_to_excel(worksheet, color_formats, people)
    try:
        workbook.close() 
        print(f"Data successfully saved to file: {output_filename}")
    except IOError as e: 
        print(f"Error closing the workbook or saving the file (IOError). Error: {e}")
    except Exception as e:
        print(f"Error closing the workbook or saving the file. Error: {e}")
def main():
    input_filename = 'Enter path to file' 
    output_filename = 'Enter path to save file'  
    process_file(input_filename, output_filename) 

if __name__ == "__main__":
    main()
