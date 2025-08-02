import svg
from textwrap import dedent
import pandas

def drawOneCast(index, character_name, person_name, output_dir):
    image = svg.SVG(
        width=280,
        height=145,
        elements=[
            svg.Style(
                text=dedent("""
                    .number { 
                        font-family: Poppins; 
                        font-size: 220px;
                        font-weight: 1000;
                        text-anchor: middle;
                        dominant-baseline: middle;
                        line-height: .5;
                        fill: #E7E7E7;
                    }
                    .character {
                        font-family: Poppins; 
                        font-size: 18px;
                        font-weight: 1000;
                        text-anchor: start;
                        fill: #2E2E2E;
                    }
                    .name {
                        font-family: Poppins; 
                        font-size: 15px;
                        font-weight: 300;
                        text-anchor: start;
                        fill: #2E2E2E;
                    }
                """),
            ),
            svg.Text(x=140, dy="22px", y=72.5, class_=["number"], text=index),
            svg.Text(x=20, y=55, class_=["character"], text=character_name),
            svg.Text(x=20, y=100, class_=["name"], text=person_name),
            svg.Rect(
                x=0, y=70.5, width=50, height=.5,
                fill="transparent", stroke="#2E2E2E", stroke_width=.5),
        ],
    )

    with open(f"{output_dir}/mic_tag_{index}.svg", "w") as f:
        f.write(image.as_str())

def drawTwoCast(index, character_name, person_one_name, person_two_name, cast_one_name, cast_two_name, output_dir):
    image = svg.SVG(
        width=280,
        height=145,
        elements=[
            svg.Style(
                text=dedent("""
                    .number { 
                        font-family: Poppins; 
                        font-size: 220px;
                        font-weight: 1000;
                        text-anchor: middle;
                        dominant-baseline: middle;
                        line-height: .5;
                        fill: #E7E7E7;
                    }
                    .character {
                        font-family: Poppins; 
                        font-size: 18px;
                        font-weight: 1000;
                        text-anchor: start;
                        fill: #2E2E2E;
                    }
                    .name {
                        font-family: Poppins; 
                        font-size: 15px;
                        font-weight: 400;
                        text-anchor: start;
                        fill: #2E2E2E;
                    }
                    .cast {
                        font-family: Poppins; 
                        font-size: 8px;
                        font-weight: 700;
                        text-anchor: start;
                        fill: #7A7A7A;
                    }
                """),
            ),
            svg.Text(x=140, dy="22px", y=72.5, class_=["number"], text=index),
            svg.Text(x=20, y=60, class_=["character"], text=character_name),
            svg.Text(x=20, y=88, class_=["cast"], text=cast_one_name),
            svg.Text(x=20, y=103, class_=["name"], text=person_one_name),
            svg.Text(x=20, y=120, class_=["cast"], text=cast_two_name),
            svg.Text(x=20, y=135, class_=["name"], text=person_two_name),
            svg.Rect(
                x=0, y=70.5, width=50, height=.5,
                fill="transparent", stroke="#2E2E2E", stroke_width=.5),
        ],
    )
    
    with open(f"{output_dir}/mic_tag_{index}.svg", "w") as f:
        f.write(image.as_str())
    
import os

def locate_xlsx_files():
    return [file for file in os.listdir(".") if file.endswith(".xlsx")]

def get_user_input(prompt, default, validation_func=None):
    while True:
        user_input = input(prompt)
        if validation_func is None or validation_func(user_input) == True:
            return user_input
        elif user_input == "":
            return default
        else:
            print("Invalid input. Please try again.")

def validate_yes_no(input_str):
    if input_str.lower() in ["y", "yes"]:
        return True
    elif input_str.lower() in ["n", "no"]:
        return True
    else:
        return None 
    
def validate_int_input(input_str):
  try:
      int(input_str)
      return True
  except ValueError:
      return None

def main():
    xlsx_files = locate_xlsx_files()

    if not xlsx_files:
        print("No xlsx files found in the current directory.")
        return

    if len(xlsx_files) > 1:
        for i, xlsx_file in enumerate(xlsx_files):
            print(f"{i + 1}. {xlsx_file}")

        file_index = int(get_user_input("Choose a file to generate snippets from (1): ", "1", validate_int_input)) -1

        if 0 <= file_index < len(xlsx_files):
            xlsx_file = xlsx_files[file_index]
        else:
            print("Invalid file selection.")
            return
    else:
        xlsx_file = xlsx_files[0]

    show_file_name = xlsx_file.replace(".xlsx", "")
    xls = pandas.ExcelFile(xlsx_file)
    
    print(f"Using file: {xlsx_file} ({show_file_name})")
    print(f"Sheet names: {xls.sheet_names}")
    sheet_name = get_user_input(
        f"Select a sheet name (default: {xls.sheet_names[0]}): ",
        xls.sheet_names[0],
        lambda x: x in xls.sheet_names
    )
    
    skip_rows = int(get_user_input("Start at row (1): ", "1", validate_int_input)) - 1
    if skip_rows < 0:
        skip_rows = 0
        
    one_or_two_cast = get_user_input(
        "Is this a one or two cast sheet? (1/2, default: 1): ",
        "1",
        lambda x: x in ["1", "2"]
    )

    if one_or_two_cast == "1":
        print("Generating tags for one cast...")
    else:
        print("Generating tags for two casts...")
    
    data_frame = read_excel_data(xlsx_file, sheet_name, skiprows=skip_rows)
    
    output_dir = "output"
    if os.path.exists(output_dir):
        for file in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, file))
    else:
        os.makedirs(output_dir)

    generate_tags(data_frame, output_dir, one_or_two_cast)

def generate_tags(data_frame, output_dir, one_or_two_cast):
    mic_index = 1

    total_tags = len(data_frame)
    print(f"Creating {total_tags} tags...")
    
    if one_or_two_cast == "2":
        cast_one_name = data_frame.columns[2]
        cast_two_name = data_frame.columns[3]
        print(f"Cast one: {cast_one_name}, Cast two: {cast_two_name}")

    for _, row in data_frame.iterrows():
        if one_or_two_cast == "2":
            character_name = row.values[1]
            person_one_name = row.values[2]
            person_two_name = row.values[3]
            
            if person_one_name == person_two_name:
                drawOneCast(mic_index, character_name, person_one_name, output_dir)
            else:
                drawTwoCast(mic_index, character_name, person_one_name, person_two_name, cast_one_name, cast_two_name, output_dir)
        else:
            character_name = row[1]
            person_name = row[2]
            drawOneCast(mic_index, character_name, person_name, output_dir)

        mic_index += 1

    print(f"\nCreated {total_tags} tags.")

def read_excel_data(xlsx_file, sheet_name, skiprows=0):
    try:
        return pandas.read_excel(xlsx_file, engine="openpyxl", sheet_name=sheet_name, skiprows=skiprows)
    except FileNotFoundError:
        print(f"Error: File '{xlsx_file}' not found.")
        exit()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        exit()
        
if __name__ == '__main__':
    main()