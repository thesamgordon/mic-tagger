import os

import pandas


def draw_one_cast(index, character_name, person_name, output_dir, title, file_name):
    with open(f"{file_name}.svg", "r") as f:
        template = f.read()

    image = template.replace("{NUMBER}", str(index))
    image = image.replace("{CHARACTER}", character_name)
    image = image.replace("{PERSON}", person_name)
    image = image.replace("{TITLE}", title)

    with open(f"{output_dir}/{file_name}_{index}.svg", "w") as f:
        f.write(image)


def draw_two_casts(
    index,
    character_name,
    person_one_name,
    person_two_name,
    cast_one_name,
    cast_two_name,
    output_dir,
    title,
    file_name,
):
    with open(f"{file_name}.svg", "r") as f:
        template = f.read()

    image = template.replace("{CAST_ONE}", cast_one_name)
    image = image.replace("{CAST_TWO}", cast_two_name)
    image = image.replace("{NUMBER}", str(index))
    image = image.replace("{CHARACTER}", character_name)
    image = image.replace("{PERSON_ONE}", person_one_name)
    image = image.replace("{PERSON_TWO}", person_two_name)
    image = image.replace("{TITLE}", title)

    with open(f"{output_dir}/{file_name}_{index}.svg", "w") as f:
        f.write(image)


def locate_xlsx_files():
    return [file for file in os.listdir(".") if file.endswith(".xlsx")]


def get_user_input(prompt, default, validation_func=None):
    while True:
        user_input = input(prompt)
        if validation_func is None or validation_func(user_input) is True:
            return user_input
        elif user_input == "":
            return default
        else:
            print("Invalid input. Please try again.")


def validate_int_input(input_str):
    try:
        int(input_str)
        return True
    except ValueError:
        return None


def read_excel_data(xlsx_file, sheet_name, skiprows=0):
    try:
        return pandas.read_excel(
            xlsx_file, engine="openpyxl", sheet_name=sheet_name, skiprows=skiprows
        )
    except FileNotFoundError:
        print(f"Error: File '{xlsx_file}' not found.")
        exit()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        exit()


def choose_xlsx_file(xlsx_files):
    if len(xlsx_files) == 1:
        return xlsx_files[0]

    for i, xlsx_file in enumerate(xlsx_files):
        print(f"{i + 1}. {xlsx_file}")

    file_index = (
        int(
            get_user_input(
                "Choose a file to generate snippets from (1): ",
                "1",
                validate_int_input,
            )
        )
        - 1
    )

    if 0 <= file_index < len(xlsx_files):
        return xlsx_files[file_index]

    print("Invalid file selection.")
    return None


def prepare_output_dir(output_dir):
    if os.path.exists(output_dir):
        for file in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, file))
    else:
        os.makedirs(output_dir)


def generate_tags(data_frame, output_dir, one_or_two_cast, title, file_name):
    mic_index = 1
    total_tags = len(data_frame)
    print(f"Creating {total_tags} tags...")

    cast_one_name = None
    cast_two_name = None

    if one_or_two_cast == "2":
        cast_one_name = data_frame.columns[2].upper()
        cast_two_name = data_frame.columns[3].upper()

        if cast_one_name is None or type(cast_one_name) is not str:
            cast_one_name = "CAST 1"
        if cast_two_name is None or type(cast_two_name) is not str:
            cast_two_name = "CAST 2"

        print(f"Cast one: {cast_one_name}, Cast two: {cast_two_name}")

    for _, row in data_frame.iterrows():
        if one_or_two_cast == "2":
            character_name = row.values[1]
            person_one_name = row.values[2]
            person_two_name = row.values[3]

            if person_two_name is None or type(person_two_name) is not str:
                person_two_name = person_one_name

            if person_one_name == person_two_name:
                draw_one_cast(
                    mic_index,
                    character_name,
                    person_one_name,
                    output_dir,
                    title,
                    "one_cast",
                )
            else:
                draw_two_casts(
                    mic_index,
                    character_name,
                    person_one_name,
                    person_two_name,
                    cast_one_name,
                    cast_two_name,
                    output_dir,
                    title,
                    "two_cast",
                )
        else:
            character_name = row.values[1]
            person_name = row.values[2]

            draw_one_cast(
                mic_index, character_name, person_name, output_dir, title, "one_cast"
            )
            draw_one_cast(
                mic_index,
                character_name,
                person_name,
                output_dir,
                title,
                "one_cast_thin",
            )

        mic_index += 1

    print(f"\nCreated {total_tags} tags.")

    with open("index.html", "r") as f:
        template = f.read()

    image_tags = ""
    for i in range(1, mic_index):
        image_tags += f'      <img src="{file_name}_{i}.svg" />\n'

    for i in range(1, mic_index):
        if os.path.exists(f"{output_dir}/{file_name}_thin_{i}.svg"):
            image_tags += f'      <img src="{file_name}_thin_{i}.svg" />\n'

    template = template.replace("{DATA}", image_tags)
    with open(f"{output_dir}/index.html", "w") as f:
        f.write(template)


def main():
    xlsx_files = locate_xlsx_files()
    if not xlsx_files:
        print("No xlsx files found in the current directory.")
        return

    xlsx_file = choose_xlsx_file(xlsx_files)
    if xlsx_file is None:
        return

    show_file_name = xlsx_file.replace(".xlsx", "")
    xls = pandas.ExcelFile(xlsx_file)

    print(f"Using file: {xlsx_file} ({show_file_name})")
    print(f"Sheet names: {xls.sheet_names}")
    sheet_name = get_user_input(
        f"Select a sheet name (default: {xls.sheet_names[0]}): ",
        xls.sheet_names[0],
        lambda x: x in xls.sheet_names,
    )

    skip_rows = int(get_user_input("Start at row (1): ", "1", validate_int_input)) - 1
    if skip_rows < 0:
        skip_rows = 0

    one_or_two_cast = get_user_input(
        "Is this a one or two cast sheet? (1/2, default: 1): ",
        "1",
        lambda x: x in ["1", "2"],
    )

    title = get_user_input("Enter a title for the tags: ", "")

    if one_or_two_cast == "1":
        print("Generating tags for one cast...")
    else:
        print("Generating tags for two casts...")

    data_frame = read_excel_data(xlsx_file, sheet_name, skiprows=skip_rows)

    output_dir = "output"
    prepare_output_dir(output_dir)

    generate_tags(
        data_frame,
        output_dir,
        one_or_two_cast,
        title,
        "two_cast" if one_or_two_cast == "2" else "one_cast",
    )


if __name__ == "__main__":
    main()
