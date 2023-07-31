import csv
from docx import Document


def find_loc_name(loc_id, locations):
    while loc_id in locations:
        loc_data = locations[loc_id]
        loc_name = loc_data["loc_name"]
        granularity = int(loc_data["granularity"])
        loc_type = loc_data["type"]

        if granularity == 1:
            if loc_type != "fac_area":
                parent_id = loc_data["parent_id"]
                parent_loc_name = find_loc_name(parent_id, locations)
                return parent_loc_name if parent_loc_name else "Parent Location"
            else:
                return loc_name
        else:
            loc_id = loc_data["parent_id"]

    return None


def create_table():
    document = Document()
    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    header_cells = table.rows[0].cells
    header_cells[0].text = "Location"
    header_cells[1].text = "Parent Location"
    header_cells[2].text = "Facility Area"

    with open("your_file.csv", "r") as file:
        reader = csv.DictReader(file)
        locations = {row["loc_id"]: row for row in reader}

    id_to_check = "10"
    result = find_loc_name(id_to_check, locations)
    if result:
        loc_name = locations[id_to_check]["loc_name"]
        parent_id = locations[id_to_check]["parent_id"]
        parent_loc_name = locations[parent_id]["loc_name"]
        row_cells = table.add_row().cells
        row_cells[0].text = loc_name
        row_cells[1].text = parent_loc_name
        row_cells[2].text = result

    document.save("location_table.docx")


create_table()
