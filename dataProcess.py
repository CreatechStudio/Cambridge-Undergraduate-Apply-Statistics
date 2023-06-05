import re
import openpyxl

def parse_data(filename):
    with open(filename, 'r') as file:
        lines = file.readlines()

    data = []
    university = ''
    year = ''
    current_name = ''
    current_categories = []
    current_data = []

    for line in lines:
        line = line.strip()
        if line.startswith('University:'):
            university = line.split(':')[1].strip()
        elif line.startswith('Year:'):
            year = line.split(':')[1].strip()
        elif line.startswith('Name:'):
            current_name = line.split(':')[1].strip()
        elif line.startswith('Data:'):
            current_data = re.findall(r'\d+', line)
            current_data = list(map(int, current_data))

            if len(current_data) < len(current_categories):
                current_data += [0] * (len(current_categories) - len(current_data))

            data.append({
                'University': university,
                'Year': year,
                'Name': current_name,
                'Categories': current_categories,
                'Data': current_data
            })

        elif line.startswith('Categories:'):
            current_categories = re.findall(r"'(.*?)'", line)

    return data


def fill_missing_data(data, categories):
    for item in data:
        missing_categories = set(categories) - set(item['Categories'])
        missing_categories_data = [0] * len(missing_categories)
        for category in missing_categories:
            index = categories.index(category)
            item['Categories'].insert(index, category)
            item['Data'].insert(index, 0)

    return data


def write_to_excel(data, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write header row with complete category list
    header = ['University', 'Year', 'Name'] + categories
    sheet.append(header)

    # Write data rows
    for item in data:
        row = [item['University'], item['Year'], item['Name']]
        row_data = item['Data']

        if len(row_data) < len(categories):
            row_data += [0] * (len(categories) - len(row_data))

        row += row_data
        sheet.append(row)

    workbook.save(output_file)


filename = 'output.txt'
output_file = 'output.xlsx'

parsed_data = parse_data(filename)

# Extract all categories from the data
categories = [
    "Anglo-Saxon, Norse, and Celtic",
    "Archaeology",
    "Architecture",
    "Asian and Middle Eastern Studies",
    "Chemical Engineering via Engineering",
    "Chemical Engineering via Natural Sciences",
    "Classics",
    "Classics (4 years)",
    "Computer Science",
    "Economics",
    "Education",
    "Engineering",
    "English",
    "Foundation Year in Arts, Humanities and Social Sciences",
    "Geography",
    "History",
    "History and Modern Languages",
    "History and Politics",
    "History of Art",
    "Human, Social, and Political Sciences",
    "Land Economy",
    "Law",
    "Linguistics",
    "Mathematics",
    "Medicine",
    "Medicine (Graduate course)",
    "Modern and Medieval Languages",
    "Music",
    "Natural Sciences",
    "Philosophy",
    "Psychological and Behavioural Sciences",
    "Theology, Religion and Philosophy of Religion",
    "Veterinary Medicine"
]

filled_data = fill_missing_data(parsed_data, categories)
write_to_excel(filled_data, output_file)
