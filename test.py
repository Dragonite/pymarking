from pdf_annotate import PdfAnnotator, Location, Appearance
import openpyxl

a = PdfAnnotator('template.pdf')

def create_annotation(node, is_name, count, value):
    coordinate_dict_name = {
        '3': {
            'x1': 95,
            'y1': 600,
            'x2': 160,
            'y2': 800
        }
    }
    coordinate_dict_id = {
        '3': {
            'x1': 215,
            'y1': 600,
            'x2': 295,
            'y2': 800
        }
    }

    if is_name:
        return node.add_annotation(
            'text',
            Location(x1=coordinate_dict_name[str(count)]['x1'], y1=coordinate_dict_name[str(count)]['y1'],
                     x2=coordinate_dict_name[str(count)]['x2'], y2=coordinate_dict_name[str(count)]['y2'], page=0),
            Appearance(stroke_color=(0, 0, 0), stroke_width=5, content=value, fill=(0, 0, 0, 1))
        )
    else:
        return node.add_annotation(
            'text',
            Location(x1=coordinate_dict_id[str(count)]['x1'], y1=coordinate_dict_id[str(count)]['y1'],
                     x2=coordinate_dict_id[str(count)]['x2'], y2=coordinate_dict_id[str(count)]['y2'], page=0),
            Appearance(stroke_color=(0, 0, 0), stroke_width=5, content=value, fill=(0, 0, 0, 1))
        )


create_annotation(a, is_name=True, count=3, value='Haolin Wu')
create_annotation(a, is_name=False, count=3, value='21706137')

# # Name 3
# a.add_annotation(
#     'text',
#     Location(x1=95, y1=600, x2=160, y2=800, page=0),
#     Appearance(stroke_color=(1, 0, 0), stroke_width=5, content="hello worldd", fill=(0.705, 0.094, 0.125, 1))
# )
# # ID 3
# a.add_annotation(
#     'text',
#     Location(x1=215, y1=600, x2=295, y2=800, page=0),
#     Appearance(stroke_color=(1, 0, 0), stroke_width=5, content="hello world", fill=(0.705, 0.094, 0.125, 1))
# )
a.write('b.pdf')

wb = openpyxl.load_workbook('demos.xlsx')
ws = wb['demos']

for row in ws.iter_rows(min_row=31, max_row=59):
    print('new row')
    filename = '{id_1}.pdf'.format(id_1=row[3].value)
    data_dict = {
        'Name_1': row[2].value,
        'ID_1': row[3].value,
        'Name_2': row[4].value,
        'ID_2': row[5].value,
        'Name_3': row[6].value,
        'ID_3': row[7].value,
        'Name_4': row[8].value,
        'ID_4': row[9].value,
    }
    print(data_dict)
