from pdf_annotate import PdfAnnotator, Location, Appearance

a = PdfAnnotator('template.pdf')


def create_annotation(node: PdfAnnotator, is_name: bool, count: int, value):
    coordinate_dict_name = {
        '1': {
            'x1': 95,
            'y1': 720,
            'x2': 160,
            'y2': 740
        },
        '2': {
            'x1': 335,
            'y1': 720,
            'x2': 405,
            'y2': 740
        },
        '3': {
            'x1': 95,
            'y1': 690,
            'x2': 160,
            'y2': 710
        },
        '4': {
            'x1': 335,
            'y1': 690,
            'x2': 405,
            'y2': 710
        }
    }
    coordinate_dict_id = {
        '1': {
            'x1': 215,
            'y1': 720,
            'x2': 295,
            'y2': 740
        },
        '2': {
            'x1': 465,
            'y1': 720,
            'x2': 550,
            'y2': 740
        },
        '3': {
            'x1': 215,
            'y1': 690,
            'x2': 295,
            'y2': 710
        },
        '4': {
            'x1': 465,
            'y1': 690,
            'x2': 550,
            'y2': 710
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


def process_name(name):
    try:
        if name:
            return name.strip().split()[0][0:12]
        else:
            return ""
    except:
        return ""


def create_data_dict(row):
    return {
        'name_1': process_name(str(row[2].value)),
        'id_1': row[3].value or "",
        'name_2': process_name(str(row[4].value)),
        'id_2': row[5].value or "",
        'name_3': process_name(str(row[6].value)),
        'id_3': row[7].value or "",
        'name_4': process_name(str(row[8].value)),
        'id_4': row[9].value or "",
    }


def annotate_pdf(node: PdfAnnotator, data: dict):
    create_annotation(node, is_name=True, count=1, value=str(data['name_1']))
    create_annotation(node, is_name=False, count=1, value=str(data['id_1']))
    create_annotation(node, is_name=True, count=2, value=str(data['name_2']))
    create_annotation(node, is_name=False, count=2, value=str(data['id_2']))
    create_annotation(node, is_name=True, count=3, value=str(data['name_3']))
    create_annotation(node, is_name=False, count=3, value=str(data['id_3']))
    create_annotation(node, is_name=True, count=4, value=str(data['name_4']))
    create_annotation(node, is_name=False, count=4, value=str(data['id_4']))
