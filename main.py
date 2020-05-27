from pdf_annotate import PdfAnnotator
import pdfrw
from utils import create_data_dict, annotate_pdf
from slugify import slugify
from pathlib import Path
import pypdftk
import openpyxl
import time

wb = openpyxl.load_workbook('demos.xlsx')
ws = wb['demos']

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

for row in ws.iter_rows(min_row=31, max_row=59):
    a = PdfAnnotator('template.pdf')

    annotate_pdf(a, create_data_dict(row))

    # def write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict):
    #     template_pdf = pdfrw.PdfReader(input_pdf_path)
    #     annotations = template_pdf.pages[0][ANNOT_KEY]
    #     for annotation in annotations:
    #         if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
    #             if annotation[ANNOT_FIELD_KEY]:
    #                 key = annotation[ANNOT_FIELD_KEY][1:-1]
    #                 if key in data_dict.keys():
    #                     annotation.update(
    #                         pdfrw.PdfDict(AP='', V='{}'.format(data_dict[key]))
    #                     )
    #     pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

    folder_directory = Path('results')
    a.write(folder_directory / '{}.pdf'.format(slugify(row[0].value)))
    unflattened = pdfrw.PdfReader(folder_directory / '{}.pdf'.format(slugify(row[0].value)))
    for annotation in unflattened.pages[0]['/Annots']:
        annotation.update(pdfrw.PdfDict(Ff=1))

    pdfrw.PdfWriter().write(folder_directory / '{}.pdf'.format(slugify(row[0].value)), unflattened)

    # pdf_writer.addPage(pdf_reader.getPage(0))
    # page = pdf_writer.getPage(0)
    # # update form fields
    # pdf_writer.updatePageFormFieldValues(page, data_dict)
    # for j in range(0, len(page['/Annots'])):
    #     writer_annot = page['/Annots'][j].getObject()
    #     for field in data_dict:
    #         if writer_annot.get('/T') == field:
    #             writer_annot.update({
    #                 NameObject("/Ff"): NumberObject(1)  # make ReadOnly
    #             })
    # output_stream = BytesIO()
    # pdf_writer.write(output_stream)
    # pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
