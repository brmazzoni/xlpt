from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

from styles import *


def gen_template(headers, header_line=1, header_fill_color=HEADER_FILL_COLOR, filename='template.xlsx', metadata=None):

  # Define (default medium gray) solid color fill for headers
  header_fill = PatternFill(
    fill_type='solid', 
    start_color=header_fill_color,
    end_color=header_fill_color)

  # Initiate new workbook and select worksheet
  wb = Workbook()
  ws = wb.active

  # Write metadata
  if metadata is not None:
    for section, data in metadata.items():
      if 'text' in data.keys():
        for cell, text in data['text'].items():
          ws[cell] = text
          ws[cell].font = Font(size=data['font'])
      for row in ws[data['range']]:
        for cell in row:
          pass
          #ws[f'{cell.column_letter}{cell.row}'] = section
        
        


  # Generate headers
  for col, header in headers.items():
    cell = col + str(header_line)
    ws[cell] = header['text']
    ws[cell].fill = header_fill
    width = header['width'] if 'width' in header.keys() else CW_NORMAL
    ws.column_dimensions[col].width = width
  

  # Style columns (10 empty lines)
  EMPTY_LINES = 10
  for row in ws[f'A{header_line}:{max(headers)}{EMPTY_LINES+header_line}']:
    for cell in row:
      cell.border = Border(
        left=Side(border_style='thin', color='FF000000'),
        right=Side(border_style='thin', color='FF000000'),
        top=Side(border_style='thin', color='FF000000'),
        bottom=Side(border_style='thin', color='FF000000')
      )
    
  for col, header in headers.items():
    for cell in ws[f'{col}{header_line}:{col}{EMPTY_LINES+header_line}']:
      wrap_text = header['wrap'] if 'wrap' in header.keys() else False
      cell[0].alignment = Alignment(wrap_text=wrap_text)


  # Save file
  wb.save(filename)

if __name__ == '__main__':

  metadata = {
    'section1': {
      'range': 'A1:B3', 
      'image': 'img/logo.png',
      'borders': 'outer' # outer/all
    },
    'section2': {
      'range': 'C1:D1',
      'text': {'C1': 'TITLE'},
      'font': 24
    },
    'section3': {
      'range': 'E1:F1',
      'text': {'E1': 'Subtitle'},
      'font': 12
    },
    'section4': {
      'range': 'C2:D3',
      'text': {'C2': 'author', 'C3': 'version'},
      'font': 10,
      'borders': 'outer'
    },
    'section5': {
      'range': 'E2:F3',
      'text': {'E2': '<author name>', 'E3': '<version number>'},
      'font': 10
    }
  }

  headers = {
    'A': {'text': 'c1', 'width': CW_SHORT},
    'B': {'text': 'column2', 'width': CW_LONG, 'wrap': True},
    'C': {'text': 'column3', 'wrap_text': True},
  }
  gen_template(headers, header_line=5, metadata=metadata)
