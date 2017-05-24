 import docx
import docx2txt
import os
from win32com import client
import PyPDF2 as pyPdf
from docx import Document
import re
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import time
import pythoncom
import threading
import Constants
 
 
 # this function will check if the paragraph in the docx file is a table of plain text.
 def iter_block_items(self, parent):
    
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
	
        # print(parent_elm.xml)
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
	
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# This function will convert docx to text file
 def docx_to_txt(self, filename):
  document = docx.Document(filename)
  lines = []
  content = []
  tempfile = filename.split('\\')
  temp_txt = ''.join(tempfile[-1:])
  tfile = temp_txt[:-5] + ".txt"
  t_file = tfile.split('/ ')
  temp_file = tempfile[:-1]
  temp_file.extend(t_file)
  txtfile = '\\'.join(temp_file).strip()
  f = open (txtfile,'w+')
 
  for block in self.iter_block_items(document):
    
    if isinstance(block, Paragraph):
 
     lines.append(" ".join(block.text.encode('ascii', 'ignore').strip().splitlines()))
  
    else:
     lines.append("<table>")       
     rw = " "      
 
     for row in block.rows:
      for cell in row.cells: 
       rw = rw + "\t"+ cell.text.encode('ascii', 'ignore').strip()
      lines.append(rw)
      rw = " "
    
     lines.append("</table>")

  for i in lines:
	if (i != ''):
		f.write (i)
		f.write('\n')
  f.close()
  self.sections(txtfile)
  self.insert_file_info(txtfile)
  return 200
 
 # To convert doc file to txt, we need to first convert it to docx then docx to txt
  def doc_to_docx(self, filename):

  if threading.currentThread().getName() != 'MainThread':
   pythoncom.CoInitialize()

  word = client.gencache.EnsureDispatch("Word.Application")
  word.Visible = False
  doc =word.Documents.Open(filename)
  doc.SaveAs(filename[:-4]+".docx", FileFormat = 16) # convert to docx
  doc.Close()
  word.Quit()
  new_file = filename[:-4]+".docx"
  self.docx_to_txt(new_file)
  return 200
