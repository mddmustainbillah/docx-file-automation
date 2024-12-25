from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

class PageLayoutProcessor:
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path
        
    def process(self):
        """Process the document's page layout"""
        try:
            # Load the document
            doc = Document(self.input_path)
            
            # Process each section in the document
            for section in doc.sections:
                # 1. Set orientation to portrait
                section.orientation = WD_ORIENT.PORTRAIT
                
                # 2. Set all margins to 1 inch
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                
                # 3. Remove headers
                if section.header:
                    section.header.is_linked_to_previous = True
                    for paragraph in section.header.paragraphs:
                        p = paragraph._element
                        p.getparent().remove(p)
                
                # 4. Remove footers
                if section.footer:
                    section.footer.is_linked_to_previous = True
                    for paragraph in section.footer.paragraphs:
                        p = paragraph._element
                        p.getparent().remove(p)
                
                # 5. Remove page borders
                try:
                    section._sectPr.remove_all("w:pgBorders")
                except:
                    pass

                # 6. Check and convert multiple columns to single column
                self._convert_to_single_column(section)
            
            # 7. Set line spacing for all text elements
            self._set_line_spacing(doc)
            
            # Save the processed document
            doc.save(self.output_path)
            print(f"Successfully processed document layout: {self.output_path}")
            
        except Exception as e:
            print(f"Error processing document: {str(e)}")

    def _convert_to_single_column(self, section):
        """Check and convert multiple columns to single column"""
        try:
            # Get the columns element
            cols = section._sectPr.xpath("./w:cols")[0]
            
            # Check if document has multiple columns
            num_cols = int(cols.get(qn('w:num'))) if cols.get(qn('w:num')) else 1
            
            if num_cols > 1:
                print(f"Converting from {num_cols} columns to single column")
                # Set to single column
                cols.set(qn('w:num'), '1')
                
                # Remove any column spacing
                if cols.get(qn('w:space')):
                    cols.set(qn('w:space'), '0')
                    
                # Remove any specific column width settings
                for child in cols.getchildren():
                    cols.remove(child)
        except Exception as e:
            print(f"Error checking/converting columns: {str(e)}")

    def _set_line_spacing(self, doc):
        """Set line spacing to 1.15 for all text elements in the document"""
        try:
            # Set spacing for all styles in the document
            for style in doc.styles:
                if hasattr(style, 'paragraph_format'):
                    style.paragraph_format.line_spacing = 1.15

            # Process all paragraphs in main document
            for paragraph in doc.paragraphs:
                self._apply_spacing_to_paragraph(paragraph)
            
            # Process paragraphs in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            self._apply_spacing_to_paragraph(paragraph)
            
            # Process headers
            for section in doc.sections:
                if section.header:
                    for paragraph in section.header.paragraphs:
                        self._apply_spacing_to_paragraph(paragraph)
                        
                # Process footers
                if section.footer:
                    for paragraph in section.footer.paragraphs:
                        self._apply_spacing_to_paragraph(paragraph)
            
            print("Successfully set line spacing to 1.15 for all text elements")
        except Exception as e:
            print(f"Error setting line spacing: {str(e)}")

    def _apply_spacing_to_paragraph(self, paragraph):
        """Apply 1.15 line spacing to a paragraph"""
        if paragraph._element is not None:
            # Set line spacing at paragraph format level
            paragraph.paragraph_format.line_spacing = 1.15
            
            # Ensure the spacing is applied at the XML level
            if paragraph._p.pPr is None:
                paragraph._p.get_or_add_pPr()
            
            # Set spacing in XML
            spacing = paragraph._p.pPr.xpath('./w:spacing')
            if not spacing:
                spacing_element = OxmlElement('w:spacing')
                spacing_element.set(qn('w:line'), str(int(240 * 1.15)))  # 240 twips = 1 line
                spacing_element.set(qn('w:lineRule'), 'auto')
                paragraph._p.pPr.append(spacing_element)
            else:
                spacing[0].set(qn('w:line'), str(int(240 * 1.15)))
                spacing[0].set(qn('w:lineRule'), 'auto')

def main():
    # Update these paths according to your file locations
    input_file = "/Users/macbookpro/Desktop/assignment_rokomari/Project eBook Automation/Ebook/90191.docx"
    output_file = "output.docx"
    
    # Process the document
    processor = PageLayoutProcessor(input_file, output_file)
    processor.process()

if __name__ == "__main__":
    main() 