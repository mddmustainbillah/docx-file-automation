from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

def main():
    # Update these paths according to your file locations
    input_file = "/Users/macbookpro/Desktop/assignment_rokomari/Project eBook Automation/Ebook/90191.docx"
    output_file = "output.docx"
    
    # Process the document
    processor = PageLayoutProcessor(input_file, output_file)
    processor.process()

if __name__ == "__main__":
    main() 