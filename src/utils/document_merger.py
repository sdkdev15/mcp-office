"""Document merger utility for Word and PDF documents."""
import os
import io
from docx import Document
from pypdf import PdfWriter

class DocumentMerger:
    def __init__(self, output_dir: str = "outputs"):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        
    def merge_pdf(self, file_paths: list[str], output_filename: str) -> str:
        """Merge multiple PDF files into one.
        
        Args:
            file_paths: List of paths to PDF files.
            output_filename: Name of the output PDF file.
            
        Returns:
            Path to the merged PDF.
        """
        if not output_filename.endswith(".pdf"):
            output_filename += ".pdf"
            
        merger = PdfWriter()
        
        for path in file_paths:
            if not os.path.exists(path):
                raise FileNotFoundError(f"Cannot merge: file not found: {path}")
            merger.append(path)
            
        out_path = os.path.join(self.output_dir, output_filename)
        with open(out_path, "wb") as f:
            merger.write(f)
            
        merger.close()
        return out_path
        
    def merge_docx(self, file_paths: list[str], output_filename: str) -> str:
        """Merge multiple Word documents into one.
        
        Args:
            file_paths: List of paths to DOCX files.
            output_filename: Name of the output DOCX file.
            
        Returns:
            Path to the merged DOCX.
        """
        if not output_filename.endswith(".docx"):
            output_filename += ".docx"
            
        if not file_paths:
            raise ValueError("No files provided for merging.")
            
        # Open the first document as the base
        base_doc = Document(file_paths[0])
        
        for path in file_paths[1:]:
            if not os.path.exists(path):
                raise FileNotFoundError(f"Cannot merge: file not found: {path}")
                
            # Add a page break between documents
            base_doc.add_page_break()
            
            # Open the document to merge
            sub_doc = Document(path)
            
            # Append elements from the sub-document's body to the base document
            for element in sub_doc.element.body:
                base_doc.element.body.append(element)
                
        out_path = os.path.join(self.output_dir, output_filename)
        base_doc.save(out_path)
        return out_path
