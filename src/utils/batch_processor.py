"""Batch processor for bulk document generation."""
import os
import json
import uuid
import zipfile
from typing import Any
from src.utils.template_engine import TemplateEngine
from src.generators.excel_generator import ExcelGenerator
from src.generators.docx_generator import DOCXGenerator
from src.generators.pptx_generator import PPTXGenerator

class BatchProcessor:
    def __init__(self, output_dir: str = "outputs"):
        self.output_dir = output_dir
        self.template_engine = TemplateEngine()
        os.makedirs(self.output_dir, exist_ok=True)
        
    def process_batch(self, format_type: str, template: str, datasets: list[dict], theme: str = "corporate") -> str:
        """Process a batch of datasets using a template.
        
        Args:
            format_type: 'excel', 'docx', or 'pptx'
            template: JSON string template
            datasets: List of dictionary contexts
            theme: Theme name
            
        Returns:
            Path to the generated ZIP file
        """
        # Create a unique batch directory
        batch_id = uuid.uuid4().hex[:8]
        batch_dir = os.path.join(self.output_dir, f"batch_{batch_id}")
        os.makedirs(batch_dir, exist_ok=True)
        
        generated_files = []
        
        for i, data in enumerate(datasets):
            try:
                # Render the template payload
                payload = self.template_engine.render_json(template, data)
                
                # Determine filename
                filename = payload.get("filename", f"doc_{i+1}")
                if not filename.endswith(f".{self._get_extension(format_type)}"):
                    filename = f"{os.path.splitext(filename)[0]}.{self._get_extension(format_type)}"
                    
                file_path = os.path.join(batch_dir, filename)
                
                # Generate document
                content_bytes = self._generate_document(format_type, payload, theme)
                
                # Write to disk
                with open(file_path, "wb") as f:
                    f.write(content_bytes)
                    
                generated_files.append(file_path)
            except Exception as e:
                # Log error and continue to next dataset
                import logging
                logging.error(f"Failed to generate document {i} in batch: {e}")
                
        # Zip the results
        zip_path = os.path.join(self.output_dir, f"batch_results_{batch_id}.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for fpath in generated_files:
                zipf.write(fpath, os.path.basename(fpath))
                
        return zip_path
        
    def _get_extension(self, format_type: str) -> str:
        if format_type == "excel": return "xlsx"
        if format_type == "docx": return "docx"
        if format_type == "pptx": return "pptx"
        return "bin"
        
    def _generate_document(self, format_type: str, payload: dict, theme: str) -> bytes:
        if format_type == "excel":
            gen = ExcelGenerator(theme_name=theme)
            return gen.create_workbook(payload.get("sheets", []))
        elif format_type == "docx":
            gen = DOCXGenerator(theme_name=theme)
            return gen.create_document_with_content(
                title=payload.get("filename", "Document").replace(".docx", ""),
                sections=payload.get("sections", [])
            )
        elif format_type == "pptx":
            gen = PPTXGenerator(theme_name=theme)
            return gen.create_presentation(
                title=payload.get("filename", "Presentation").replace(".pptx", ""),
                slides=payload.get("slides", [])
            )
        else:
            raise ValueError(f"Unsupported format type: {format_type}")
