import pytest
import os
import zipfile
import json
from src.utils.template_engine import TemplateEngine
from src.utils.batch_processor import BatchProcessor
from src.utils.document_merger import DocumentMerger

def test_template_engine_basic():
    engine = TemplateEngine()
    
    # Test variables
    assert engine.render("Hello {{name}}!", {"name": "World"}) == "Hello World!"
    
    # Test nested variables
    assert engine.render("Value: {{user.info.age}}", {"user": {"info": {"age": 25}}}) == "Value: 25"
    
    # Test filters
    assert engine.render("{{name | uppercase}}", {"name": "test"}) == "TEST"
    
def test_template_engine_json_filter():
    engine = TemplateEngine()
    # If the user's name has quotes, the json filter should escape them
    res = engine.render('{"name": "{{name | json}}"}', {"name": 'John "Danger" Doe'})
    assert res == '{"name": "John \\"Danger\\" Doe"}'

def test_template_engine_conditionals():
    engine = TemplateEngine()
    template = "Start {{#if show}}Visible{{/if}} End"
    assert engine.render(template, {"show": True}) == "Start Visible End"
    assert engine.render(template, {"show": False}) == "Start  End"

def test_template_engine_loops():
    engine = TemplateEngine()
    template = "Items: {{#each items}}- {{name}} {{/each}}"
    data = {"items": [{"name": "A"}, {"name": "B"}]}
    assert engine.render(template, data) == "Items: - A - B "

def test_batch_processor(tmp_path):
    output_dir = str(tmp_path / "outputs")
    processor = BatchProcessor(output_dir=output_dir)
    
    template = json.dumps({
        "filename": "test_{{id}}.docx",
        "sections": [
            {"type": "paragraph", "text": "Hello {{name}}"}
        ]
    })
    
    datasets = [
        {"id": 1, "name": "Alice"},
        {"id": 2, "name": "Bob"}
    ]
    
    zip_path = processor.process_batch("docx", template, datasets)
    
    assert os.path.exists(zip_path)
    assert zip_path.endswith(".zip")
    
    # Verify zip contents
    with zipfile.ZipFile(zip_path, 'r') as zf:
        namelist = zf.namelist()
        assert "test_1.docx" in namelist
        assert "test_2.docx" in namelist

def test_document_merger(tmp_path):
    output_dir = str(tmp_path / "outputs")
    merger = DocumentMerger(output_dir=output_dir)
    
    # Create two dummy docx files to merge using batch processor
    processor = BatchProcessor(output_dir=output_dir)
    template = json.dumps({
        "filename": "merge_{{id}}.docx",
        "sections": [{"type": "paragraph", "text": "Doc {{id}}"}]
    })
    processor.process_batch("docx", template, [{"id": 1}, {"id": 2}])
    
    # Find the generated batch dir
    batch_dirs = [d for d in os.listdir(output_dir) if d.startswith("batch_")]
    assert len(batch_dirs) > 0
    batch_dir = os.path.join(output_dir, batch_dirs[0])
    
    path1 = os.path.join(batch_dir, "merge_1.docx")
    path2 = os.path.join(batch_dir, "merge_2.docx")
    
    merged_path = merger.merge_docx([path1, path2], "merged_output.docx")
    
    assert os.path.exists(merged_path)
    assert merged_path.endswith(".docx")
