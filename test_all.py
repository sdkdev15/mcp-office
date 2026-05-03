"""Test all generators directly."""
import sys, os
from pathlib import Path

# Ensure src is importable
sys.path.insert(0, str(Path(__file__).parent))
os.makedirs("outputs/test", exist_ok=True)

print("=" * 50)
print("Testing All Generators")
print("=" * 50)

# Test Excel
print("\n[1/3] Testing Excel Generator...")
from src.generators.excel_generator import ExcelGenerator
gen = ExcelGenerator("corporate")
sheets = [
    {
        "name": "Penjualan",
        "headers": ["Produk", "Revenue", "Units", "Growth"],
        "rows": [
            ["Product A", 15000000, 120, "12.5%"],
            ["Product B", 25000000, 200, "8.3%"],
            ["Product C", 18000000, 150, "15.1%"],
            ["Product D", 32000000, 310, "22.0%"],
            ["Product E", 12000000, 95, "5.7%"],
        ],
    },
    {
        "name": "Summary",
        "headers": ["Metric", "Value"],
        "rows": [
            ["Total Revenue", "Rp 102.000.000"],
            ["Total Units", "875"],
            ["Avg Growth", "12.7%"],
            ["Top Product", "Product D"],
        ],
    },
]
data = gen.create_workbook(sheets, metadata={"author": "Test", "title": "Q1 Sales"})
path = "outputs/test/Laporan_Penjualan_Q1.xlsx"
Path(path).write_bytes(data)
print(f"  OK: {path} ({len(data):,} bytes)")

# Test Word
print("\n[2/3] Testing Word Generator...")
from src.generators.docx_generator import DOCXGenerator
gen = DOCXGenerator("corporate")
data = gen.create_document(
    "Laporan Bisnis Q1 2025",
    page_size="A4",
    orientation="portrait",
    metadata={"author": "Test", "title": "Laporan Bisnis"},
)
path = "outputs/test/Laporan_Bisnis_Q1.docx"
Path(path).write_bytes(data)
print(f"  OK: {path} ({len(data):,} bytes)")

# Test PowerPoint
print("\n[3/3] Testing PowerPoint Generator...")
from src.generators.pptx_generator import PPTXGenerator
gen = PPTXGenerator("corporate")
data = gen.create_presentation(
    "Presentasi Q1 2025",
    slide_size="widescreen",
    metadata={"author": "Test", "title": "Presentasi Q1"},
)
path = "outputs/test/Presentasi_Q1.pptx"
Path(path).write_bytes(data)
print(f"  OK: {path} ({len(data):,} bytes)")

print("\n" + "=" * 50)
print("All 3 generators work!")
print(f"Files saved to: outputs/test/")
print("=" * 50)