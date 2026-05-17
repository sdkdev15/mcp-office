import pytest
from src.analysis.analyzer import Analyzer
from src.analysis.summary_generator import SummaryGenerator
from src.analysis.faq_generator import FAQGenerator
from src.analysis.chart_recommender import ChartRecommender

@pytest.fixture
def sample_data():
    return [
        {"month": "Jan", "revenue": 50000, "expenses": 30000, "category": "A"},
        {"month": "Feb", "revenue": 62000, "expenses": 35000, "category": "B"},
        {"month": "Mar", "revenue": 58000, "expenses": 32000, "category": "A"},
        {"month": "Apr", "revenue": 70000, "expenses": 40000, "category": "C"},
        {"month": "May", "revenue": 75000, "expenses": 42000, "category": "B"},
    ]

def test_analyzer(sample_data):
    analyzer = Analyzer()
    result = analyzer.analyze(sample_data, target_columns=["revenue", "expenses"], breakdown_by="month")
    
    assert "statistics" in result
    assert "revenue" in result["statistics"]
    assert result["statistics"]["revenue"]["mean"] == 63000.0
    assert result["statistics"]["revenue"]["max"] == 75000.0
    
    assert "correlations" in result
    assert len(result["correlations"]) > 0
    assert result["correlations"][0]["columns"] == ["revenue", "expenses"]
    
    assert "trends" in result
    assert len(result["trends"]) > 0

def test_analyzer_empty():
    analyzer = Analyzer()
    result = analyzer.analyze([])
    assert "error" in result

def test_summary_generator(sample_data):
    gen = SummaryGenerator()
    result = gen.generate(sample_data)
    
    assert "summary" in result
    assert "key_metrics" in result
    assert "insights" in result
    assert len(result["key_metrics"]) >= 2
    assert any(m["name"] == "Total Revenue" for m in result["key_metrics"])

def test_faq_generator(sample_data):
    gen = FAQGenerator()
    result = gen.generate(sample_data, num_questions=4)
    
    assert "faq" in result
    assert len(result["faq"]) > 0
    assert "question" in result["faq"][0]
    assert "answer" in result["faq"][0]

def test_chart_recommender(sample_data):
    gen = ChartRecommender()
    result = gen.recommend(sample_data)
    
    assert "charts" in result
    charts = result["charts"]
    assert len(charts) > 0
    types = [c["type"] for c in charts]
    # Expect line/bar/scatter etc. based on rules
    assert "bar" in types or "line" in types or "scatter" in types
