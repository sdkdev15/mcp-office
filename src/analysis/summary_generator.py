"""Summary generation module."""
import pandas as pd
import numpy as np
from src.analysis.analyzer import Analyzer

class SummaryGenerator:
    def generate(self, data: list[dict], style: str = "professional", include_metrics: bool = True, max_insights: int = 5) -> dict:
        """Generate high-level summaries and key metrics from data."""
        if not data:
            return {"error": "Empty dataset"}
            
        analyzer = Analyzer()
        df = pd.DataFrame(data)
        num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        analysis = analyzer.analyze(data, target_columns=num_cols)
        if "error" in analysis:
            return analysis
            
        stats = analysis.get("statistics", {})
        
        summary_text = "Data analysis complete. "
        key_metrics = []
        insights = []
        highlights = []
        
        for col, stat in stats.items():
            total = df[col].sum() if col in df.columns else 0
            key_metrics.append({
                "name": f"Total {col.capitalize()}",
                "value": round(float(total), 2),
                "mean": round(stat["mean"], 2),
                "max": round(stat["max"], 2)
            })
            
            insights.append(f"The average {col.lower()} is {round(stat['mean'], 2)}, with a peak of {round(stat['max'], 2)}.")
            
        for trend in analysis.get("trends", []):
            insights.append(trend)
            if "increasing" in trend:
                highlights.append(f"Growth observed in {trend.split()[0]}")
                
        for corr in analysis.get("correlations", []):
            if abs(corr["coefficient"]) > 0.7:
                insights.append(f"Strong correlation identified between {corr['columns'][0]} and {corr['columns'][1]} ({corr['coefficient']:.2f}).")
                
        if stats:
            summary_text += f"The dataset spans {len(data)} records across {len(stats)} key numeric metrics. "
            
        if style == "casual":
            summary_text = f"Here's a quick look at your data! We analyzed {len(data)} rows. "
            
        insights = insights[:max_insights]
            
        return {
            "summary": summary_text.strip(),
            "key_metrics": key_metrics if include_metrics else [],
            "insights": insights,
            "highlights": highlights
        }
