"""FAQ Generation module."""
from src.analysis.analyzer import Analyzer

class FAQGenerator:
    def generate(self, data: list[dict], num_questions: int = 5, question_style: str = "practical") -> dict:
        """Generate Q&A pairs from data insights."""
        if not data:
            return {"error": "Empty dataset"}
            
        analyzer = Analyzer()
        import pandas as pd
        import numpy as np
        
        df = pd.DataFrame(data)
        num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        analysis = analyzer.analyze(data, target_columns=num_cols)
        if "error" in analysis:
            return analysis
            
        stats = analysis.get("statistics", {})
        faqs = []
        
        for col, stat in stats.items():
            if len(faqs) >= num_questions:
                break
                
            if question_style == "practical":
                q1 = f"What is the average {col.lower()}?"
            else:
                q1 = f"Could you provide the mean value for {col}?"
                
            a1 = f"The average {col.lower()} across the dataset is {round(stat['mean'], 2)}."
            faqs.append({"question": q1, "answer": a1})
            
            if len(faqs) >= num_questions:
                break
                
            if question_style == "practical":
                q2 = f"What was the maximum {col.lower()} recorded?"
            else:
                q2 = f"What is the peak {col} value?"
                
            a2 = f"The highest recorded {col.lower()} is {round(stat['max'], 2)}."
            faqs.append({"question": q2, "answer": a2})
            
        for trend in analysis.get("trends", []):
            if len(faqs) >= num_questions:
                break
                
            col = trend.split()[0]
            q = f"What is the general trend for {col}?"
            faqs.append({"question": q, "answer": trend})
            
        return {"faq": faqs}
