"""Chart Recommendation module."""
import pandas as pd
import numpy as np

class ChartRecommender:
    def recommend(self, data: list[dict], data_types: dict = None, num_recommendations: int = 3) -> dict:
        """Recommend chart types based on data."""
        if not data:
            return {"error": "Empty dataset"}
            
        df = pd.DataFrame(data)
        recommendations = []
        
        # Auto-detect data types if not provided
        if not data_types:
            data_types = {}
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    data_types[col] = "numeric"
                elif pd.api.types.is_datetime64_any_dtype(df[col]) or "date" in col.lower() or "month" in col.lower() or "year" in col.lower():
                    data_types[col] = "time"
                else:
                    data_types[col] = "categorical"
                    
        time_cols = [k for k, v in data_types.items() if v == "time"]
        num_cols = [k for k, v in data_types.items() if v == "numeric"]
        cat_cols = [k for k, v in data_types.items() if v == "categorical"]
        
        # Rule 1: Time Series -> Line Chart
        if time_cols and num_cols:
            recommendations.append({
                "type": "line",
                "columns": [time_cols[0], num_cols[0]],
                "title": f"{num_cols[0].capitalize()} Over Time",
                "reason": "Line charts are best for showing trends over time periods."
            })
            
        # Rule 2: Categorical + Numeric -> Bar/Column Chart
        if cat_cols and num_cols:
            recommendations.append({
                "type": "bar",
                "columns": [cat_cols[0], num_cols[0]],
                "title": f"{num_cols[0].capitalize()} by {cat_cols[0].capitalize()}",
                "reason": "Bar charts are excellent for comparing quantities across different categories."
            })
            
        # Rule 3: 2 Numeric -> Scatter Plot
        if len(num_cols) >= 2:
            recommendations.append({
                "type": "scatter",
                "columns": [num_cols[0], num_cols[1]],
                "title": f"{num_cols[0].capitalize()} vs {num_cols[1].capitalize()}",
                "reason": "Scatter plots reveal relationships or correlations between two numerical variables."
            })
            
        # Rule 4: Parts of a whole -> Pie Chart (if small number of categories)
        if cat_cols and num_cols:
            unique_cats = df[cat_cols[0]].nunique()
            if 2 <= unique_cats <= 6:
                recommendations.append({
                    "type": "pie",
                    "columns": [cat_cols[0], num_cols[0]],
                    "title": f"Distribution of {num_cols[0].capitalize()} across {cat_cols[0].capitalize()}",
                    "reason": "Pie charts effectively show parts-of-a-whole when there are few categories."
                })
                
        # Limit and format output
        recommendations = recommendations[:num_recommendations]
        
        # Add pseudo data_ranges for excel usage if possible
        for i, rec in enumerate(recommendations):
            rec["data_range"] = f"A1:B{len(df) + 1}"
            
        return {"charts": recommendations}
