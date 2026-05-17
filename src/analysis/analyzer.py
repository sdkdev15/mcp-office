"""Statistical data analysis module."""
import pandas as pd
import numpy as np

class Analyzer:
    def analyze(self, data: list[dict], target_columns: list[str] = None, breakdown_by: str = None) -> dict:
        """Analyze a dataset and return statistics, trends, and correlations."""
        if not data:
            return {"error": "Empty dataset"}
            
        df = pd.DataFrame(data)
        
        if not target_columns:
            target_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            
        results = {
            "statistics": {},
            "trends": [],
            "outliers": [],
            "correlations": [],
            "distributions": {}
        }
        
        for col in target_columns:
            if col not in df.columns:
                continue
                
            # Convert to numeric, coercing errors
            df[col] = pd.to_numeric(df[col], errors='coerce')
            series = df[col].dropna()
            
            if series.empty:
                continue
                
            stats = {
                "mean": float(series.mean()),
                "median": float(series.median()),
                "stddev": float(series.std()) if len(series) > 1 else 0.0,
                "min": float(series.min()),
                "max": float(series.max()),
                "q1": float(series.quantile(0.25)),
                "q3": float(series.quantile(0.75))
            }
            results["statistics"][col] = stats
            
            # Outliers (IQR method)
            iqr = stats["q3"] - stats["q1"]
            lower_bound = stats["q1"] - 1.5 * iqr
            upper_bound = stats["q3"] + 1.5 * iqr
            outliers = series[(series < lower_bound) | (series > upper_bound)]
            for idx, val in outliers.items():
                # Get a simple identifier if possible
                identifier = df.iloc[idx].get(breakdown_by) if breakdown_by and breakdown_by in df.columns else f"Row {idx}"
                results["outliers"].append({
                    "column": col,
                    "identifier": str(identifier),
                    "value": float(val),
                    "reason": "Outside 1.5 * IQR"
                })
                
        # Correlations
        if len(target_columns) >= 2:
            # Only use numeric columns for correlation
            num_df = df[target_columns].select_dtypes(include=[np.number])
            if len(num_df.columns) >= 2:
                corr_matrix = num_df.corr()
                for i in range(len(num_df.columns)):
                    for j in range(i + 1, len(num_df.columns)):
                        col1 = num_df.columns[i]
                        col2 = num_df.columns[j]
                        coef = corr_matrix.loc[col1, col2]
                        if not pd.isna(coef):
                            results["correlations"].append({
                                "columns": [col1, col2],
                                "coefficient": float(coef)
                            })
                            
        # Trends
        if breakdown_by and breakdown_by in df.columns:
            grouped = df.groupby(breakdown_by)[target_columns].mean()
            for col in target_columns:
                if col in grouped.columns:
                    values = grouped[col].dropna().tolist()
                    if len(values) >= 2:
                        start_val = values[0]
                        end_val = values[-1]
                        change = ((end_val - start_val) / start_val) * 100 if start_val != 0 else 0
                        direction = "increasing" if change > 0 else ("decreasing" if change < 0 else "stable")
                        results["trends"].append(
                            f"{col.capitalize()} is {direction} overall ({change:+.1f}% from first to last {breakdown_by})."
                        )
                        
        return results
