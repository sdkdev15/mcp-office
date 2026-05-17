# MCP Office Analysis Guide

This guide explains how to use the Analysis & Content Generation tools (Track B) to process raw data and produce intelligent insights and recommendations for Office document generation.

## Overview

The Analysis tools are designed to work together to turn raw datasets into ready-to-present material. Instead of providing the AI with raw data to figure out manually, you can use these tools to extract statistical meaning, summaries, and Q&A content.

The tools provided are:
1. `analyze_data` - Deep statistical analysis (mean, median, correlation, outliers, trends).
2. `generate_summary` - High-level text and key metrics suitable for executive summaries.
3. `generate_faq` - Auto-generated Q&A based on the data.
4. `recommend_charts` - Suggested visualizations based on data dimensions.

---

## 1. analyze_data

Performs a full statistical analysis on a JSON dataset.

**When to use:** When you need deep numerical insights, want to discover correlations, or identify anomalies.

### Example Payload
```json
{
  "data": [
    {"month": "Jan", "revenue": 50000, "expenses": 30000},
    {"month": "Feb", "revenue": 62000, "expenses": 35000}
  ],
  "target_columns": ["revenue", "expenses"],
  "breakdown_by": "month"
}
```

---

## 2. generate_summary

Summarizes the data in plain text and extracts key performance metrics.

**When to use:** Ideal for generating the "Executive Summary" section in a Word document (`docx_create`) or the introductory slide in a PowerPoint (`pptx_create`).

### Example Payload
```json
{
  "data": [...],
  "style": "professional",
  "include_metrics": true,
  "max_insights": 5
}
```

---

## 3. generate_faq

Automatically formulates insightful questions and accurate answers from the data.

**When to use:** Perfect for an "Appendix" or "Q&A" section in a document.

### Example Payload
```json
{
  "data": [...],
  "num_questions": 5,
  "question_style": "practical"
}
```

---

## 4. recommend_charts

Recommends the best chart types (bar, line, scatter, pie) based on the shape and types of data columns.

**When to use:** Before calling `excel_create` or `pptx_create` when you are unsure what visualizations will best represent the dataset.

### Example Payload
```json
{
  "data": [...],
  "num_recommendations": 3
}
```

---

## Recommended Workflow

1. **Step 1:** You have a dataset (e.g., from an API or database).
2. **Step 2:** Call `generate_summary` to get the text narrative.
3. **Step 3:** Call `recommend_charts` to figure out what visuals to include.
4. **Step 4:** Call `excel_create` or `docx_create`, injecting the text from Step 2 into the paragraphs and using the chart configurations from Step 3.
