"""Document planning engine - analyzes requests and generates structured plans for document generation.

This is the CORE helper that agents MUST call before using any generate tool.
It prevents common mistakes, suggests optimal layouts, and provides ready-to-use templates.
"""

from __future__ import annotations

import json
from typing import Optional

from src.styles.themes import list_themes, get_theme
from src.utils.logger import get_logger

log = get_logger("planner")


class DocumentPlanner:
    """Analyze document request and generate comprehensive plan.
    
    This planner is designed to be called BEFORE any document generation tool.
    It will:
    1. Analyze the user's request to understand requirements
    2. Determine the optimal tool (pptx_create, excel_create, docx_create)
    3. Recommend theme, layout, and slide structure
    4. Generate a JSON template ready to fill with data
    5. Warn about constraints and common mistakes
    """
    
    def __init__(self):
        self.themes = list_themes()
    
    def analyze(
        self,
        request: str,
        doc_type: str = "auto",
    ) -> dict:
        """Analyze a document request and return a comprehensive plan.
        
        Args:
            request: Natural language description of what document to create
            doc_type: Desired document type ('auto', 'presentation', 'spreadsheet', 'document')
            
        Returns:
            Comprehensive plan dictionary with:
            - tool: Recommended tool name
            - reasoning: Why this tool was chosen
            - theme: Theme recommendations
            - slide_plan: Ordered list of slides/sections
            - data_requirements: What data is needed
            - template: Ready-to-use JSON template
            - constraints: Important limitations
            - common_mistakes: What to avoid
            - tips: Best practices
        """
        request_lower = request.lower()
        
        # Step 1: Determine document type
        doc_type = self._detect_doc_type(request, doc_type)
        
        # Step 2: Select appropriate tool
        tool = self._select_tool(doc_type)
        
        # Step 3: Analyze requirements
        theme_suggestion = self._suggest_theme(request, doc_type)
        
        # Step 4: Generate slide/section plan
        slide_plan = self._generate_slide_plan(request, doc_type, theme_suggestion)
        
        # Step 5: Identify data requirements
        data_requirements = self._identify_data_requirements(request, doc_type, slide_plan)
        
        # Step 6: Generate template
        template = self._generate_template(request, doc_type, theme_suggestion, slide_plan)
        
        # Step 7: Generate warnings and tips
        common_mistakes = self._get_common_mistakes(doc_type)
        tips = self._get_best_practices(doc_type)
        constraints = self._get_constraints(doc_type)
        
        return {
            "tool": tool,
            "reasoning": self._explain_tool_choice(doc_type, theme_suggestion),
            "theme": theme_suggestion,
            "slide_plan": slide_plan,
            "data_requirements": data_requirements,
            "template": template,
            "constraints": constraints,
            "common_mistakes": common_mistakes,
            "tips": tips,
        }
    
    def _detect_doc_type(self, request: str, explicit_type: str) -> str:
        """Auto-detect document type from request keywords."""
        if explicit_type != "auto":
            return explicit_type
        
        # Presentation keywords
        pres_keywords = [
            "presentasi", "ppt", "slide", "slide deck", "pitch deck",
            "presentation", "powerpoint", "ppts", "slide show", "slide deck",
            "presentasi", "deklarasi"
        ]
        
        # Spreadsheet keywords
        excel_keywords = [
            "excel", "spreadsheet", "spread sheet", "worksheet", "workbook",
            "table", "tabular", "grid data", "spreadsheet",
        ]
        
        # Document keywords
        doc_keywords = [
            "document", "doc", "word", "report", "laporan",
            "dokumen", "document", "report document", "memorandum",
        ]
        
        for keyword in pres_keywords:
            if keyword in request_lower:
                return "presentation"
        
        for keyword in excel_keywords:
            if keyword in request_lower:
                return "spreadsheet"
        
        for keyword in doc_keywords:
            if keyword in request_lower:
                return "document"
        
        # Default to presentation (most common use case)
        return "presentation"
    
    def _select_tool(self, doc_type: str) -> str:
        """Select the appropriate tool based on document type."""
        tools = {
            "presentation": "pptx_create",
            "spreadsheet": "excel_create",
            "document": "docx_create",
        }
        return tools.get(doc_type, "pptx_create")
    
    def _suggest_theme(self, request: str, doc_type: str) -> dict:
        """Suggest theme based on context and content."""
        request_lower = request.lower()
        
        theme_map = {
            "corporate": [
                "bisnis", "business", "kantor", "company", "professional",
                "meeting", "meeting", "corporate", "office", "laporan",
                "report", "quarterly", "annual",
            ],
            "creative": [
                "creative", "desain", "design", "marketing", "ic", "branding",
                "iklan", "advertising", "campaign", "iklan", "iklan",
            ],
            "dark": [
                "tech", "teknologi", "technology", "startup", "digital",
                "software", "data", "analytics", "developer",
            ],
            "minimal": [
                "minimal", "simple", "clean", "simple", "minimalis",
                "minimalis", "clean", "clean", "minimal", "simple",
            ],
            "academic": [
                "academic", "akademik", "university", "university", "universitas",
                "research", "penelitian", "thesis", "skripsi", "jurnal",
            ],
        }
        
        for theme, keywords in theme_map.items():
            for keyword in keywords:
                if keyword in request_lower:
                    theme_info = next((t for t in self.themes if t['name'] == theme), None)
                    return {
                        "recommended": theme,
                        "alternative": "corporate" if theme != "corporate" else "dark",
                        "reason": f"Theme {theme} matches context of request",
                        "info": theme_info,
                    }
        
        # Default to corporate
        return {
            "recommended": "corporate",
            "alternative": "dark",
            "reason": "Default corporate theme for general business use",
            "info": next((t for t in self.themes if t['name'] == "corporate"), None),
        }
    
    def _generate_slide_plan(self, request: str, doc_type: str, theme: dict) -> list:
        """Generate slide plan based on request analysis."""
        if doc_type == "presentation":
            return self._plan_presentation_slides(request)
        elif doc_type == "spreadsheet":
            return self._plan_spreadsheet_sheets(request)
        else:  # document
            return self._plan_document_sections(request)
    
    def _plan_presentation_slides(self, request: str) -> list:
        """Plan presentation slides with recommended layouts."""
        return [
            {
                "order": 1,
                "visual_type": "cover",
                "title": request[:40] if len(request) > 40 else request,
                "subtitle": "Presentation",
                "icon": "auto",
                "gradient": True,
            },
            {
                "order": 2,
                "visual_type": "exec_summary",
                "title": "Executive Summary",
                "stats_count": 3,
            },
            {
                "order": 3,
                "visual_type": "section_header",
                "title": "Agenda",
            },
            {
                "order": 4,
                "layout": "title_and_content",
                "title": "Content Slide 1",
                "bullets_hint": True,
            },
            {
                "order": 5,
                "layout": "title_and_content",
                "title": "Content Slide 2",
                "bullets_hint": True,
            },
            {
                "order": 6,
                "visual_type": "timeline",
                "title": "Timeline",
                "events_hint": True,
            },
            {
                "order": 7,
                "visual_type": "cta",
                "title": "Thank You",
            },
        ]
    
    def _plan_spreadsheet_sheets(self, request: str) -> list:
        """Plan spreadsheet sheets."""
        return [
            {
                "sheet_name": "Overview",
                "purpose": "Main data summary",
                "columns": ["Category", "Value", "Status"],
                "charts": ["bar", "pie"],
            },
            {
                "sheet_name": "Details",
                "purpose": "Detailed data entry",
                "columns": ["ID", "Item", "Quantity", "Price", "Total"],
            },
        ]
    
    def _plan_document_sections(self, request: str) -> list:
        """Plan document sections."""
        return [
            {"type": "title", "text": request[:60]},
            {"type": "heading_1", "text": "Introduction"},
            {"type": "paragraph", "text": "Brief introduction here"},
            {"type": "heading_1", "text": "Main Content"},
            {"type": "list_bullet", "items": ["Point 1", "Point 2", "Point 3"]},
            {"type": "table", "headers": ["Column 1", "Column 2", "Column 3"]},
            {"type": "heading_1", "text": "Conclusion"},
            {"type": "paragraph", "text": "Summary and recommendations"},
        ]
    
    def _identify_data_requirements(self, request: str, doc_type: str, slide_plan: list) -> list:
        """Identify what data is needed to fill the plan."""
        requirements = []
        
        if doc_type == "presentation":
            requirements.append("Presentation title and subtitle")
            requirements.append("Key metrics or stats for executive summary")
            requirements.append("Content points for each slide")
            requirements.append("Timeline events (if applicable)")
            requirements.append("Action items or next steps")
        
        elif doc_type == "spreadsheet":
            requirements.append("Sheet names and purpose")
            requirements.append("Column headers for each sheet")
            requirements.append("Data rows for each sheet")
            requirements.append("Chart data ranges")
        
        else:  # document
            requirements.append("Document title")
            requirements.append("Section headings")
            requirements.append("Paragraph content for each section")
            requirements.append("List items (if any)")
            requirements.append("Table data (if any)")
        
        return requirements
    
    def _generate_template(self, request: str, doc_type: str, theme: dict, slide_plan: list) -> dict:
        """Generate a ready-to-use JSON template."""
        template = {"filename": "generated_document", "theme": theme["recommended"]}
        
        if doc_type == "presentation":
            template["filename"] = "presentation.pptx"
            template["slide_size"] = "widescreen"
            template["auto_icons"] = True
            template["slides"] = []
            
            for slide in slide_plan:
                slide_template = {
                    "visual_type": slide.get("visual_type", "title_and_content"),
                    "title": slide.get("title", "Slide Title"),
                }
                
                if slide["visual_type"] == "cover":
                    slide_template["subtitle"] = "Subtitle here"
                    slide_template["icon"] = "auto"
                
                elif slide["visual_type"] == "exec_summary":
                    slide_template["stats"] = [
                        {"number": "100%", "label": "Metric 1"},
                        {"number": "50", "label": "Metric 2"},
                        {"number": "$1M", "label": "Metric 3"},
                    ]
                
                elif slide["layout"] == "title_and_content":
                    slide_template["bullets"] = ["Point 1", "Point 2", "Point 3"]
                
                elif slide["visual_type"] == "timeline":
                    slide_template["events"] = [
                        {"title": "Event 1", "date": "Jan 2026"},
                        {"title": "Event 2", "date": "Jun 2026"},
                        {"title": "Event 3", "date": "Dec 2026"},
                    ]
                
                elif slide["visual_type"] == "cta":
                    slide_template["content"] = "Questions?"
                
                template["slides"].append(slide_template)
        
        return template
    
    def _explain_tool_choice(self, doc_type: str, theme: dict) -> str:
        """Explain why a specific tool was chosen."""
        explanations = {
            "presentation": f"Presentation ({doc_type}) requires slide-based layouts. Using pptx_create with {theme['recommended']} theme for professional look.",
            "spreadsheet": f"Spreadsheet ({doc_type}) requires tabular data organization. Using excel_create with formulas and charts support.",
            "document": f"Document ({doc_type}) requires structured sections. Using docx_create with headings, lists, and tables.",
        }
        return explanations.get(doc_type, "Using default presentation tool")
    
    def _get_common_mistakes(self, doc_type: str) -> list:
        """Get common mistakes to avoid for this doc type."""
        mistakes = {
            "presentation": [
                "Don't put too much text per slide - max 3-5 bullets",
                "Use visual_type for premium layouts (cover, exec_summary, timeline)",
                "Set 'auto_icons': true for auto icon detection",
                "Don't use invalid layouts - only: title, title_and_content, section_header, etc.",
                "Keep title slide as first slide (order 1)",
                "Use gradient backgrounds for modern look",
                "Keep text consistent font sizes (28+ for titles, 18+ for content)",
            ],
            "spreadsheet": [
                "Always include headers for each sheet",
                "Ensure row lengths match header count",
                "Use '=' prefix for Excel formulas",
                "Don't use string numbers (e.g., '100' instead of 100)",
                "Max 50 sheets per workbook",
            ],
            "document": [
                "Don't use 'content_paragraphs' (deprecated)",
                "Always use 'sections' array with proper types",
                "Only use supported section types: title, heading_1, paragraph, list_bullet, table",
                "Ensure table rows match header count",
            ],
        }
        return mistakes.get(doc_type, [])
    
    def _get_best_practices(self, doc_type: str) -> list:
        """Get best practices for this doc type."""
        practices = {
            "presentation": [
                "Use visual_type: cover for first slide, exec_summary for overview",
                "Add section_header slides between major topics",
                "Use timeline or roadmap for project updates",
                "End with CTA slide for call-to-action",
                "Use themes: corporate, dark, or creative based on audience",
                "For tables, use build_table_slide method",
                "For comparisons, use build_comparison_slide method",
                "For team slides, use build_team_slide method",
            ],
            "spreadsheet": [
                "Use multiple sheets for different data categories",
                "Add charts to visualize key metrics",
                "Use conditional formatting for status indicators",
                "Include summary sheet with key formulas",
                "Apply consistent styling with themes",
            ],
            "document": [
                "Start with title, subtitle, and TOC",
                "Use heading levels (heading_1, heading_2, heading_3) properly",
                "Include lists for key points",
                "Add tables for structured data",
                "Use themes for consistent styling",
            ],
        }
        return practices.get(doc_type, [])
    
    def _get_constraints(self, doc_type: str) -> dict:
        """Get constraints and limitations."""
        constraints = {
            "presentation": {
                "max_slides": 500,
                "recommended_slides": "10-20 slides for 30-min presentation",
                "max_bullets_per_slide": 6,
                "max_stat_items": 6,
                "file_size_warning": "Keep under 50MB for large presentations",
            },
            "spreadsheet": {
                "max_sheets": 50,
                "max_rows_per_sheet": 1048576,
                "max_columns_per_sheet": 16384,
                "sheet_name_max_length": 31,
            },
            "document": {
                "max_pages": 500,
                "max_table_rows": 10000,
                "max_sections": 500,
            },
        }
        return constraints.get(doc_type, {})
