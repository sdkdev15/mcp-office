"""Conditional formatting module for Excel."""
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import DataBarRule, ColorScaleRule, CellIsRule
from openpyxl.styles import PatternFill, Font

class ConditionalFormatter:
    def apply_rules(self, ws: Worksheet, rules: list[dict]):
        """Apply a list of conditional formatting rules to a worksheet."""
        for rule in rules:
            rule_type = rule.get("type")
            data_range = rule.get("data_range")
            
            if not rule_type or not data_range:
                continue
                
            if rule_type == "data_bar":
                # openpyxl requires color as hex string without #
                color = rule.get("color", "638EC6").replace("#", "") 
                ws.conditional_formatting.add(
                    data_range,
                    DataBarRule(start_type='min', end_type='max', color=color, showValue="None", minLength=None, maxLength=None)
                )
            elif rule_type == "color_scale":
                start_color = rule.get("start_color", "F8696B").replace("#", "")
                mid_color = rule.get("mid_color", "FFEB84").replace("#", "")
                end_color = rule.get("end_color", "63BE7B").replace("#", "")
                ws.conditional_formatting.add(
                    data_range,
                    ColorScaleRule(start_type='min', start_color=start_color,
                                   mid_type='percentile', mid_value=50, mid_color=mid_color,
                                   end_type='max', end_color=end_color)
                )
            elif rule_type == "cell_is":
                operator = rule.get("operator", "greaterThan")
                formula = rule.get("formula", ["0"])
                if isinstance(formula, str):
                    formula = [formula]
                elif isinstance(formula, (int, float)):
                    formula = [str(formula)]
                    
                fill = None
                font = None
                if "fill_color" in rule:
                    c = rule["fill_color"].replace("#", "")
                    fill = PatternFill(start_color=c, end_color=c, fill_type="solid")
                if "font_color" in rule:
                    c = rule["font_color"].replace("#", "")
                    font = Font(color=c)
                    
                ws.conditional_formatting.add(
                    data_range,
                    CellIsRule(operator=operator, formula=formula, stopIfTrue=True, fill=fill, font=font)
                )
