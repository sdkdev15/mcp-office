"""Template engine for JSON payload generation and batch processing."""
import json
import re
from typing import Any

class TemplateEngine:
    """A lightweight template engine supporting variables, conditionals, loops, and filters."""
    
    def render(self, template_str: str, data: dict) -> str:
        """Render a template string with the provided data context."""
        tokens = self._tokenize(template_str)
        return self._evaluate(tokens, data)
        
    def render_json(self, template_str: str, data: dict) -> dict:
        """Render a template string and parse the result as JSON."""
        rendered = self.render(template_str, data)
        try:
            return json.loads(rendered)
        except json.JSONDecodeError as e:
            # Provide some context around the error
            lines = rendered.split('\n')
            err_line = lines[e.lineno - 1] if e.lineno <= len(lines) else ""
            raise ValueError(f"Template rendered invalid JSON: {str(e)}\nNear: {err_line}")
        
    def _tokenize(self, text: str) -> list[tuple[str, str]]:
        # Split by {{...}}
        pattern = r'(\{\{.*?\}\})'
        parts = re.split(pattern, text)
        tokens = []
        for part in parts:
            if not part: continue
            if part.startswith('{{') and part.endswith('}}'):
                inner = part[2:-2].strip()
                if inner.startswith('#if '):
                    tokens.append(('IF', inner[4:].strip()))
                elif inner == '/if':
                    tokens.append(('ENDIF', ''))
                elif inner.startswith('#each '):
                    tokens.append(('EACH', inner[6:].strip()))
                elif inner == '/each':
                    tokens.append(('ENDEACH', ''))
                else:
                    tokens.append(('VAR', inner))
            else:
                tokens.append(('TEXT', part))
        return tokens

    def _evaluate(self, tokens: list[tuple[str, str]], data: dict, ctx: list[dict] = None) -> str:
        if ctx is None:
            ctx = [data]
            
        result = []
        i = 0
        while i < len(tokens):
            kind, val = tokens[i]
            if kind == 'TEXT':
                result.append(val)
            elif kind == 'VAR':
                result.append(self._eval_var(val, ctx))
            elif kind == 'IF':
                # find matching endif
                depth = 1
                j = i + 1
                while j < len(tokens):
                    if tokens[j][0] == 'IF': depth += 1
                    elif tokens[j][0] == 'ENDIF':
                        depth -= 1
                        if depth == 0: break
                    j += 1
                
                if j == len(tokens):
                    raise ValueError(f"Unclosed {{{{str('#if')}}}} block")
                
                cond_val = self._resolve(val, ctx)
                # Truthy evaluation
                if cond_val:
                    inner_res = self._evaluate(tokens[i+1:j], data, ctx)
                    result.append(inner_res)
                i = j
            elif kind == 'EACH':
                depth = 1
                j = i + 1
                while j < len(tokens):
                    if tokens[j][0] == 'EACH': depth += 1
                    elif tokens[j][0] == 'ENDEACH':
                        depth -= 1
                        if depth == 0: break
                    j += 1
                    
                if j == len(tokens):
                    raise ValueError(f"Unclosed {{{{str('#each')}}}} block")
                
                items = self._resolve(val, ctx)
                if isinstance(items, list):
                    for item in items:
                        # If item is not a dict, wrap it so we can reference it via 'this'
                        item_ctx = item if isinstance(item, dict) else {"this": item}
                        new_ctx = ctx + [item_ctx]
                        inner_res = self._evaluate(tokens[i+1:j], data, new_ctx)
                        result.append(inner_res)
                i = j
            i += 1
        return "".join(result)

    def _resolve(self, path: str, ctx: list[dict]) -> Any:
        if not path or path == "this":
            return ctx[-1] if ctx else None
            
        for c in reversed(ctx):
            if not isinstance(c, dict): continue
            
            parts = path.split('.')
            val = c
            found = True
            for p in parts:
                if isinstance(val, dict) and p in val:
                    val = val[p]
                else:
                    found = False
                    break
            if found:
                return val
        return None

    def _eval_var(self, val: str, ctx: list[dict]) -> str:
        parts = val.split('|')
        var_name = parts[0].strip()
        resolved = self._resolve(var_name, ctx)
        
        if resolved is None:
            res_str = ""
        else:
            res_str = str(resolved)
            
        # apply filters
        is_json_escaped = False
        for filter_name in parts[1:]:
            f = filter_name.strip().lower()
            if f == 'uppercase': res_str = res_str.upper()
            elif f == 'lowercase': res_str = res_str.lower()
            elif f == 'json': 
                res_str = json.dumps(res_str)[1:-1]
                is_json_escaped = True
                
        # If the variable contains quotes or newlines and is likely injected into JSON
        # and wasn't explicitly formatted as json, we should escape newlines at least.
        # But this is a basic engine. Best practice is for users to use | json.
        return res_str
