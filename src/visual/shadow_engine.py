"""Shadow engine — applies outer & inner shadow effects to shapes via direct XML manipulation.

Supports:
- Outer shadows (blur, offset, color, opacity)
- Inner shadows
- Apply to shapes or text runs
"""

from __future__ import annotations

from typing import Optional
from pptx.oxml.ns import qn
from lxml import etree


class ShadowDef:
    """Definition of a shadow effect."""

    def __init__(
        self,
        shadow_type: str = "outer",
        blur_radius: int = 6,
        offset_x: float = 2.0,
        offset_y: float = 2.0,
        color: str = "#000000",
        opacity: int = 40,
    ):
        self.shadow_type = shadow_type  # "outer" or "inner"
        self.blur_radius = blur_radius
        self.offset_x = offset_x
        self.offset_y = offset_y
        self.color = color
        self.opacity = opacity  # 0-100

    @classmethod
    def soft(cls, color: str = "#000000", opacity: int = 30) -> ShadowDef:
        """Soft outer shadow, no offset."""
        return cls(shadow_type="outer", blur_radius=8, offset_x=0, offset_y=0, color=color, opacity=opacity)

    @classmethod
    def drop(cls, offset_x: float = 3, offset_y: float = 3, blur: int = 6, color: str = "#000000", opacity: int = 40) -> ShadowDef:
        """Drop shadow with offset."""
        return cls(shadow_type="outer", blur_radius=blur, offset_x=offset_x, offset_y=offset_y, color=color, opacity=opacity)

    @classmethod
    def glow(cls, color: str = "#3B82F6", size: int = 10) -> ShadowDef:
        """Glow effect (colored soft shadow, centered)."""
        return cls(shadow_type="outer", blur_radius=size, offset_x=0, offset_y=0, color=color, opacity=60)

    @classmethod
    def inner(cls, blur: int = 4, opacity: int = 30) -> ShadowDef:
        """Inner shadow for inset effect."""
        return cls(shadow_type="inner", blur_radius=blur, offset_x=1, offset_y=1, color="#000000", opacity=opacity)


class ShadowEngine:
    """Apply shadow effects to shapes and text."""

    @staticmethod
    def _create_outer_shadow_xml(shadow: ShadowDef) -> etree.Element:
        """Build an a:outerShdw XML element."""
        os_elem = etree.Element(qn("a:outerShdw"))
        os_elem.set("blurRad", str(shadow.blur_radius * 635))  # Convert pt to EMU
        os_elem.set("dist", str(int((shadow.offset_x ** 2 + shadow.offset_y ** 2) ** 0.5 * 635)))
        os_elem.set("dir", str(int(135 * 60000)))  # Default direction
        # Color
        srgb = etree.SubElement(os_elem, qn("a:srgbClr"))
        srgb.set("val", shadow.color.lstrip("#"))
        alpha = etree.SubElement(srgb, qn("a:alpha"))
        alpha.set("val", str(int(shadow.opacity * 1000)))  # 0-100000
        return os_elem

    @staticmethod
    def _create_inner_shadow_xml(shadow: ShadowDef) -> etree.Element:
        """Build an a:innerShdw XML element."""
        is_elem = etree.Element(qn("a:innerShdw"))
        is_elem.set("blurRad", str(shadow.blur_radius * 635))
        is_elem.set("dist", str(int(((shadow.offset_x ** 2 + shadow.offset_y ** 2) ** 0.5) * 635)))
        is_elem.set("dir", str(int(135 * 60000)))
        srgb = etree.SubElement(is_elem, qn("a:srgbClr"))
        srgb.set("val", shadow.color.lstrip("#"))
        alpha = etree.SubElement(srgb, qn("a:alpha"))
        alpha.set("val", str(int(shadow.opacity * 1000)))
        return is_elem

    @staticmethod
    def apply_to_shape(shape, shadow: ShadowDef) -> None:
        """Apply shadow to a shape."""
        try:
            sp_pr = shape._element.find(qn("p:spPr"))
            if sp_pr is None:
                return
            # Remove existing shadows
            for old in sp_pr.findall(qn("a:outerShdw")):
                sp_pr.remove(old)
            for old in sp_pr.findall(qn("a:innerShdw")):
                sp_pr.remove(old)

            if shadow.shadow_type == "outer":
                sh_xml = ShadowEngine._create_outer_shadow_xml(shadow)
            else:
                sh_xml = ShadowEngine._create_inner_shadow_xml(shadow)
            sp_pr.append(sh_xml)
        except Exception:
            pass

    @staticmethod
    def apply_to_text_run(run, shadow: ShadowDef) -> None:
        """Apply shadow to a text run."""
        try:
            r_pr = run._r.find(qn("a:rPr"))
            if r_pr is None:
                from lxml import etree
                r_pr = etree.SubElement(run._r, qn("a:rPr"))
            for old in r_pr.findall(qn("a:outerShdw")):
                r_pr.remove(old)
            for old in r_pr.findall(qn("a:innerShdw")):
                r_pr.remove(old)

            if shadow.shadow_type == "outer":
                sh_xml = ShadowEngine._create_outer_shadow_xml(shadow)
            else:
                sh_xml = ShadowEngine._create_inner_shadow_xml(shadow)
            r_pr.append(sh_xml)
        except Exception:
            pass


# ── Convenience Functions ──

def add_drop_shadow(shape, offset_x: float = 3, offset_y: float = 3, blur: int = 6):
    """Quick drop shadow."""
    ShadowEngine.apply_to_shape(shape, ShadowDef.drop(offset_x=offset_x, offset_y=offset_y, blur=blur))


def add_glow(shape, color: str = "#3B82F6", size: int = 10):
    """Quick glow effect."""
    ShadowEngine.apply_to_shape(shape, ShadowDef.glow(color=color, size=size))


def add_soft_shadow(shape):
    """Quick soft shadow."""
    ShadowEngine.apply_to_shape(shape, ShadowDef.soft())
