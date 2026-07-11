"""Visual engine — gradient, shadow, shapes, and premium layout helpers."""

from src.visual.gradient_engine import GradientEngine, GradientDef, apply_gradient_to_shape, apply_gradient_to_slide
from src.visual.shadow_engine import ShadowEngine, ShadowDef, add_drop_shadow, add_glow, add_soft_shadow
from src.visual.shapes import NeoShape, add_pill_card, add_numbered_circle

__all__ = [
    "GradientEngine", "GradientDef", "apply_gradient_to_shape", "apply_gradient_to_slide",
    "ShadowEngine", "ShadowDef", "add_drop_shadow", "add_glow", "add_soft_shadow",
    "NeoShape", "add_pill_card", "add_numbered_circle",
]
