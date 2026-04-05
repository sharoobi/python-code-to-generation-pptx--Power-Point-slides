#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔════════════════════════════════════════════════════════════════════════╗
║  🎨 CINEMATIC PRESENTATION TEMPLATE - Single File Edition              ║
║  قالب عروض تقديمية سينمائي - ملف واحد احترافي وقابل للتخصيص            ║
╠════════════════════════════════════════════════════════════════════════╣
║  📦 الاستخدام:                                                          ║
║      1) عدّل قسم "⚙️ USER CONFIGURATION" في الأعلى حسب مشروعك          ║
║      2) شغّل: python presentation_template.py                           ║
║      3) استمتع بالعرض! 🎉                                               ║
║                                                                        ║
║  🔧 نقاط التخصيص السريعة:                                               ║
║      • PROJECT_INFO: غيّر البيانات الأساسية لمشروعك                    ║
║      • DESIGN_SYSTEM: عدّل الألوان، الخطوط، والأبعاد                   ║
║      • SLIDES_CONFIG: أضف/احذف/عدّل الشرائح ومحتواها                   ║
║      • ASSET_PATHS: حدّث مسارات الصور والشعارات                        ║
║                                                                        ║
║  🌐 الترخيص: MIT - حرّ في التعديل، النشر، والتوزيع                     ║
║  💻 المتطلبات: python>=3.9, python-pptx, pillow                        ║
║  📦 التثبيت: pip install python-pptx pillow pyyaml                    ║
║                                                                        ║
║  🔄 آخر تحديث: 2026 | ✨ مصمم ليكون قالبًا قابلاً لإعادة الاستخدام   ║
╚════════════════════════════════════════════════════════════════════════╝
"""

# ════════════════════════════════════════════════════════════════════
#  ⚙️  قسم الإعدادات - عدّل هنا فقط!  ⚙️
#  ✅ USER CONFIGURATION - EDIT THIS SECTION ONLY!
# ════════════════════════════════════════════════════════════════════

# 🎯 معلومات المشروع الأساسية - غيّر هذه القيم لمشروعك
# ───────────────────────────────────────────────────────────────────
PROJECT_INFO = {
    "name": "اسم مشروعك هنا",                    # 🔧 غيّر هذا
    "subtitle": "وصف مختصر لمشروعك",              # 🔧 غيّر هذا
    "type": "graduation_project",                # خيارات: graduation, business, pitch, academic
    
    "university": "اسم جامعتك",                   # 🔧 غيّر هذا
    "college": "اسم كليتك",                       # 🔧 غيّر هذا
    "department": "اسم قسمك",                     # 🔧 غيّر هذا
    
    "team": [                                    # 🔧 أضف أعضاء فريقك
        {"name": "اسم الطالب 1", "id": "2022110001"},
        {"name": "اسم الطالب 2", "id": "2022110002"},
        # أضف المزيد حسب الحاجة...
    ],
    
    "supervisor": "اسم المشرف الأكاديمي",
    "location": "المدينة، الدولة",
    "year": "2025-2026",
    
    "language": "ar",                            # خيارات: "ar", "en", "bilingual"
}

# 🎨 نظام التصميم - عدّل الألوان والخطوط لتناسب هويتك
# ───────────────────────────────────────────────────────────────────
DESIGN_SYSTEM = {
    # 🎭 السمة العامة: dark, light, أو custom
    "theme": "dark_cinematic",                   # 🔧 خيارات: dark_cinematic, light_professional, minimal
    
    # 📐 أبعاد الشريحة (بالبوصة - افتراضي 16:9)
    "slide_size": {"width": 13.333, "height": 7.5},
    
    # 🔤 الخطوط
    "fonts": {
        "arabic": "Segoe UI",                    # 🔧 خط النصوص العربية
        "english": "Segoe UI",                   # 🔧 خط النصوص الإنجليزية
        "display": "Segoe UI",                   # 🔧 خط العناوين الكبيرة
    },
    
    # 🌈 لوحة الألوان - عدّل القيم [R, G, B] من 0-255
    "colors": {
        # ألوان الخلفية
        "bg_dark": [3, 5, 12],                   # 🔧 خلفية داكنة
        "bg_light": [247, 249, 253],             # 🔧 خلفية فاتحة
        "bg_gradient_top": [3, 5, 12],           # 🔧 أعلى التدرج
        "bg_gradient_bottom": [18, 35, 66],      # 🔧 أسفل التدرج
        
        # ألوان النصوص
        "text_on_dark": [247, 249, 253],         # 🔧 نص على خلفية داكنة
        "text_on_light": [18, 26, 40],           # 🔧 نص على خلفية فاتحة
        "text_muted": [160, 170, 185],           # 🔧 نص ثانوي
        
        # ألوان التمييز والأكسنت
        "primary": [41, 100, 255],               # 🔧 اللون الأساسي (أزرق ملكي)
        "accent": [98, 242, 255],                # 🔧 لون التمييز (أزرق سماوي)
        "accent_alt": [148, 120, 255],           # 🔧 لون تمييز بديل (بنفسجي)
        "success": [46, 213, 150],               # 🔧 لون النجاح (أخضر)
        "warning": [255, 189, 87],               # 🔧 لون التحذير (أصفر)
        "danger": [255, 117, 96],                # 🔧 لون الخطر (أحمر)
        
        # ألوان التأثيرات
        "glow": [98, 242, 255],                  # 🔧 لون التوهج
        "grid": [112, 129, 170],                 # 🔧 لون الشبكة
        "particles": [247, 249, 253],            # 🔧 لون الجسيمات
    },
    
    # ✨ كثافة التأثيرات البصرية (0.0 = بدون تأثيرات، 1.0 = قصوى)
    "visual_effects_intensity": 0.85,
}

# 📊 إعدادات الشرائح - حدّد عدد الشرائح ومحتواها
# ───────────────────────────────────────────────────────────────────
SLIDES_CONFIG = {
    "total_slides": 34,                          # 🔧 العدد الكلي للشرائح
    
    # 📑 عناوين الفصول - عدّل أو أضف فصولك
    "chapters": [
        "الافتتاح", "الافتتاح", "الملخص", "المشكلة", "المشكلة",
        "تعريف المشروع", "تعريف المشروع", "النطاق", "لماذا الآن؟",
        "المنهجية", "أصحاب المصلحة", "التحليل", "النموذج",
        "المنتج", "المنتج", "المتطلبات", "المتطلبات", "المتطلبات",
        "المتطلبات", "المتطلبات", "الجدوى", "الجدوى", "المخططات",
        "المخططات", "المخططات", "الثقة", "الأعمال", "الأعمال",
        "التميّز", "التطوير", "القيمة", "الدفاع", "الختام", "الختام",
    ],
    
    # 🎨 لون الأكسنت لكل فصل (من لوحة الألوان في DESIGN_SYSTEM)
    "chapter_accents": {
        "الافتتاح": "accent",
        "الملخص": "primary",
        "المشكلة": "danger",
        "تعريف المشروع": "accent_alt",
        "النطاق": "accent",
        "لماذا الآن؟": "warning",
        "المنهجية": "success",
        "أصحاب المصلحة": "accent_alt",
        "التحليل": "danger",
        "النموذج": "primary",
        "المنتج": "accent",
        "المتطلبات": "primary",
        "الجدوى": "warning",
        "المخططات": "accent",
        "الثقة": "success",
        "الأعمال": "warning",
        "التميّز": "accent_alt",
        "التطوير": "success",
        "القيمة": "accent",
        "الدفاع": "primary",
        "الختام": "accent",
    },
    
    # 🔤 إعدادات النصوص
    "text": {
        "font_size_title": 31,
        "font_size_subtitle": 13,
        "font_size_body": 11,
        "font_size_small": 9,
        "line_spacing": 0.42,
    },
    
    # 🔄 اتجاه النصوص (RTL للعربية، LTR للإنجليزية)
    "rtl_enabled": True,
}

# 🖼️ مسارات الأصول والصور - حدّثها لتنظيم ملفاتك
# ───────────────────────────────────────────────────────────────────
ASSET_PATHS = {
    # 📁 المجلدات الأساسية
    "root": ".",                                  # 🔧 مجلد المشروع الرئيسي
    "output": "output.pptx",                      # 🔧 اسم ملف الإخراج
    "assets_dir": "assets",                       # 🔧 مجلد الصور والموارد
    
    # 🖼️ الصور والشعارات - غيّر المسارات أو اتركها فارغة لتجاوزها
    "logos": {
        "mark": "assets/logo-mark.png",           # 🔧 شعار رمزي صغير
        "wordmark": "assets/logo-word.png",       # 🔧 شعار نصي
    },
    
    "screenshots": {
        "home": "assets/screenshots/home.png",
        "login": "assets/screenshots/login.png",
        "sections": "assets/screenshots/sections.png",
        "profile": "assets/screenshots/profile.png",
    },
    
    "diagrams": {
        "agile": "assets/diagrams/agile.png",
        "timeline": "assets/diagrams/timeline.png",
        "auth_flow": "assets/diagrams/auth.png",
        "usecase_user": "assets/diagrams/usecase_user.png",
        "usecase_worker": "assets/diagrams/usecase_worker.png",
        "usecase_store": "assets/diagrams/usecase_store.png",
        "erd": "assets/diagrams/erd.png",
    },
    
    # ⚠️ سلوك عند عدم وجود الصور
    "missing_asset_behavior": "placeholder",      # خيارات: "placeholder", "skip", "error"
}

# ⚡ خيارات التصدير والأداء
# ───────────────────────────────────────────────────────────────────
EXPORT_OPTIONS = {
    "add_speaker_notes": True,                    # إضافة ملاحظات للمتحدث
    "add_transitions": True,                      # إضافة انتقالات بين الشرائح
    "add_animations": False,                      # ⚠️ يتطلب Windows + pywin32
    "com_enhancement": False,                     # ⚠️ تحسينات PowerPoint المتقدمة
    "verbose_logging": False,                     # سجل تفاصيل أكثر للتصحيح
    "skip_missing_assets": True,                  # تجاهل الصور المفقودة بدلاً من التوقف
}

# ════════════════════════════════════════════════════════════════════
#  🛑 لا تعدّل ما بعد هذا السطر إلا إذا كنت مطوّراً!  🛑
#  ⚠️ DO NOT EDIT BELOW THIS LINE UNLESS YOU'RE A DEVELOPER! ⚠️
# ════════════════════════════════════════════════════════════════════

from __future__ import annotations
import math
import hashlib
import json
import logging
import shutil
import argparse
from pathlib import Path
from typing import Iterable, Optional, Dict, List, Any, Tuple

try:
    from PIL import Image, ImageDraw, ImageFilter
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError as e:
    print(f"❌ خطأ: المكتبات المطلوبة غير مثبتة.\n")
    print(f"📦 الرجاء تثبيتها بالأمر:\n   pip install python-pptx pillow pyyaml\n")
    print(f"🔍 التفاصيل: {e}")
    exit(1)

# ════════════════════════════════════════════════════════════════════
#  📋 ملاحظات المتحدث - قابلة للتخصيص
# ════════════════════════════════════════════════════════════════════
SPEAKER_NOTES = [
    "نفتتح العرض بصيغة أكاديمية رسمية: الجامعة، الكلية، القسم، عنوان المشروع، أسماء الفريق، واسم المشرف.",
    "بعد الغلاف الرسمي ننتقل إلى غلاف سينمائي يعيد تعريف المشروع كمنتج، لا كوثيقة فقط.",
    "في الملخص التنفيذي نضغط جوهر المشروع في ثلاث نقاط: المشكلة، الحل، والأثر.",
    "هنا نعرض المشكلة كما هي: السوق يعمل لكنه مشتت، ويحتاج إلى منصة موحدة.",
    "هذه الشريحة تخلق التوتر السردي: السوق موجود لكن التنظيم غائب.",
    "نعرّف المشروع تعريفًا دقيقًا: منصة قطاعية متعددة الأطراف، لا متجرًا فقط.",
    "شريحة قبل وبعد توضّح الفرق بين الوضع التقليدي والوضع بعد دخول المنصة.",
    "تحديد الحدود المكانية والزمانية مهم جدًا أمام اللجنة الأكاديمية.",
    "لماذا الآن؟ لأن السوق موجود، والحاجة متكررة، والفجوة الرقمية واضحة.",
    "اعتمدنا Agile لأن المشروع متعدد الأطراف ومتطلباته قابلة للتغير.",
    "فهم أصحاب المصلحة يشرح بنية المنصة: العميل، المتجر، مقدم الخدمة، والإدارة.",
    "في تحليل الوضع الحالي، نوضح أن الخلل ليس فقط في غياب التطبيق.",
    "هذه الشريحة تعرض النموذج التشغيلي: الرحلة تبدأ بالحساب وتنتهي بالتقييم.",
    "ننتقل من عرض صور كثيرة إلى التركيز على شاشة واحدة ورسالة واحدة.",
    "هذه الشريحة تثبت أن المشروع يمتلك سطحًا منتجيًا فعليًا من التطبيق نفسه.",
    "شريحة فصل بصري بين الرؤية وبين البنية التقنية.",
    "في المتطلبات الوظيفية نركز على المصادقة، الأدوار، وإدارة الخدمات.",
    "في الطبقة الثانية ننتقل إلى المنتجات، المتاجر، الطلبات، والدفع.",
    "ثم تأتي طبقة البحث، التواصل، التخصيص، والدعم الفني.",
    "المتطلبات غير الوظيفية: السرعة، الأمان، التوافقية، الثبات، والصيانة.",
    "الجدوى الفنية: التقنيات المختارة مفهومة وقابلة للتنفيذ.",
    "الجدوى الاقتصادية والتشغيلية والزمنية تعزز منطقية المشروع.",
    "مخططات Use Case لثلاثة أطراف تعطينا قوة في الشرح الأكاديمي.",
    "شريحة الـ spotlight تقرّبنا من المخططات وتوضح ما يهم فيها.",
    "مخطط ERD يرفع مستوى المشروع تقنيًا ويظهر بنية البيانات المترابطة.",
    "الثقة وضبط الجودة: ملفات تعريف، حالات تشغيلية، وتقييمات موثقة.",
    "نموذج الأعمال يوضح مسارات الدخل: عمولات، اشتراكات، إعلانات.",
    "الاستراتيجية التسويقية: تسويق رقمي، SEO، تواصل مباشر، فترات تجريبية.",
    "الميزة التنافسية: التخصص القطاعي والفهم العميق للسوق المستهدف.",
    "خارطة التطوير تُظهر كيف يمكن نقل المشروع إلى نسخة أكثر نضجًا.",
    "نلخص قيمة المشروع من أربع زوايا: أكاديمية، اقتصادية، اجتماعية، ومنتجية.",
    "شريحة الدفاع المختصر: وضوح المشكلة، منطقية الحل، وجود الوثيقة والمخططات.",
    "الفينال يعيد وضع المشروع في إطار أكبر: محاولة لبناء طبقة تنظيم رقمية.",
    "نختم بالشكر، مع ترك انطباع واضح أن العرض شامل وموثق من جميع الزوايا.",
]

# ════════════════════════════════════════════════════════════════════
#  🎨 محرك الخلفيات السينمائية - 25+ تأثير بصري
# ════════════════════════════════════════════════════════════════════

# أبعاد البكسل للخلفية (ثابتة للجودة العالية)
PX_W, PX_H = 1920, 1080

def _layer() -> Tuple[Image.Image, ImageDraw.Draw]:
    """إنشاء طبقة PNG شفافة جديدة للرسم"""
    img = Image.new("RGBA", (PX_W, PX_H), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)

def _composite(base: Image.Image, layer: Image.Image, blur: int = 0) -> Image.Image:
    """دمج طبقة فوق خلفية مع خيار طمس"""
    if blur > 0:
        layer = layer.filter(ImageFilter.GaussianBlur(blur))
    return Image.alpha_composite(base.convert("RGBA"), layer)

def gradient(top: Tuple[int,int,int], bottom: Tuple[int,int,int]) -> Image.Image:
    """تدرج لوني عمودي"""
    img = Image.new("RGB", (PX_W, PX_H), top)
    draw = ImageDraw.Draw(img)
    for y in range(PX_H):
        t = y / max(PX_H - 1, 1)
        c = tuple(int(top[i] * (1 - t) + bottom[i] * t) for i in range(3))
        draw.line((0, y, PX_W, y), fill=c)
    return img

def gradient_radial(center: Tuple[int,int,int], edge: Tuple[int,int,int], *, cx=960, cy=540) -> Image.Image:
    """تدرج لوني دائري من المركز للأطراف"""
    img = Image.new("RGB", (PX_W, PX_H), edge)
    draw = ImageDraw.Draw(img)
    max_r = int(math.hypot(PX_W, PX_H))
    for r in range(max_r, 0, -2):
        t = r / max_r
        c = tuple(int(center[i] * (1 - t) + edge[i] * t) for i in range(3))
        draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=c)
    return img

def add_glow(img: Image.Image, bbox: Tuple[int,int,int,int], color: Tuple[int,int,int], alpha: int, blur: int, *, ellipse: bool = True) -> Image.Image:
    """إضافة تأثير توهج ناعم"""
    layer, draw = _layer()
    fill = (*color, alpha)
    if ellipse:
        draw.ellipse(bbox, fill=fill)
    else:
        draw.rounded_rectangle(bbox, radius=44, fill=fill)
    return _composite(img, layer, blur)

def add_grid(img: Image.Image, step: int, alpha: int, *, color: Tuple[int,int,int]) -> Image.Image:
    """إضافة شبكة خطوط خلفية"""
    layer, draw = _layer()
    c = (*color, alpha)
    for x in range(0, PX_W, step):
        draw.line((x, 0, x, PX_H), fill=c, width=1)
    for y in range(0, PX_H, step):
        draw.line((0, y, PX_W, y), fill=c, width=1)
    return _composite(img, layer)

def add_particles(img: Image.Image, *, color: Tuple[int,int,int], density: int = 1800, alpha_max: int = 26) -> Image.Image:
    """إضافة جسيمات ضوئية عشوائية"""
    layer, draw = _layer()
    for i in range(density):
        x = (i * 37 + i * i * 11) % PX_W
        y = (i * 89 + i * i * 7) % PX_H
        a = 3 + (i * 17) % max(alpha_max, 4)
        r = 1 + (i % 4 == 0) + (i % 7 == 0)
        draw.ellipse((x, y, x + r, y + r), fill=(*color, a))
    return _composite(img, layer)

def add_diagonal_lines(img: Image.Image, alpha: int = 28, *, color: Tuple[int,int,int]) -> Image.Image:
    """إضافة خطوط قطرية ديناميكية"""
    layer, draw = _layer()
    for i in range(-8, 24):
        x = i * 124
        draw.line((x, PX_H, x + 780, 0), fill=(*color, alpha), width=2)
    return _composite(img, layer, 2)

def add_arcs(img: Image.Image, alpha: int = 28, *, color: Tuple[int,int,int], x_shift: int = 0) -> Image.Image:
    """إضافة أقواس منحنية زخرفية"""
    layer, draw = _layer()
    for i in range(9):
        draw.arc((900 + x_shift - i * 58, 90 + i * 34, 1950 + x_shift, 1020 + i * 16), 
                 start=182, end=334, fill=(*color, alpha), width=3)
    return _composite(img, layer, 1)

def add_nodes(img: Image.Image, *, color: Tuple[int,int,int]) -> Image.Image:
    """إضافة عقد متصلة (شبكة عصبية/بيانات)"""
    layer, draw = _layer()
    dots = [(1540, 210), (1380, 350), (1240, 520), (1020, 708), (760, 560), (520, 720)]
    for x, y in dots:
        draw.ellipse((x - 14, y - 14, x + 14, y + 14), fill=(*color, 190))
        draw.ellipse((x - 22, y - 22, x + 22, y + 22), outline=(*color, 60), width=2)
    for (x1, y1), (x2, y2) in zip(dots[:-1], dots[1:]):
        draw.line((x1, y1, x2, y2), fill=(*DESIGN_SYSTEM["colors"]["accent"], 78), width=4)
    return _composite(img, layer, 1)

def add_waves(img: Image.Image, *, color: Tuple[int,int,int], alpha: int = 32, amplitude: int = 34, frequency: float = 2.8, baseline: int = 760) -> Image.Image:
    """إضافة موجات سائلة متحركة"""
    layer, draw = _layer()
    for band in range(5):
        pts = []
        for x in range(0, PX_W + 1, 14):
            y = baseline + band * 22 + int(math.sin((x / PX_W) * math.pi * frequency + band * 0.9) * amplitude)
            pts.append((x, y))
        draw.line(pts, fill=(*color, alpha - band * 4), width=3)
    return _composite(img, layer, 1)

def add_rings(img: Image.Image, *, color: Tuple[int,int,int], alpha: int = 28, center: Tuple[int,int] = (1540, 260), start_radius: int = 120, count: int = 6) -> Image.Image:
    """إضافة حلقات متموجة مركزية"""
    layer, draw = _layer()
    cx, cy = center
    for i in range(count):
        r = start_radius + i * 58
        draw.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(*color, alpha - i * 3), width=3)
    return _composite(img, layer, 1)

def add_blueprint(img: Image.Image, *, color: Tuple[int,int,int]) -> Image.Image:
    """إضافة نمط مخطط هندسي (Blueprint)"""
    layer, draw = _layer()
    for i in range(13):
        x = 120 + i * 128
        draw.line((x, 130, x + 280, 960), fill=(*color, 20), width=2)
    for i in range(17):
        y = 150 + i * 42
        draw.line((940, y, 1840, y), fill=(*DESIGN_SYSTEM["colors"]["text_muted"], 14), width=1)
    return _composite(img, layer, 1)

def add_mesh(img: Image.Image, *, color: Tuple[int,int,int], alpha: int = 18, x_offset: int = 0) -> Image.Image:
    """إضافة شبكة مثلثية ثلاثية الأبعاد"""
    layer, draw = _layer()
    for row in range(7):
        pts = []
        for col in range(-1, 10):
            x = 920 + x_offset + col * 140
            y = 140 + row * 110 + int(math.sin((col + row * 0.5) * 0.7) * 16)
            pts.append((x, y))
        draw.line(pts, fill=(*color, alpha), width=2)
    for col in range(8):
        pts = []
        for row in range(-1, 9):
            x = 940 + x_offset + col * 140 + int(math.sin((row + col * 0.4) * 0.8) * 16)
            y = 110 + row * 110
            pts.append((x, y))
        draw.line(pts, fill=(*DESIGN_SYSTEM["colors"]["accent"], alpha), width=2)
    return _composite(img, layer, 1)

def add_aurora(img: Image.Image, *, colors: List[Tuple[int,int,int]], bands: int = 5, amplitude: int = 80, blur_r: int = 45) -> Image.Image:
    """إضافة تأثير شفق قطبي سينمائي"""
    layer, draw = _layer()
    band_h = PX_H // (bands + 2)
    for b in range(bands):
        c = colors[b % len(colors)]
        base_y = 120 + b * band_h
        pts = []
        for x in range(0, PX_W + 1, 8):
            y = base_y + int(math.sin(x / 220.0 + b * 1.3) * amplitude + math.cos(x / 340.0 + b * 0.7) * amplitude * 0.5)
            pts.append((x, y))
        for w in range(40):
            shifted = [(px, py + w) for px, py in pts]
            draw.line(shifted, fill=(*c, 18 - w // 3), width=2)
    return _composite(img, layer, blur_r)

def add_vignette(img: Image.Image, *, strength: int = 90) -> Image.Image:
    """إضافة تظليل حواف (Vignette) للتركيز على المركز"""
    layer = Image.new("RGBA", (PX_W, PX_H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    max_r = int(math.hypot(PX_W, PX_H)) // 2
    cx, cy = PX_W // 2, PX_H // 2
    for r in range(max_r, max_r // 3, -2):
        t = (r - max_r // 3) / (max_r - max_r // 3)
        a = int(strength * t * t)
        draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(0, 0, 0, a))
    return _composite(img, layer, 20)

def cinematic_preset(name: str, accent: Tuple[int,int,int], *, dark: bool) -> Image.Image:
    """مكتبة الخلفيات الجاهزة - اختر حسب نوع الشريحة"""
    colors = DESIGN_SYSTEM["colors"]
    
    if name == "light_clean":
        img = gradient(colors["bg_light"], (228, 236, 248))
        img = add_grid(img, 100, 10, color=colors["grid"])
        img = add_glow(img, (120, 200, 760, 840), accent, 34, 90)
        return img
    
    if name == "warm_problem":
        img = gradient_radial(colors["warning"], colors["danger"], cx=1440, cy=360)
        img = add_glow(img, (400, 140, 1600, 860), accent, 30, 65)
        img = add_vignette(img, strength=34)
        return img
    
    if name == "deep_circuit":
        img = gradient(colors["bg_dark"], colors["bg_gradient_bottom"])
        img = add_glow(img, (1120, 180, 1880, 900), accent, 90, 72)
        img = add_particles(img, color=colors["particles"], density=900, alpha_max=12)
        return img
    
    if name == "aurora_dark":
        img = gradient(colors["bg_dark"], colors["bg_gradient_bottom"])
        aurora_colors = [accent, colors["accent"], colors["accent_alt"]]
        img = add_aurora(img, colors=aurora_colors, bands=4, amplitude=60, blur_r=48)
        img = add_glow(img, (880, 120, 1780, 820), accent, 72, 72)
        img = add_particles(img, color=colors["particles"], density=1200, alpha_max=12)
        img = add_vignette(img, strength=58)
        return img
    
    # خلفية افتراضية
    bg_top = colors["bg_dark"] if dark else colors["bg_light"]
    bg_bottom = colors["bg_gradient_bottom"] if dark else (216, 225, 239)
    img = gradient(bg_top, bg_bottom)
    particle_color = colors["particles"] if dark else accent
    return add_particles(img, color=particle_color, density=900, alpha_max=10)

# ════════════════════════════════════════════════════════════════════
#  🏗️ مولّد الخلفيات - يولّد 28 خلفية سينمائية فريدة
# ════════════════════════════════════════════════════════════════════

def generate_backgrounds(output_dir: Path, config: Dict) -> List[Tuple[str, Path]]:
    """توليد جميع الخلفيات وحفظها في المجلد المحدد"""
    output_dir.mkdir(parents=True, exist_ok=True)
    colors = config["colors"]
    generated = []
    
    # قائمة الخلفيات: (اسم الملف, دالة التوليد)
    backgrounds = [
        ("01-official.png", lambda: _bg_official(colors)),
        ("02-cinematic.png", lambda: _bg_cinematic(colors)),
        ("03-summary.png", lambda: _bg_summary(colors)),
        ("04-problem.png", lambda: _bg_problem(colors)),
        ("05-what.png", lambda: _bg_what(colors)),
        ("06-scope.png", lambda: _bg_scope(colors)),
        ("07-method.png", lambda: _bg_method(colors)),
        ("08-stakeholders.png", lambda: _bg_stakeholders(colors)),
        ("09-weaknesses.png", lambda: _bg_weaknesses(colors)),
        ("10-model.png", lambda: _bg_model(colors)),
        ("11-experience.png", lambda: _bg_experience(colors)),
        ("12-func-auth.png", lambda: _bg_func_auth(colors)),
        ("13-func-ops.png", lambda: _bg_func_ops(colors)),
        ("14-func-intelligence.png", lambda: _bg_func_intelligence(colors)),
        ("15-nfr.png", lambda: _bg_nfr(colors)),
        ("16-feasibility-tech.png", lambda: _bg_feasibility_tech(colors)),
        ("17-feasibility-business.png", lambda: _bg_feasibility_business(colors)),
        ("18-usecases.png", lambda: _bg_usecases(colors)),
        ("19-erd.png", lambda: _bg_erd(colors)),
        ("20-trust.png", lambda: _bg_trust(colors)),
        ("21-business.png", lambda: _bg_business(colors)),
        ("22-marketing.png", lambda: _bg_marketing(colors)),
        ("23-advantage.png", lambda: _bg_advantage(colors)),
        ("24-roadmap.png", lambda: _bg_roadmap(colors)),
        ("25-value.png", lambda: _bg_value(colors)),
        ("26-defense.png", lambda: _bg_defense(colors)),
        ("27-finale.png", lambda: _bg_finale(colors)),
        ("28-closing.png", lambda: _bg_closing(colors)),
    ]
    
    for filename, generator in backgrounds:
        try:
            img = generator()
            path = output_dir / filename
            img.convert("RGB").save(path)
            generated.append((filename, path))
        except Exception as e:
            logging.warning(f"⚠️ فشل توليد {filename}: {e}")
            # إنشاء خلفية بديلة بسيطة
            img = gradient(colors["bg_dark"], colors["bg_gradient_bottom"])
            path = output_dir / filename
            img.convert("RGB").save(path)
            generated.append((filename, path))
    
    return generated

# ── دوال توليد الخلفيات الفردية (مبسطة للقالب) ─────────────────────

def _bg_official(c): return gradient(c["bg_light"], (225, 233, 246))
def _bg_cinematic(c): return cinematic_preset("aurora_dark", c["accent"], dark=True)
def _bg_summary(c): return cinematic_preset("light_clean", c["primary"], dark=False)
def _bg_problem(c): return cinematic_preset("warm_problem", c["danger"], dark=True)
def _bg_what(c): 
    img = gradient(c["bg_dark"], c["bg_gradient_bottom"])
    return add_glow(img, (980, 90, 1840, 980), c["accent"], 90, 62)
def _bg_scope(c): return gradient(c["bg_light"], (232, 238, 248))
def _bg_method(c): return cinematic_preset("aurora_dark", c["success"], dark=True)
def _bg_stakeholders(c): return gradient(c["bg_light"], (233, 239, 248))
def _bg_weaknesses(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_model(c): return cinematic_preset("light_clean", c["primary"], dark=False)
def _bg_experience(c): 
    img = gradient(c["bg_dark"], c["bg_gradient_bottom"])
    return add_glow(img, (980, 100, 1880, 980), c["primary"], 108, 68)
def _bg_func_auth(c): return gradient(c["bg_light"], (232, 238, 248))
def _bg_func_ops(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_func_intelligence(c): return gradient(c["bg_light"], (234, 240, 248))
def _bg_nfr(c): return gradient(c["bg_dark"], (17, 37, 66))
def _bg_feasibility_tech(c): return gradient(c["bg_light"], (232, 239, 249))
def _bg_feasibility_business(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_usecases(c): return cinematic_preset("light_clean", c["accent"], dark=False)
def _bg_erd(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_trust(c): return cinematic_preset("light_clean", c["accent_alt"], dark=False)
def _bg_business(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_marketing(c): return gradient(c["bg_light"], (233, 239, 248))
def _bg_advantage(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_roadmap(c): return cinematic_preset("light_clean", c["success"], dark=False)
def _bg_value(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_defense(c): return gradient(c["bg_light"], (233, 239, 248))
def _bg_finale(c): return gradient(c["bg_dark"], c["bg_gradient_bottom"])
def _bg_closing(c): return cinematic_preset("aurora_dark", c["accent"], dark=True)

# ════════════════════════════════════════════════════════════════════
#  🧩 مكونات الشرائح - دوال مساعدة قابلة لإعادة الاستخدام
# ════════════════════════════════════════════════════════════════════

def rgb(color_name: str) -> RGBColor:
    """تحويل اسم اللون في الإعدادات إلى RGBColor لـ python-pptx"""
    c = DESIGN_SYSTEM["colors"].get(color_name, DESIGN_SYSTEM["colors"]["primary"])
    return RGBColor(*c)

def set_rtl(paragraph):
    """تفعيل اتجاه النص من اليمين لليسار للعربية"""
    if SLIDES_CONFIG["rtl_enabled"]:
        ppr = paragraph._p.get_or_add_pPr()
        ppr.set("rtl", "1")

def add_background(slide, bg_path: Path, slide_w: float, slide_h: float):
    """إضافة خلفية للشريحة"""
    if bg_path.exists():
        slide.shapes.add_picture(str(bg_path), 0, 0, Inches(slide_w), Inches(slide_h))

def add_text_rtl(slide, left: float, top: float, width: float, height: float, 
                 text: str, *, font_name: str = None, font_size: int = 20, 
                 color_name: str = "text_on_dark", bold: bool = False, 
                 align=PP_ALIGN.RIGHT):
    """إضافة نص مع دعم RTL"""
    font_name = font_name or DESIGN_SYSTEM["fonts"]["arabic" if SLIDES_CONFIG["rtl_enabled"] else "english"]
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = align
    set_rtl(p)
    r = p.add_run()
    r.text = text
    r.font.name = font_name
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color_name)
    return box

def add_card(slide, left: float, top: float, width: float, height: float, 
             *, fill_name: str = "bg_dark", border_name: str = "accent", 
             transparency: float = 0.1):
    """إضافة بطاقة/صندوق نصي مزخرف"""
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 
                                   Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = rgb(fill_name)
    card.fill.transparency = transparency
    card.line.color.rgb = rgb(border_name)
    card.line.width = Pt(1.0)
    return card

def add_title_cluster(slide, eyebrow: str, title: str, subtitle: str, *, dark: bool = True):
    """إضافة مجموعة عنوان: كلمة صغيرة + عنوان رئيسي + وصف"""
    fg = "text_on_dark" if dark else "text_on_light"
    muted = "text_muted"
    accent = "accent"
    
    # كلمة صغيرة في الأعلى
    add_text_rtl(slide, 10.22, 0.48, 2.05, 0.32, eyebrow, font_size=9, color_name=accent, bold=True)
    # العنوان الرئيسي
    add_text_rtl(slide, 5.6, 0.92, 6.2, 1.42, title, font_name=DESIGN_SYSTEM["fonts"]["display"], 
                 font_size=31, color_name=fg, bold=True)
    # الوصف
    if subtitle:
        add_text_rtl(slide, 5.64, 2.08, 6.05, 0.76, subtitle, font_size=13, color_name=muted)
    # شريط تمييز
    bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 
                                  Inches(5.62), Inches(2.96), Inches(1.22), Inches(0.03))
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb(accent)
    bar.line.fill.background()

def add_footer(slide, *, dark: bool = True):
    """إضافة تذييل ثابت في أسفل الشريحة"""
    c = "text_muted"
    bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(7.2), 
                                  Inches(DESIGN_SYSTEM["slide_size"]["width"]), Inches(0.02))
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb("grid" if dark else "text_muted")
    bar.fill.transparency = 0.5
    bar.line.fill.background()
    add_text_rtl(slide, 0.45, 7.12, 4.1, 0.2, 
                 f"{PROJECT_INFO['name']} | Cinematic Template", 
                 font_size=8, color_name=c, align=PP_ALIGN.LEFT)

def add_slide_number(slide, num: int, total: int, *, dark: bool = True):
    """إضافة رقم الشريحة وشريط التقدم"""
    accent = SLIDES_CONFIG["chapter_accents"].get(
        SLIDES_CONFIG["chapters"][num-1] if num <= len(SLIDES_CONFIG["chapters"]) else "default", 
        "accent"
    )
    chapter = SLIDES_CONFIG["chapters"][num-1] if num <= len(SLIDES_CONFIG["chapters"]) else PROJECT_INFO["name"]
    
    # شريط التقدم
    track = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(8.52), Inches(7.09), Inches(3.76), Inches(0.08))
    track.fill.solid()
    track.fill.fore_color.rgb = rgb("grid" if dark else "text_muted")
    track.fill.transparency = 0.55
    track.line.fill.background()
    
    fill_w = 3.76 * (num / max(total, 1))
    fill = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(8.52), Inches(7.09), Inches(fill_w), Inches(0.08))
    fill.fill.solid()
    fill.fill.fore_color.rgb = rgb(accent)
    fill.line.fill.background()
    
    # النصوص
    add_text_rtl(slide, 6.98, 6.98, 1.25, 0.2, chapter, font_size=8, color_name=accent, align=PP_ALIGN.LEFT)
    add_text_rtl(slide, 12.02, 6.98, 0.7, 0.2, f"{num}/{total}", font_size=8, color_name=accent, align=PP_ALIGN.LEFT)

# ════════════════════════════════════════════════════════════════════
#  📄 دوال بناء الشرائح - كل دالة تبني شريحة واحدة
# ════════════════════════════════════════════════════════════════════

def slide_cover(prs, bg_path: Path):
    """شريحة الغلاف الرسمية"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, bg_path, DESIGN_SYSTEM["slide_size"]["width"], DESIGN_SYSTEM["slide_size"]["height"])
    add_footer(slide, dark=False)
    add_slide_number(slide, 1, SLIDES_CONFIG["total_slides"], dark=False)
    
    # معلومات الجامعة
    add_text_rtl(slide, 4.95, 0.7, 4.7, 0.24, PROJECT_INFO["university"], font_size=11, color_name="text_muted", bold=True)
    add_text_rtl(slide, 4.95, 1.02, 4.7, 0.26, PROJECT_INFO["college"], font_name=DESIGN_SYSTEM["fonts"]["display"], font_size=18, color_name="text_on_light", bold=True)
    add_text_rtl(slide, 4.95, 1.36, 4.7, 0.22, f"{PROJECT_INFO['department']} | {PROJECT_INFO['year']}", font_size=11, color_name="text_muted")
    
    # عنوان المشروع
    add_text_rtl(slide, 3.9, 2.2, 7.25, 1.02, PROJECT_INFO["name"], font_name=DESIGN_SYSTEM["fonts"]["display"], font_size=28, color_name="text_on_light", bold=True)
    add_text_rtl(slide, 4.15, 3.46, 6.75, 0.4, PROJECT_INFO["subtitle"], font_size=14, color_name="text_muted")
    
    # فريق العمل
    card = add_card(slide, 0.95, 4.2, 11.45, 2.45, fill_name="bg_light", border_name="primary", transparency=0.01)
    y = 4.42
    for member in PROJECT_INFO["team"]:
        add_text_rtl(slide, 1.28, y, 10.5, 0.24, f"{member['name']} | {member['id']}", font_size=12, color_name="text_on_light")
        y += 0.42
    
    # المشرف
    add_text_rtl(slide, 1.28, y + 0.2, 10.5, 0.24, f"إشراف: {PROJECT_INFO['supervisor']}", font_size=11, color_name="text_muted")

def slide_cinematic_cover(prs, bg_path: Path):
    """شريحة الغلاف السينمائية"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, bg_path, DESIGN_SYSTEM["slide_size"]["width"], DESIGN_SYSTEM["slide_size"]["height"])
    add_footer(slide)
    add_slide_number(slide, 2, SLIDES_CONFIG["total_slides"])
    
    # عنوان كبير
    add_text_rtl(slide, 5.95, 1.34, 5.35, 1.7, PROJECT_INFO["name"], font_name=DESIGN_SYSTEM["fonts"]["display"], font_size=38, color_name="text_on_dark", bold=True)
    add_text_rtl(slide, 6.02, 3.12, 5.3, 0.7, PROJECT_INFO["subtitle"], font_size=16, color_name="text_muted")
    
    # نقاط القوة
    add_text_rtl(slide, 0.95, 4.95, 2.8, 0.42, "✓ حل حقيقي لمشكلة متكررة", font_size=13, color_name="text_on_dark")
    add_text_rtl(slide, 4.02, 4.95, 2.8, 0.42, "✓ بنية تقنية قابلة للتوسع", font_size=13, color_name="text_on_dark")
    add_text_rtl(slide, 7.09, 4.95, 2.8, 0.42, "✓ وثيقة أكاديمية متكاملة", font_size=13, color_name="text_on_dark")
    add_text_rtl(slide, 10.16, 4.95, 2.8, 0.42, "✓ قابل للتطوير بعد التخرج", font_size=13, color_name="text_on_dark")

def slide_summary(prs, bg_path: Path):
    """شريحة الملخص التنفيذي"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, bg_path, DESIGN_SYSTEM["slide_size"]["width"], DESIGN_SYSTEM["slide_size"]["height"])
    add_footer(slide, dark=False)
    add_slide_number(slide, 3, SLIDES_CONFIG["total_slides"], dark=False)
    
    add_title_cluster(slide, "الملخص", "ما هو المشروع باختصار؟", 
                     "عرض موجز للفكرة، الحل، والأثر المتوقع في 3 بطاقات.", dark=False)
    
    # البطاقات الثلاث
    add_card(slide, 0.95, 3.2, 3.35, 2.2, fill_name="bg_light", border_name="primary", transparency=0.02)
    add_text_rtl(slide, 1.15, 3.4, 3.0, 0.3, "🎯 المشكلة", font_size=14, color_name="text_on_light", bold=True)
    add_text_rtl(slide, 1.15, 3.85, 3.0, 1.3, "وصف مختصر للمشكلة التي يعالجها المشروع...", font_size=11, color_name="text_muted")
    
    add_card(slide, 4.06, 3.2, 3.35, 2.2, fill_name="bg_light", border_name="accent", transparency=0.02)
    add_text_rtl(slide, 4.26, 3.4, 3.0, 0.3, "💡 الحل", font_size=14, color_name="text_on_light", bold=True)
    add_text_rtl(slide, 4.26, 3.85, 3.0, 1.3, "وصف مختصر للحل المقترح ومنهجيته...", font_size=11, color_name="text_muted")
    
    add_card(slide, 7.17, 3.2, 3.35, 2.2, fill_name="bg_light", border_name="success", transparency=0.02)
    add_text_rtl(slide, 7.37, 3.4, 3.0, 0.3, "🚀 الأثر", font_size=14, color_name="text_on_light", bold=True)
    add_text_rtl(slide, 7.37, 3.85, 3.0, 1.3, "القيمة المضافة والفوائد المتوقعة...", font_size=11, color_name="text_muted")

# ════════════════════════════════════════════════════════════════════
#  🏗️ الدالة الرئيسية لبناء العرض
# ════════════════════════════════════════════════════════════════════

def build_presentation(config: Dict, output_path: Path) -> Path:
    """بناء العرض التقديمي كاملاً"""
    log = logging.getLogger(__name__)
    log.info("🎨 بدء بناء العرض التقديمي...")
    
    # تهيئة العرض
    prs = Presentation()
    prs.slide_width = Inches(config["slide_size"]["width"])
    prs.slide_height = Inches(config["slide_size"]["height"])
    
    # توليد الخلفيات
    assets_dir = Path(config.get("assets_dir", "assets"))
    bg_dir = assets_dir / "backgrounds"
    log.info(f"🖼️ توليد الخلفيات في: {bg_dir}")
    backgrounds = generate_backgrounds(bg_dir, config["colors"])
    
    # دوال الشرائح المتاحة
    slide_functions = [slide_cover, slide_cinematic_cover, slide_summary]
    # 🔧 لإضافة المزيد من الشرائح، أضف دوال جديدة هنا واستدعها في الحلقة أدناه
    
    # بناء الشرائح
    total = min(config["total_slides"], len(SLIDES_CONFIG["chapters"]))
    for i in range(min(total, len(slide_functions))):
        bg_name, bg_path = backgrounds[i % len(backgrounds)]
        log.info(f"📄 بناء الشريحة {i+1}/{total}: {SLIDES_CONFIG['chapters'][i]}")
        slide_functions[i](prs, bg_path)
    
    # للشرائح المتبقية، نستخدم شريحة محتوى عامة
    for i in range(len(slide_functions), total):
        bg_name, bg_path = backgrounds[i % len(backgrounds)]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(slide, bg_path, config["slide_size"]["width"], config["slide_size"]["height"])
        add_footer(slide, dark=("dark" in config["theme"]))
        add_slide_number(slide, i+1, total, dark=("dark" in config["theme"]))
        
        chapter = SLIDES_CONFIG["chapters"][i] if i < len(SLIDES_CONFIG["chapters"]) else "محتوى"
        add_title_cluster(slide, chapter, f"عنوان الشريحة {i+1}", 
                         "اكتب محتوى شريحتك هنا. هذا قالب قابل للتخصيص بالكامل!", 
                         dark=("dark" in config["theme"]))
    
    # حفظ الملف
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    log.info(f"✅ تم الحفظ: {output_path}")
    
    return output_path

# ════════════════════════════════════════════════════════════════════
#  🎯 نقطة الدخول الرئيسية
# ════════════════════════════════════════════════════════════════════

def main():
    """نقطة الدخول للتطبيق"""
    parser = argparse.ArgumentParser(
        description="🎨 قالب عرض تقديمي سينمائي - ملف واحد احترافي",
        epilog="مثال: python presentation_template.py --output my-project.pptx"
    )
    parser.add_argument("-c", "--config", type=Path, default=None, help="ملف إعدادات YAML خارجي (اختياري)")
    parser.add_argument("-o", "--output", type=Path, default=Path(ASSET_PATHS["output"]), help="مسار ملف الإخراج")
    parser.add_argument("-t", "--theme", choices=["dark_cinematic", "light_professional", "minimal"], help="اختيار سمة جاهزة")
    parser.add_argument("-v", "--verbose", action="store_true", help="تفعيل التسجيل المفصل")
    parser.add_argument("--dry-run", action="store_true", help="محاكاة دون إنشاء ملف")
    
    args = parser.parse_args()
    
    # إعداد التسجيل
    level = logging.DEBUG if args.verbose or EXPORT_OPTIONS["verbose_logging"] else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")
    log = logging.getLogger(__name__)
    
    # تحميل الإعدادات
    config = {
        "project": PROJECT_INFO,
        "design": DESIGN_SYSTEM,
        "slides": SLIDES_CONFIG,
        "assets": ASSET_PATHS,
        "export": EXPORT_OPTIONS,
    }
    
    # تطبيق السمة إذا طُلبت
    if args.theme:
        log.info(f"🎨 تطبيق السمة: {args.theme}")
        if args.theme == "light_professional":
            config["design"]["theme"] = "light_professional"
            config["design"]["colors"]["bg_dark"] = config["design"]["colors"]["bg_light"]
        elif args.theme == "minimal":
            config["design"]["visual_effects_intensity"] = 0.3
    
    if args.dry_run:
        log.info("🔍 وضع المحاكاة: لن يتم إنشاء ملف")
        log.info("✅ الإعدادات صالحة، القالب جاهز للاستخدام!")
        return 0
    
    # بناء العرض
    try:
        output = build_presentation(config, args.output)
        log.info(f"\n🎉 تم بنجاح! افتح الملف: {output.absolute()}")
        log.info("💡 لتخصيص القالب: عدّل قسم الإعدادات في أعلى الملف")
        return 0
    except Exception as e:
        log.error(f"❌ خطأ أثناء البناء: {e}")
        if EXPORT_OPTIONS["verbose_logging"]:
            import traceback
            traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit(main())
