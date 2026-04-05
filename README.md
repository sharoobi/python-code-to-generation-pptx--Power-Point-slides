# 🎨 Cinematic Presentation Template

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9+-blue?style=for-the-badge&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License">
  <img src="https://img.shields.io/badge/Arabic-RTL-orange?style=for-the-badge" alt="Arabic Support">
  <img src="https://img.shields.io/badge/Single_File-✅-purple?style=for-the-badge" alt="Single File">
</p>

<p align="center">
  <strong>قالب عروض تقديمية احترافي في ملف بايثون واحد!</strong><br>
  <em>سينمائي • قابل للتخصيص • يدعم العربية • جاهز للنشر</em>
</p>

<p align="center">
  <a href="#-بدء-الاستخدام-السريع">🚀 بدء سريع</a> •
  <a href="#⚙️-تخصيص-القالب">⚙️ التخصيص</a> •
  <a href="#📚-مرجع-الإعدادات">📚 المرجع</a> •
  <a href="#❓-الأسئلة-الشائعة">❓ الأسئلة</a>
</p>

---

## ✨ لماذا هذا القالب؟

| الميزة | الفائدة لك |
|--------|-----------|
| 📦 **ملف واحد** | لا حاجة لمجلدات معقدة، كل شيء في مكان واحد |
| ⚙️ **إعدادات في الأعلى** | عدّل مشروعك من قسم واحد واضح، لا تلمس الكود |
| 🌙🌞 **سمات متعددة** | اختر بين الوضع الداكن، الفاتح، أو المينيمال |
| 🕌 **دعم عربي كامل** | نصوص من اليمين لليسار، خطوط عربية، تواريخ هجرية |
| 🎬 **خلفيات سينمائية** | 28 خلفية جاهزة بتأثيرات بصرية احترافية |
| 🔧 **قابل للتوسع** | أضف شرائح، ألوان، تأثيرات بسهولة تامة |
| 🛡️ **آمن عند الخطأ** | يتجاهل الصور المفقودة ولا يتوقف فجأة |

---

## 🚀 بدء الاستخدام السريع

### الخطوة 1: التثبيت

```bash
# 1️⃣ أنشئ مجلدًا جديدًا لمشروعك
mkdir my-presentation && cd my-presentation

# 2️⃣ حمّل ملف القالب
#    انسخ محتوى presentation_template.py وضعه في مجلدك

# 3️⃣ ثبّت المكتبات المطلوبة
pip install python-pptx pillow pyyaml

# ✅ جاهز! جرب التشغيل:
python presentation_template.py
```

### الخطوة 2: التخصيص (5 دقائق فقط!)

افتح الملف وعدّل **فقط هذا القسم في الأعلى**:

```python
# ════════════════════════════════════════════════════════════
#  ⚙️  قسم الإعدادات - عدّل هنا فقط!  ⚙️
# ════════════════════════════════════════════════════════════

# 🎯 معلومات المشروع
PROJECT_INFO = {
    "name": "اسم مشروعك هنا",              # ← غيّر هذا
    "subtitle": "وصف مختصر لمشروعك",        # ← وهذا
    "team": [
        {"name": "أحمد محمد", "id": "2025001"},  # ← أضف فريقك
        {"name": "سارة علي", "id": "2025002"},
    ],
    "supervisor": "د. خالد أحمد",
    "university": "جامعة ...",
    # ... بقية الحقول
}

# 🎨 الألوان والخطوط
DESIGN_SYSTEM = {
    "theme": "dark_cinematic",  # خيارات: dark, light, minimal
    "colors": {
        "primary": [41, 100, 255],   # ← غيّر اللون الأساسي [R, G, B]
        "accent": [98, 242, 255],    # ← لون التمييز
        # ... أضف ألوانك الخاصة
    },
}

# 🖼️ مسارات الصور
ASSET_PATHS = {
    "logos": {
        "mark": "assets/my-logo.png",  # ← مسار شعارك
        "wordmark": "assets/logo-text.png",
    },
    # ... بقية المسارات
}
```

### الخطوة 3: التشغيل

```bash
# تشغيل بسيط
python presentation_template.py

# مع خيارات متقدمة
python presentation_template.py \
    --output عرضي_النهائي.pptx \
    --theme light_professional \
    --verbose

# محاكاة دون إنشاء ملف (للتجربة)
python presentation_template.py --dry-run
```

### الخطوة 4: الاستمتاع! 🎉

سيظهر ملف `output.pptx` (أو الاسم الذي حددته) جاهز للعرض أو التعديل في PowerPoint.

---

## ⚙️ تخصيص القالب

### 🎨 تغيير السمة (Theme)

```bash
# السمة الداكنة السينمائية (افتراضية)
python presentation_template.py --theme dark_cinematic

# السمة الفاتحة الاحترافية
python presentation_template.py --theme light_professional

# السمة المينيمال البسيطة
python presentation_template.py --theme minimal
```

أو عدّل في الكود مباشرة:
```python
DESIGN_SYSTEM = {
    "theme": "light_professional",  # ← غيّر هنا
    # ...
}
```

### 🌈 تغيير الألوان

كل الألوان في قسم `DESIGN_SYSTEM["colors"]`. القيم هي [أحمر، أخضر، أزرق] من 0-255:

```python
"colors": {
    "primary": [255, 100, 100],      # أحمر بدلاً من أزرق
    "accent": [255, 215, 0],         # ذهبي
    "bg_dark": [10, 10, 20],         # خلفية داكنة مخصصة
    "text_on_dark": [255, 255, 255], # نص أبيض نقي
}
```

> 💡 **نصيحة**: استخدم موقع مثل [rgb.to](https://rgb.to/) لاختيار ألوانك وتحويلها لقيم RGB.

### 🔤 تغيير الخطوط

```python
"fonts": {
    "arabic": "Segoe UI",        # الخط الافتراضي للعربية
    "english": "Segoe UI",       # الخط الافتراضي للإنجليزية
    "display": "Tahoma",         # خط العناوين الكبيرة
},
```

> ⚠️ تأكد أن الخطوط مثبتة على جهازك، أو استخدم خطوطًا شائعة مثل: `Arial`, `Tahoma`, `Segoe UI`, `Calibri`.

### 🖼️ إضافة شعارك وصور مشروعك

1. أنشئ مجلد `assets/` بجانب الملف:
```
my-presentation/
├── presentation_template.py
└── assets/
    ├── logos/
    │   ├── my-logo.png
    │   └── logo-text.png
    └── screenshots/
        ├── home.png
        └── dashboard.png
```

2. حدّث المسارات في `ASSET_PATHS`:
```python
ASSET_PATHS = {
    "logos": {
        "mark": "assets/logos/my-logo.png",
        "wordmark": "assets/logos/logo-text.png",
    },
    "screenshots": {
        "home": "assets/screenshots/home.png",
        # ...
    },
}
```

### 📊 إضافة/تعديل الشرائح

#### لإضافة شريحة جديدة:

1. أنشئ دالة جديدة بنفس النمط:
```python
def slide_my_custom(prs, bg_path: Path):
    """شريحة مخصصة لمحتواك"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, bg_path, 
                   DESIGN_SYSTEM["slide_size"]["width"], 
                   DESIGN_SYSTEM["slide_size"]["height"])
    add_footer(slide, dark=("dark" in DESIGN_SYSTEM["theme"]))
    add_slide_number(slide, 35, SLIDES_CONFIG["total_slides"])
    
    # أضف محتواك هنا
    add_text_rtl(slide, 1.0, 1.0, 11.0, 1.0, 
                 "عنوان شريحتك هنا", 
                 font_name=DESIGN_SYSTEM["fonts"]["display"],
                 font_size=28, 
                 color_name="text_on_dark", 
                 bold=True)
    
    add_text_rtl(slide, 1.0, 2.5, 11.0, 4.0,
                 "اكتب محتوى شريحتك هنا...\n"
                 "• نقطة أولى\n"
                 "• نقطة ثانية\n"
                 "• نقطة ثالثة",
                 font_size=14,
                 color_name="text_muted")
```

2. أضف الدالة إلى القائمة في `build_presentation()`:
```python
slide_functions = [
    slide_cover, 
    slide_cinematic_cover, 
    slide_summary,
    # ... الشرائح الأخرى
    slide_my_custom,  # ← أضف دالتك هنا
]
```

3. زِد عدد الشرائح في الإعدادات:
```python
SLIDES_CONFIG = {
    "total_slides": 35,  # ← زِد الرقم
    # ...
}
```

#### لتعديل محتوى شريحة موجودة:

ابحث عن الدالة المقابلة (مثلاً `slide_summary`) وعدّل النصوص داخلها:
```python
def slide_summary(prs, bg_path: Path):
    # ...
    add_text_rtl(slide, 1.15, 3.85, 3.0, 1.3, 
                 "✏️ اكتب وصف مشكلتك هنا...",  # ← عدّل هذا النص
                 font_size=11, color_name="text_muted")
```

---

## 📚 مرجع الإعدادات

### 🎯 `PROJECT_INFO` - معلومات المشروع

| الحقل | النوع | الوصف | مثال |
|-------|-------|-------|--------|
| `name` | str | اسم المشروع | `"منصة سين للخدمات"` |
| `subtitle` | str | وصف مختصر | `"ربط السوق بالخدمة والثقة"` |
| `type` | str | نوع المشروع | `"graduation_project"` |
| `university` | str | اسم الجامعة | `"جامعة أزال"` |
| `college` | str | اسم الكلية | `"كلية الحاسوب"` |
| `department` | str | اسم القسم | `"قسم تكنولوجيا المعلومات"` |
| `team` | list[dict] | أعضاء الفريق | `[{"name": "أحمد", "id": "2025001"}]` |
| `supervisor` | str | اسم المشرف | `"د. مختار غيلان"` |
| `location` | str | الموقع الجغرافي | `"صنعاء، اليمن"` |
| `year` | str | السنة الدراسية | `"2025-2026"` |
| `language` | str | لغة العرض | `"ar"` أو `"en"` أو `"bilingual"` |

### 🎨 `DESIGN_SYSTEM` - نظام التصميم

#### السمة العامة
```python
"theme": "dark_cinematic"  # خيارات: dark_cinematic, light_professional, minimal
```

#### أبعاد الشريحة
```python
"slide_size": {"width": 13.333, "height": 7.5}  # بالبوصة (16:9 افتراضي)
```

#### الخطوط
```python
"fonts": {
    "arabic": "Segoe UI",      # خط النصوص العربية
    "english": "Segoe UI",     # خط النصوص الإنجليزية  
    "display": "Segoe UI",     # خط العناوين الكبيرة
},
```

#### لوحة الألوان (كل القيم [R, G, B] من 0-255)
```python
"colors": {
    # الخلفيات
    "bg_dark": [3, 5, 12],           # خلفية الوضع الداكن
    "bg_light": [247, 249, 253],     # خلفية الوضع الفاتح
    "bg_gradient_top": [3, 5, 12],   # أعلى التدرج اللوني
    "bg_gradient_bottom": [18, 35, 66], # أسفل التدرج
    
    # النصوص
    "text_on_dark": [247, 249, 253], # نص على خلفية داكنة
    "text_on_light": [18, 26, 40],   # نص على خلفية فاتحة
    "text_muted": [160, 170, 185],   # نص ثانوي/باهت
    
    # ألوان التمييز
    "primary": [41, 100, 255],       # اللون الأساسي
    "accent": [98, 242, 255],        # لون التمييز الرئيسي
    "accent_alt": [148, 120, 255],   # لون تمييز بديل
    "success": [46, 213, 150],       # لون النجاح (أخضر)
    "warning": [255, 189, 87],       # لون التحذير (أصفر)
    "danger": [255, 117, 96],        # لون الخطر (أحمر)
    
    # التأثيرات البصرية
    "glow": [98, 242, 255],          # لون التوهج
    "grid": [112, 129, 170],         # لون الشبكة الخلفية
    "particles": [247, 249, 253],    # لون الجسيمات الضوئية
},
```

#### كثافة التأثيرات
```python
"visual_effects_intensity": 0.85  # من 0.0 (بدون تأثيرات) إلى 1.0 (قصوى)
```

### 📊 `SLIDES_CONFIG` - إعدادات الشرائح

```python
SLIDES_CONFIG = {
    "total_slides": 34,  # العدد الكلي للشرائح
    
    # عناوين الفصول (تظهر في شريط التقدم)
    "chapters": [
        "الافتتاح", "الملخص", "المشكلة", "الحل",  # ← عدّل أو أضف فصولك
        # ...
    ],
    
    # لون الأكسنت لكل فصل (من أسماء الألوان في DESIGN_SYSTEM["colors"])
    "chapter_accents": {
        "الافتتاح": "accent",
        "الملخص": "primary", 
        "المشكلة": "danger",
        # ...
    },
    
    # أحجام النصوص
    "text": {
        "font_size_title": 31,    # حجم عنوان الشريحة
        "font_size_subtitle": 13, # حجم الوصف
        "font_size_body": 11,     # حجم النص العادي
        "font_size_small": 9,     # حجم النصوص الصغيرة
        "line_spacing": 0.42,     # تباعد الأسطر
    },
    
    # دعم اللغة العربية
    "rtl_enabled": True,  # True للنصوص من اليمين لليسار
}
```

### 🖼️ `ASSET_PATHS` - مسارات الصور والموارد

```python
ASSET_PATHS = {
    "root": ".",                          # مجلد المشروع الرئيسي
    "output": "output.pptx",              # اسم ملف الإخراج
    "assets_dir": "assets",               # مجلد الصور
    
    "logos": {
        "mark": "assets/logo-mark.png",   # شعار رمزي صغير
        "wordmark": "assets/logo-word.png", # شعار نصي
    },
    
    "screenshots": {
        "home": "assets/screenshots/home.png",      # لقطة الشاشة الرئيسية
        "login": "assets/screenshots/login.png",    # لقطة تسجيل الدخول
        # ... أضف لقطاتك
    },
    
    "diagrams": {
        "agile": "assets/diagrams/agile.png",       # مخطط المنهجية
        "erd": "assets/diagrams/erd.png",           # مخطط قواعد البيانات
        # ... أضف مخططاتك
    },
    
    # سلوك عند عدم وجود الصور
    "missing_asset_behavior": "placeholder",  # خيارات: "placeholder", "skip", "error"
}
```

### ⚡ `EXPORT_OPTIONS` - خيارات التصدير

```python
EXPORT_OPTIONS = {
    "add_speaker_notes": True,      # إضافة ملاحظات للمتحدث أسفل كل شريحة
    "add_transitions": True,        # إضافة انتقالات بين الشرائح
    "add_animations": False,        # ⚠️ يتطلب Windows + pywin32
    "com_enhancement": False,       # ⚠️ تحسينات PowerPoint المتقدمة (Windows فقط)
    "verbose_logging": False,       # تسجيل تفاصيل أكثر للتصحيح
    "skip_missing_assets": True,    # تجاهل الصور المفقودة بدلاً من التوقف
}
```

---

## 🖥️ خيارات سطر الأوامر

```bash
python presentation_template.py [خيارات]

الخيارات المتاحة:
  -c, --config PATH     ملف إعدادات YAML خارجي (اختياري)
  -o, --output PATH     مسار ملف الإخراج (افتراضي: output.pptx)
  -t, --theme NAME      اختيار سمة جاهزة: dark_cinematic, light_professional, minimal
  -v, --verbose         تفعيل التسجيل المفصل للتصحيح
  --dry-run             محاكاة البناء دون إنشاء ملف فعلي
  --help                عرض رسالة المساعدة والخروج
```

### أمثلة عملية:

```bash
# 🎯 تشغيل أساسي
python presentation_template.py

# 📁 تغيير اسم ملف الإخراج
python presentation_template.py --output "عرض_مشروعي_النهائي.pptx"

# 🎨 تغيير السمة
python presentation_template.py --theme light_professional

# 🔍 وضع التصحيح (عرض تفاصيل الأخطاء)
python presentation_template.py --verbose

# 🧪 تجربة دون إنشاء ملف
python presentation_template.py --dry-run

# 🎯 كل الخيارات معًا
python presentation_template.py \
    --output "تخرج_2026.pptx" \
    --theme minimal \
    --verbose
```

---

## ❓ الأسئلة الشائعة

### ❓ لماذا لا تظهر الصور التي أضفتها؟

✅ تأكد من:
1. أن المسار في `ASSET_PATHS` صحيح نسبيًا للملف الرئيسي
2. أن اسم الملف يطابق تمامًا (حساس لحالة الأحرف)
3. أن الصيغة مدعومة (.png, .jpg, .jpeg)

```python
# ❌ خطأ شائع
"mark": "Logo.png"  # إذا كان الاسم الفعلي: logo.png

# ✅ الصحيح
"mark": "assets/logos/logo.png"
```

### ❓ كيف أغير اتجاه النص للإنجليزية؟

```python
# في قسم الإعدادات:
PROJECT_INFO = {
    "language": "en",  # غيّر من "ar" إلى "en"
    # ...
}

SLIDES_CONFIG = {
    "rtl_enabled": False,  # عطّل دعم RTL
    # ...
}
```

### ❓ الخلفية تبدو مختلفة عن المتوقع؟

✅ الحلول:
1. جرب سمة أخرى: `--theme light_professional`
2. قلّل كثافة التأثيرات: `"visual_effects_intensity": 0.3`
3. عدّل ألوان الخلفية في `DESIGN_SYSTEM["colors"]`

### ❓ كيف أضيف ملاحظات للمتحدث؟

```python
# في قسم الإعدادات:
EXPORT_OPTIONS = {
    "add_speaker_notes": True,  # تأكد أنها True
    # ...
}

# ثم عدّل قائمة SPEAKER_NOTES في الكود:
SPEAKER_NOTES = [
    "ملاحظة للشريحة 1: افتح العرض بابتسامة ورحّب باللجنة...",
    "ملاحظة للشريحة 2: ركّز على أن المشروع حل لمشكلة حقيقية...",
    # أضف ملاحظة لكل شريحة بالترتيب
]
```

### ❓ هل يعمل على Mac/Linux؟

✅ **نعم!** مع ملاحظة:
- الميزات الأساسية تعمل على جميع الأنظمة
- ميزة `add_animations` و `com_enhancement` تتطلب Windows + pywin32
- إذا أردت حركات انتقالية على Mac، استخدم PowerPoint مباشرة بعد إنشاء الملف

### ❓ كيف أشارك القالب مع زملائي؟

1. ارفع الملف `presentation_template.py` على GitHub
2. فعّل خيار **Template repository** من إعدادات المستودع
3. شارك الرابط مع زملائك
4. كل شخص ينشئ نسخة خاصة به ويعدّل الإعدادات فقط

---

## 🤝 المساهمة في تطوير القالب

نرحب بمساهماتك! 🎉

### 🐛 الإبلاغ عن مشكلة
1. ابحث في [المشاكل المفتوحة](../../issues) أولاً
2. إذا لم تجد مشكلتك، أنشئ [مشكلة جديدة](../../issues/new) مع:
   - وصف واضح للمشكلة
   - خطوات إعادة الإنتاج
   - نظام التشغيل وإصدار بايثون
   - لقطات شاشة إذا أمكن

### 💡 اقتراح ميزة جديدة
1. افتح [نقاشًا جديدًا](../../discussions/new) لوصف فكرتك
2. اشرح الفائدة المتوقعة للمستخدمين
3. نناقش الفكرة ونخطط للتنفيذ معًا

### 🔧 إرسال تحسينات كودية
```bash
# 1. انشئ Fork من المستودع
# 2. انشئ فرعًا جديدًا
git checkout -b feature/my-awesome-feature

# 3. عدّل الكود مع الحفاظ على نمط التنسيق الحالي
# 4. اختبر التغييرات محليًا
python presentation_template.py --verbose

# 5. أرسل Pull Request مع وصف واضح للتغييرات
```

### 📝 معايير الكود
- ✅ تعليقات عربية/إنجليزية تشرح الأجزاء المعقدة
- ✅ دوال صغيرة ومركزة على مهمة واحدة
- ✅ أسماء متغيرات ودوال واضحة ومعبرة
- ✅ عدم كسر التوافق مع الإعدادات الحالية

---

## 📄 الترخيص

هذا المشروع مرخص تحت **رخصة MIT** - يمكنك:

✅ استخدامه في مشاريع شخصية أو تجارية  
✅ تعديله وتخصيصه كما تشاء  
✅ توزيعه ومشاركته مع الآخرين  
✅ استخدامه كأساس لمشاري

> : هذا قالب مفتوح المصدر — مساهمتك تجعله أفضل للجميع! 🚀

---

<p align="center">

</p>
