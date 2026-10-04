# Prompt — Google AI Studio (Build)

> انسخ كل شي تحت الخط وحطّو في Google AI Studio ← Build.

---

## Role

You are a senior full-stack engineer and an expert in the Tunisian primary-school curriculum (البرامج الرسمية التونسية). Build a complete, production-quality web app called **«مولّد المذكرات»** (Lesson-Memo Generator) for Tunisian primary-school teachers.

## What the app does (one sentence)

The teacher uploads the official **teacher's guide** (دليل المعلم, PDF) and **student book** (كتاب التلميذ, PDF) for a subject, plus their own **Word memo template** (.docx). They pick a lesson and its page numbers. The app reads those pages with Gemini and splits the content into the official lesson stages. The teacher reviews and edits every field, then downloads a filled Word memo (مذكرة / جذاذة).

## Tech stack (mandatory)

- React + TypeScript + Tailwind CSS, single-page app, everything runs in the browser.
- `@google/genai` for Gemini. Use the latest **Gemini Flash** model available in AI Studio, with JSON output (`responseMimeType: "application/json"` + `responseSchema`).
- `pdfjs-dist`: read the page count and show page thumbnails so the teacher can check they chose the right pages.
- `pdf-lib`: copy only the selected pages into a small new PDF, which is sent to Gemini (never send the whole book).
- `pizzip` + `docxtemplater`: fill `{{...}}` placeholders. Use `paragraphLoop: true`, `linebreaks: true`, delimiters `{{ }}`.
- `jszip`: direct XML editing for templates without placeholders (see "Label mode").
- `file-saver`: download the result.
- IndexedDB (via `idb-keyval`): store the library (PDFs, templates, settings) so the teacher uploads once.

## UI and language

- **All UI in Arabic, `dir="rtl"`**, font "Noto Naskh Arabic" or "Cairo" (Google Fonts).
- Mobile-first: most teachers use a phone. Big touch targets. A fixed bottom bar on the form screen, and a fixed top bar (رجوع / حفظ) on the review screen.
- Messages use simple, friendly Arabic that a Tunisian teacher understands. No technical jargon in errors.
- Light, calm, professional design: one primary colour (teal), cards, clear section titles.

## Screens

### 1. المكتبة (Library)

- The teacher creates **levels → subjects**, e.g. «السنة السادسة» → «الإيقاظ العلمي».
- Each subject has three upload slots:
  - `دليل المعلم` (PDF)
  - `كتاب التلميذ` (PDF, optional)
  - `نموذج المذكرة` (.docx, optional; if missing, use the built-in default template)
- Show the file name, size and page count. Allow replacing or deleting a file, with a confirmation.

### 2. توليد مذكرة (Generator form)

Fields:
- المستوى / المادة (dropdowns from the library)
- عنوان الدرس (text, required)
- رقم الأسبوع (number 1–36)
- التاريخ (date, default today, format `dd/mm/yyyy`)
- صفحات الدليل: من … إلى …
- صفحات الكتاب: من … إلى … (hidden if there is no book)

Behaviour:
- Validate the page ranges against the real page count.
- Show small thumbnails of the first and last selected page so the teacher can check them.
- Warn if the range is a single page 1–1 (usually the cover): «اختر صفحات الدرس، موش صفحة الغلاف».
- A big «تحليل ومراجعة» button shows a progress state: «نقرا في الصفحات…».

### 3. المراجعة (Review)

- One editable multi-line text area per section, in the stage order below. Each label shows both official names, e.g. «الوضعية المشكل — ألاحظ وأتساءل».
- Fields that came back empty get a **yellow background**, plus a note at the top listing them: «أقسام ما تلقاتش لا في الدليل لا في الكتاب: …».
- A collapsible «النص الأصلي» panel shows the raw text Gemini read, so the teacher can copy from it.
- Buttons: «رجوع» (keeps the form values) and «حفظ وتحميل».

### 4. Save

- File name: `المادة - الدرس - أسبوع X.docx`, with characters invalid in file names removed.
- After download, show a summary:
  - which template cells were filled;
  - which stayed empty because no text was found: «خانات بقات فارغة خاطر ما لقيتلهمش نص: …».

### 5. الإعدادات (Settings)

- An editor for the section names and their aliases (the JSON described below), with a «القيم الافتراضية» reset button.
- Export / import the whole library settings as JSON.

## The official lesson stages (domain knowledge — use exactly these)

The same stage has one name in the teacher's guide and another in the student book. The **section key** is what appears in the template as `{{key}}`.

| Section key | Name in دليل المعلم | Name in كتاب التلميذ |
|---|---|---|
| المكتسبات السابقة | المكتسبات السابقة | أتعهد مكتسباتي السابقة |
| الوضعية المشكل | الوضعية المشكل / الوضعية المشكلة / الوضعية الانطلاقية | ألاحظ وأتساءل |
| تحليل الوضعية ورصد التصورات | تحليل الوضعية ورصد التصورات / رصد التصورات | أفترض |
| التحقق العلمي | التحقق العلمي + النشاط الأول، النشاط الثاني… | أجرب وأتثبت (التجارب) |
| الاستنتاج | الاستنتاج (1)، (2)… | أستنتج |
| التطبيق والتوظيف | التطبيق والتوظيف / التطبيق | أطبق وأوظف |
| التقييم | التقييم / التقويم | أقيم تعلمي الجديد |
| التوسع والامتداد | التوسع والامتداد | — |
| معجمي في العلوم | — | معجمي في العلوم |
| أتهيأ لتعلمي اللاحق | — | أتهيأ لتعلمي اللاحق |

Lesson card fields (mostly from the guide):

- الكفاية النهائية
- المكون الأول
- المكون الثاني
- الوحدة
- المفاهيم
- المحتوى
- الهدف (الهدف المميز)
- المستلزمات (المستلزمات البيداغوجية)
- الحواجز
- مؤشرات التجاوز
- مؤشرات القدرة المستهدفة

Standard fields filled from the form: المادة، الدرس، الأسبوع، التاريخ.

## Gemini extraction (core logic)

1. Use `pdf-lib` to build a PDF with **only** the selected guide pages, and another with the selected book pages. Send them as `inlineData` (`application/pdf`).
   - Gemini reads the **rendered pages**. This matters: many Tunisian ministry PDFs use legacy "Arabic XT / AXt" fonts whose text layer is garbage Latin characters (e.g. `á«Hô©dG` instead of `العربية`). Never rely on the PDF text layer.
2. System instruction (put it in the code, in English):
   - You are an assistant to a Tunisian primary teacher.
   - Copy the text **verbatim** from the pages: same words, same order, same numbering (النشاط الأول، الاستنتاج 1…).
   - **Never invent, summarise, rephrase or add pedagogy.** If a section is not on the pages, return an empty string.
   - Classify each passage under the section whose heading it follows in the document. Use the guide/book name table above, including all aliases.
   - Keep line breaks between list items. Remove page numbers, running headers and footers.
   - Read Arabic right-to-left correctly and keep Arabic punctuation.
3. `responseSchema`:

   ```json
   {
     "sections": {
       "<section key>": { "guide": "string", "book": "string" }
     },
     "raw_guide": "string",
     "raw_book": "string"
   }
   ```

   One property per section key, all required, empty string allowed.
4. Merge rule for the review field of each section:
   - both parts present → `guide + "\n\nكتاب التلميذ:\n" + book`;
   - otherwise → whichever part is non-empty.
   - Also keep the separate values for the template keys `{{القسم - الدليل}}` and `{{القسم - الكتاب}}`.
5. Errors, each shown in Arabic:
   - no API key or quota exceeded;
   - pages unreadable;
   - empty result: «ما لقيتش نص في الصفحات هذي. ثبّت أرقام الصفحات».
   - Retry once on a network error.

## Filling the Word template

Before writing, **strip XML-invalid characters** from every value: `/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F￾￿]/g`. PDFs often contain them, and Word refuses the file.

### Mode A — placeholders (`{{...}}` present)

- Supported keys:
  - every section key;
  - `{{القسم - الدليل}}` and `{{القسم - الكتاب}}` for each section;
  - `{{المادة}} {{الدرس}} {{الأسبوع}} {{التاريخ}}`;
  - `{{الكتاب}}` (all the book text).
- Placeholders may be split across several runs by Word. docxtemplater handles this; keep the run formatting of the placeholder.
- Fill the body, tables (including nested ones), text boxes, headers and footers.
- After filling, warn about placeholders in the template with no matching key. List them; do not fail.

### Mode B — label mode (template has **no** `{{...}}`)

Most teachers have their own table template, e.g.:

| المراحل | نشاط المدرس | نشاط المتعلم | الوسائل |
|---|---|---|---|
| تعهد المكتسبات | | ينجز المطلوب | |
| الوضعية الاستكشافية | | عمل جماعي | فيديو |
| رصد التصورات | (existing text) | | السبورة |
| … | | | |

Rules (edit `word/document.xml` with JSZip and DOMParser/XMLSerializer):

1. **Find label cells.** Compare text ignoring spaces, tashkeel and the letter variants أ/إ/آ→ا, ة→ه, ى→ي. A cell matches a label if its text equals an alias, or starts with an alias followed by a non-letter or by a long explanation (e.g. «التحقق العلمي المنهج التجريبي / البحث الوثائقي»). Longest alias first.
2. **Header row.** If a row contains «نشاط المدرس» (aliases: نشاط المعلم، نشاط المربي، سير الأنشطة، الأنشطة), write stage text into that column, using its grid position and counting `gridSpan`. If the row also contains «المراحل» (aliases: المرحلة، مراحل الدرس، سير الدرس), only look for labels in that column.
3. **Both directions.** Tables may be stored RTL (`w:bidiVisual`, the label is the first `w:tc`) or LTR (the label is the last `w:tc`). It must work in both.
4. **No header.** For card tables («الهدف المميز | …»), write into the neighbouring cell: the next one, or the previous one if the label is the last cell.
5. **Continuation tables.** A table continued on the next page with the same column count and no header row inherits the previous table's column layout.
6. **Writing a cell.**
   - If the cell is empty, remove its extra empty paragraphs and write from the first one.
   - If it already has text, **keep it** and append below.
   - Write one `w:p` per line, with `w:bidi` on the paragraph and `w:rtl` on the run.
   - Copy the paragraph-mark run properties (`w:pPr/w:rPr`) so the teacher's font and size are kept.
7. **Same section, two labels** (e.g. «رصد التصورات» and «بناء الفرضيات»): the first gets the guide part, the second the book part. If there is no book part, both get the full text.
8. Never write into a cell that itself matches a label. Never touch the «نشاط المتعلم» or «الوسائل» columns.

Default label aliases (editable in Settings):

```json
{
  "المادة": ["المادة"],
  "الدرس": ["الدرس", "عنوان الدرس", "موضوع الدرس", "الموضوع"],
  "الأسبوع": ["الأسبوع"],
  "التاريخ": ["التاريخ"],
  "الوحدة": ["الوحدة"],
  "الكفاية النهائية": ["الكفاية النهائية", "نص الكفاية النهائية", "الكفاية"],
  "المكون الأول+المكون الثاني": ["مكون الكفاية", "مكونات الكفاية", "المكون"],
  "الهدف": ["الهدف المميز", "الأهداف المميزة", "الهدف", "الأهداف"],
  "المحتوى": ["المحتوى", "المحتويات"],
  "المفاهيم": ["المفاهيم"],
  "المستلزمات": ["المستلزمات", "المستلزمات البيداغوجية"],
  "الحواجز": ["الحواجز"],
  "مؤشرات التجاوز": ["مؤشرات التجاوز"],
  "مؤشرات القدرة المستهدفة": ["مؤشرات القدرة المستهدفة", "مؤشرات القدرة"],
  "المكتسبات السابقة": ["تعهد المكتسبات", "المكتسبات السابقة", "أتعهد مكتسباتي", "المكتسبات"],
  "الوضعية المشكل": ["الوضعية المشكل", "الوضعية المشكلة", "الوضعية الاستكشافية", "الوضعية الانطلاقية", "ألاحظ وأتساءل", "الانطلاق"],
  "تحليل الوضعية ورصد التصورات": ["تحليل الوضعية ورصد التصورات", "تحليل الوضعية", "رصد التصورات", "التصورات", "بناء الفرضيات", "الفرضيات", "أفترض"],
  "التحقق العلمي": ["التحقق العلمي", "أجرب وأتثبت", "التجريب"],
  "الاستنتاج": ["الاستنتاج", "أستنتج"],
  "التطبيق والتوظيف": ["التطبيق والتوظيف", "التطبيق", "أطبق وأوظف", "التوظيف", "التعلم الآلي"],
  "التوسع والامتداد": ["التوسع والامتداد", "التعلم الإدماجي", "الإدماج"],
  "التقييم": ["التقييم", "التقويم", "أقيم تعلمي الجديد", "أقيم تعلمي"]
}
```

(`"A+B"` means: join both sections' text with a line break.)

### Built-in default template

If the subject has no template, generate one with the `docx` library:

- A4, RTL, Arabic font.
- A title.
- A lesson card table (المادة، الدرس، الأسبوع، التاريخ، الوحدة، الكفاية النهائية، المكون، الهدف، المحتوى، المستلزمات).
- A «سير الحصّة» table with columns «المرحلة | نشاط المدرس | نشاط المتعلم | الوسائل». It has one row per stage, showing the guide name and the book name, with the matching `{{…}}` in «نشاط المدرس».

## Privacy & robustness

- Files never leave the browser except the **selected pages** sent to Gemini. Say this in a small note on the generator screen.
- Handle large PDFs (100+ MB) without freezing the UI: do the PDF work in a Web Worker or `await` in chunks, and show a spinner.
- If the review or form screen is left open, do not lose the edited values.
- If a template is corrupt, show «ما نجمتش نحل النموذج، ربما الملف تالف».

## Acceptance checklist (test these yourself before finishing)

1. Upload a guide PDF and pick its lesson pages. Every stage field is filled with verbatim text, and the numbered activities are kept.
2. An AXt-encoded ministry PDF gives correct Arabic, not Latin garbage.
3. A template with `{{الهدف}}` split across runs is filled, and the formatting is kept.
4. A template with **no** placeholders and an LTR stages table (labels in the last column) gets every stage written into «نشاط المدرس». Existing text in a cell is kept.
5. Text containing `\x01` or `\x0C` saves without error.
6. The downloaded file name is `الإيقاظ العلمي - الهواء - أسبوع 3.docx`.
7. Everything works on a 360 px wide phone screen in RTL.

Deliver clean, commented, modular code: `pdf/`, `gemini/`, `docx/fillPlaceholders.ts`, `docx/fillByLabels.ts`, `ui/`, and `config/sections.ts` holding the tables above as typed constants.
