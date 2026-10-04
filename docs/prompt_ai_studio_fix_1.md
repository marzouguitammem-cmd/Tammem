# Fix prompt — Google AI Studio (round 1)

> انسخ كل شي تحت الخط وابعثو في نفس المحادثة متاع التطبيق في AI Studio.

---

The app works, but it has **critical bugs**. Fix all of them without changing the design, the navigation or the features that already work.

## Bug 1 — Hard-coded demo content (CRITICAL)

When the app opens, fields are already filled with a sample lesson about air:

- الكفاية النهائية: «حل وضعيات مشكل دالة بإنجاز مشاريع متصلة بالمادة…»
- مكون الكفاية: «يفسر بعض الظواهر الفيزيائية المتصلة بخصائص الهواء والغازات»
- الهدف المميز: «يتعرف المتعلم مكونات الهواء…»
- مؤشرات القدرة: «يقيس ارتفاع الماء في المخبار…»
- عنوان الدرس «الهواء ومكوناته» and الوحدة «المادة والطاقة (الوحدة 3)».

Whatever lesson or pages I choose, the generated memo is **always the same air lesson** (candle under a glass, 1/5 oxygen…). That content is **not** in the pages I selected. Pages 2–3 of my guide are the book's introduction, and page 1 of the student book is the cover, yet it still produced a full air lesson. So the app is either using mock data or letting the model invent.

Required fixes:

1. **Delete every hard-coded sample, mock, demo or default lesson text** in the whole codebase: initial state, placeholders used as values, `useState` defaults, sample JSON, fallback results, seed data, "example" constants. Search the code for «الهواء», «الأكسجين», «شمعة», «المادة والطاقة» and remove all of them from data and state.
2. Every form field starts **empty**. Light grey `placeholder` hints are fine (e.g. «مثال: الهواء ومكوناته»), but a placeholder must never be submitted as a value.
3. Remove the «اقتراحات سريعة لكفايات البرامج التونسية» chips. If you keep them, they must only fill a field **when I tap one**, never automatically.
4. Make the lesson-card fields (الكفاية النهائية، المكون، الهدف المميز، مؤشرات القدرة) **optional**. They must not be required inputs on the form. These values must come **from the PDF pages**. If I type something, it is only an override shown in the review screen.
5. **Never fall back to sample data.** If the API call fails or returns nothing, show an error. Never show a fake result.

## Bug 2 — The memo is not built from the pages I select (CRITICAL)

The memo must be built **only** from the guide and book pages I type, every time, for every lesson.

1. Use `pdf-lib`'s `copyPages` to build a new PDF containing **exactly** pages `from…to` of the guide (1-based, inclusive), and do the same for the book. Send those bytes to Gemini as `inlineData` with `mimeType: "application/pdf"`. Do not send the whole file, thumbnails only, a cached earlier PDF, or text from a previous lesson.
2. **Do not send the lesson title, unit, competencies or any form text to Gemini as content.** Gemini receives only the selected pages and the extraction instructions. The title, unit, week and date are used only for the memo header and the file name. If the title is sent, the model invents a lesson from it.
3. Gemini settings: `temperature: 0`, `responseMimeType: "application/json"`, a strict `responseSchema`.
4. System instruction (replace the current one):
   > You extract text from the attached PDF pages of a Tunisian primary-school teacher's guide and student book. Copy text **verbatim** from these pages only: same words, same order, same numbering. **Do not invent, summarise, rephrase, complete or add anything**, and do not use your own knowledge of the curriculum. If a section does not appear on these pages, return an empty string for it. If the pages contain no lesson (cover, introduction, table of contents), return all sections empty and set `"no_lesson_found": true`. For every non-empty section, return the page number(s) where you found it.
5. Schema:
   - for each section: `{ "guide": string, "book": string, "pages": number[] }`;
   - plus `"raw_guide"`, `"raw_book"` (the full transcription of the pages);
   - plus `"no_lesson_found"`: boolean.
6. **Grounding check** after the response. For each section, normalise its text and `raw_guide` / `raw_book`:
   - normalising means removing tashkeel and spaces, and treating أإآ→ا, ة→ه, ى→ي;
   - check that at least 80 % of the section's words appear in the raw text of the pages;
   - if not, empty that section, highlight it in yellow, and add a note: «هالجزء ما لقيتوش في الصفحات المختارة».
7. If `no_lesson_found` is true, or every section is empty, show «ما لقيتش درس في الصفحات هذي. ثبّت أرقام الصفحات (موش الغلاف ولا المقدمة)». Stay on the form; do not open the review screen.
8. **No stale state.** Each «توليد» click must:
   - clear the previous result;
   - abort any request still in flight (`AbortController`);
   - start a fresh request.

   Do not cache results by subject. The history tab («السجل») stores past memos separately; opening a new generation must never load a history item.
9. Add a small collapsible «تفاصيل تقنية» panel on the review screen, with:
   - the page ranges sent;
   - the number of pages and size (KB) of each PDF sent;
   - the first 300 characters of `raw_guide` and `raw_book`.

   This lets me check the app really read my pages.

## Bug 3 — Page-number inputs cannot be cleared

In «تحديد أرقام الصفحات», the fields already contain a number (e.g. 2, 3, 1). When I delete it, the number comes back, so I cannot type my own value.

Required fixes:

1. Store each page field as a **string** in state (`""` allowed), not as a number. Do not use `Number(value) || 1`, `parseInt(...) || 1`, `Math.max(1, …)` or any clamp inside `onChange`.
2. Use `<input type="text" inputMode="numeric" pattern="[0-9]*">`. In `onChange`, keep only digits (`value.replace(/[^0-9٠-٩]/g, "")`) and convert Arabic-Indic digits to Latin.
3. All page fields start **empty**, for the guide and the book, with placeholder «من» / «إلى».
4. Validate **only when I press «توليد»**:
   - both fields are filled;
   - 1 ≤ from ≤ to ≤ total pages;
   - at most 15 pages per document.

   Show the error under the field in red, in Arabic.
5. Show the thumbnails only when the field holds a valid number. Update them after typing stops (debounce 400 ms), not on every keystroke.
6. The book range is optional. If it is empty, send only the guide.

## Acceptance tests (run them before you answer)

1. Open the app fresh. No field contains lesson text and no page field contains a number.
2. In a page field, type `12`, delete it completely: the field stays empty. Type `45`: it shows `45`.
3. Guide pages 2–3 (introduction) and book page 1 (cover) give the message «ما لقيتش درس…». No air lesson appears.
4. Two different lessons with two different page ranges give two **different** memos. Each memo contains only sentences that exist in its own pages (check with «تفاصيل تقنية»).
5. Search the codebase: no lesson sample text remains (`grep` for «شمعة», «الأكسجين», «الهواء ومكوناته»).

Keep the Arabic RTL UI, the review screen, Word export and PDF printing exactly as they are. Show me the list of files you changed.
