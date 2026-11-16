import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from build_full_assessment import answers

def create_document():
    doc = Document()

    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
        # Header / Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "JustGo | Data Solution Engineer | Written Assessment (V4)"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.name = "Calibri"
        hp.style.font.size = Pt(8.5)
        hp.style.font.color.rgb = RGBColor(128, 128, 128)

        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "JustGo, Dhaka — Data Solution Engineer Written Assessment"
        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp.style.font.name = "Calibri"
        fp.style.font.size = Pt(8.5)
        fp.style.font.color.rgb = RGBColor(128, 128, 128)

    # Styles
    COLOR_PRIMARY = RGBColor(27, 54, 93)     # #1B365D Dark Navy
    COLOR_SECONDARY = RGBColor(0, 102, 204)  # #0066CC Blue
    COLOR_TEXT = RGBColor(51, 51, 51)        # #333333 Charcoal
    COLOR_MUTED = RGBColor(102, 102, 102)

    def add_title(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(22)
        run.font.color.rgb = COLOR_PRIMARY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        return p

    def add_subtitle(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(13)
        run.font.color.rgb = COLOR_SECONDARY
        p.paragraph_format.space_after = Pt(14)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(15)
        run.font.color.rgb = COLOR_PRIMARY
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(12.5)
        run.font.color.rgb = COLOR_SECONDARY
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        return p

    def add_p(text, bold=False, italic=False):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_TEXT
        run.bold = bold
        run.italic = italic
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        return p

    def add_bullet(text, bold_prefix="", level=0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.12
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(10)
            r_pre.font.color.rgb = COLOR_TEXT
        r_txt = p.add_run(text)
        r_txt.font.name = "Calibri"
        r_txt.font.size = Pt(10)
        r_txt.font.color.rgb = COLOR_TEXT
        return p

    def add_answer_box(content, is_code=False):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        # Border
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="999999"/><w:left w:val="single" w:sz="6" w:space="0" w:color="999999"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="999999"/><w:right w:val="single" w:sz="6" w:space="0" w:color="999999"/></w:tcBorders>')
        tcPr.append(borders)
        
        # Fill background - very light grey for code, off-white for text
        fill_hex = "F8F9FA" if is_code else "FFFFFF"
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

        # Margins inside cell
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

        # Populate cell content
        cp = cell.paragraphs[0]
        lines = content.strip().split('\n')
        for i, line in enumerate(lines):
            if i > 0:
                cp = cell.add_paragraph()
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(2 if not is_code else 0)
            cp.paragraph_format.line_spacing = 1.05 if is_code else 1.15
            run = cp.add_run(line)
            if is_code:
                run.font.name = "Consolas"
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(34, 34, 34)
            else:
                run.font.name = "Calibri"
                run.font.size = Pt(10)
                run.font.color.rgb = COLOR_TEXT

        # Add spacing after table
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(4)
        sp.paragraph_format.space_after = Pt(8)

    # -----------------------------------------------------------------------
    # DOCUMENT FRONT MATTER
    # -----------------------------------------------------------------------
    add_title("Data Solution Engineer: Written Assessment")
    add_subtitle("JustGo, Dhaka")

    add_h1("Overview")
    add_p("This assignment is designed to assess your skills in:")
    add_bullet(" Complex queries, stored procedures, and data migration scenarios.", "T-SQL:")
    add_bullet(" Application development, API/data integration, and problem-solving.", "C#:")
    add_bullet(" Your ability to analyze scenarios, propose strategies, and explain your design decisions.", "Technical Consultancy:")
    add_bullet(" Your ability to understand what non-technical customers and colleagues actually need, reason from evidence, explain technical issues in plain language, and recommend the option that best fits the purpose.", "Customer Scenarios:")
    add_p("All of these skills are assessed through realistic work scenarios of the kind you would handle in this role. Several scenarios form one storyline, so read them in order.")

    add_h1("About this assessment")
    add_bullet("There are ten questions (Q1 to Q10), most with more than one part. You will write SQL and C# from scratch, review and correct existing code, scope work with customers and write to them in plain language.")
    add_bullet("There is no time limit for the work itself, but there is a fixed deadline: please submit by the deadline stated in your invitation email. Submissions after the deadline cannot be accepted. We suggest spreading the work over several sittings.")
    add_bullet("All organisations, people, data, schemas, code and API details are fictional and simplified for this exercise. They are not real JustGo data or system internals. Where something is not stated, make a sensible assumption and say what it is.")
    add_bullet("A follow-up interview and live session will discuss your answers, including your code and design choices.")

    add_h1("How to answer")
    add_bullet("Type your answers in the matching boxes in the Answer Section at the end of this document.")
    add_bullet("Do not delete, rename or move any heading in the Answer Section. Your answers are matched to questions by their headings. Replace the text 'Type your answer here' in each box with your answer. Boxes grow as you type.")
    add_bullet("Every answer must cite the evidence it relies on (for example 'Exhibit Q4-B, row M-203', 'Exhibit Q5, line 24' or 'Exhibit Q7, batch B3') and explain your reasoning. Statements without evidence or reasoning score lower, however polished they are.")
    add_bullet("Write customer and colleague replies exactly as you would send them. Plain, friendly, professional English is what we are looking for.")
    add_bullet("Word limits are shown next to each answer heading. They are a guide to the length we expect; stay close to them.")
    add_bullet("Code answers (SQL, C#) go in the boxes formatted in Consolas. Paste code as plain text only (in Word: right-click, Paste Options, Keep Text Only). Keep required comments inside the code. Do not paste screenshots or images of code. You are welcome to run your code; where we ask for output, paste what your program actually printed.")
    add_bullet("Keep everything inside this one document. Do not attach extra files, links or images.")

    add_h1("How to submit")
    add_bullet("Save this document as a Word file (.docx). Do not convert it to PDF or any other format.")
    add_bullet("Name the file: Full Name - email address - DSE Assessment.docx, for example: Rahim Uddin - rahim.uddin@example.com - DSE Assessment.docx. Use the same email address you applied with.")
    add_bullet("Upload it using the submission form link in your invitation. The form also asks for your name and email address. Only one file is accepted, and it must be a .docx file. Do not add your name, contact details or any declaration inside the document itself; the form collects these.")
    add_p("We look forward to reviewing your assignment and getting insights into your unique approach and problem-solving methodology. Good luck, and be prepared to showcase both your technical and personal expertise!", italic=True)

    doc.add_page_break()

    # -----------------------------------------------------------------------
    # QUESTIONS & EXHIBITS OVERVIEW
    # -----------------------------------------------------------------------
    add_title("Questions & Exhibits Reference")

    add_h1("Storyline 1 - Events (Q1 and Q2)")
    add_h2("Q1. 'Which events are nearly full?'")
    add_p("Midland Gymnastics Association tracks events, registrations, and payments. Claire Dawson (Events Manager) wants to identify events exceeding 90% capacity to open waiting lists or schedule extra sessions. Paid registrations only count; some events lack capacity settings; delivery requested every Monday morning.")

    add_h2("Q2. 'Looking up one event on demand'")
    add_p("Support team lead requires a reusable stored procedure GetEventRegistrationDetails accepting EventID to return event details, confirmed paid registrants, and error handling for missing events.")

    add_h1("Storyline 2 - Reporting (Q3)")
    add_h2("Q3. 'The family membership report'")
    add_p("Helen Carter (Membership Secretary, Northshire Swimming Association) requests a family membership report. Treasurer Graham notes renewal totals don't match bank receipts (£1,440 vs £540) due to one-line-per-person output. Helen also plans junior squad communications.")

    add_h1("Storyline 3 - A migration (Q4 to Q7)")
    add_h2("Q4. 'This is NOT what we were promised'")
    add_p("Mark Ellis (Chair, Riverside Rowing Club) reports active members incorrectly flagged as Expired ahead of Friday regatta entries. Exhibit Q4-B reveals US/UK date swap (MM/dd/yyyy vs dd/MM/yyyy) when Day <= 12.")

    add_h2("Q5. 'Working with existing code: the import utility'")
    add_p("Review C# import utility in Exhibit Q5. Identify culture flaws, swallowed exceptions (2,412 source rows vs 2,389 imported = 23 missing members), and provide corrected ParseDate and ImportAll implementations.")

    add_h2("Q6. 'Writing new code: a pre-import validation tool'")
    add_p("Develop standalone C# console application to validate regatta entry CSV data against schema rules, positive IDs, and allowed PaymentStatus values, outputting an execution summary.")

    add_h2("Q7. 'Planning the rest of the migration'")
    add_p("Three Rivers Rowing Association migration for remaining 37 clubs (2.55M rows, 58h window). Analyze trial log (FK constraint error on MemberID 40917, 2,000 duplicate registrations upon retry, transaction log exhaustion Msg 9002), and present a 1-page optimization architecture.")

    add_h1("Storyline 4 - Day-to-day requests (Q8 to Q10)")
    add_h2("Q8. 'Can you run a script tonight?'")
    add_p("Rafi (Onboarding) requests an urgent Thursday evening script to lapse and remove unrenewed Eastvale Badminton members before season starts on 1 Sept. Analyze ambiguities, assess members E-01 to E-07, and propose safe governance.")

    add_h2("Q9. 'Explaining a technical limit'")
    add_p("Coastline Triathlon requests full live sync every 10 minutes for 18,000 members. Mathematical evaluation against 120 req/min rate limit (18,180 calls take 151 mins). Recommend webhooks / modifiedSince incremental sync.")

    add_h2("Q10. 'Your dashboard is wrong'")
    add_p("Daniel Ross (Westfield Netball) questions discrepancy between 1,240 active dashboard metric and 1,312 exported report rows for national funding return. Reconcile dual memberships (58) and month-to-date turnover (+17 new, -31 expired).")

    doc.add_page_break()

    # -----------------------------------------------------------------------
    # ANSWER SECTION
    # -----------------------------------------------------------------------
    add_title("Answer Section")
    add_p("Type each answer in the box under its heading. Do not delete, rename or move the headings. Cite the exhibit row, line or figure you rely on and explain your reasoning.", italic=True)

    # 1. Answer Q1 Part A
    add_h1("Answer Q1 Part A - T-SQL query")
    add_p("Code box (Consolas). Paste code as plain text only. Include indexing recommendations and personalization comments as SQL comments.", italic=True)
    add_answer_box(answers["Q1_A"]["text"], is_code=True)

    # 2. Answer Q1 Part B
    add_h1("Answer Q1 Part B - Reply to Claire")
    add_p("Up to 300 words.", italic=True)
    add_answer_box(answers["Q1_B"]["text"], is_code=False)

    # 3. Answer Q2 Part A
    add_h1("Answer Q2 Part A - Stored procedure GetEventRegistrationDetails")
    add_p("Code box (Consolas). Paste code as plain text only. Include your comments in the code.", italic=True)
    add_answer_box(answers["Q2_A"]["text"], is_code=True)

    # 4. Answer Q2 Part B
    add_h1("Answer Q2 Part B - Note to the Support team")
    add_p("Up to 200 words.", italic=True)
    add_answer_box(answers["Q2_B"]["text"], is_code=False)

    # 5. Answer Q3 Part A1
    add_h1("Answer Q3 Part A1 - Reply to Helen")
    add_p("Up to 400 words.", italic=True)
    add_answer_box(answers["Q3_A1"]["text"], is_code=False)

    # 6. Answer Q3 Part A2
    add_h1("Answer Q3 Part A2 - Internal scoping note")
    add_p("Up to 300 words.", italic=True)
    add_answer_box(answers["Q3_A2"]["text"], is_code=False)

    # 7. Answer Q3 Part B
    add_h1("Answer Q3 Part B - SQL for your recommended design")
    add_p("Code box (Consolas). Paste code as plain text only.", italic=True)
    add_answer_box(answers["Q3_B"]["text"], is_code=True)

    # 8. Answer Q4 Part A
    add_h1("Answer Q4 Part A - Reply to Mark")
    add_p("Up to 250 words.", italic=True)
    add_answer_box(answers["Q4_A"]["text"], is_code=False)

    # 9. Answer Q4 Part B
    add_h1("Answer Q4 Part B - Internal note")
    add_p("Up to 450 words, including a short SQL sketch.", italic=True)
    add_answer_box(answers["Q4_B"]["text"], is_code=False)

    # 10. Answer Q5 Part A
    add_h1("Answer Q5 Part A - Problems found in the existing code")
    add_p("Up to 450 words. Refer to line numbers.", italic=True)
    add_answer_box(answers["Q5_A"]["text"], is_code=False)

    # 11. Answer Q5 Part B
    add_h1("Answer Q5 Part B - Corrected ParseDate and ImportAll")
    add_p("Code box (Consolas). Paste code as plain text only.", italic=True)
    add_answer_box(answers["Q5_B"]["text"], is_code=True)

    # 12. Answer Q6 Part A
    add_h1("Answer Q6 Part A - C# console application")
    add_p("Code box (Consolas). Paste code as plain text only. Include your comments, including the extension comment, in the code.", italic=True)
    add_answer_box(answers["Q6_A"]["text"], is_code=True)

    # 13. Answer Q6 Part B
    add_h1("Answer Q6 Part B - Explanation of approach and assumptions")
    add_p("1 to 2 paragraphs.", italic=True)
    add_answer_box(answers["Q6_B"]["text"], is_code=False)

    # 14. Answer Q6 Part C
    add_h1("Answer Q6 Part C - Program output for the sample CSV")
    add_p("Code box (Consolas). Paste code as plain text only. Paste exactly what your program printed.", italic=True)
    add_answer_box(answers["Q6_C"]["text"], is_code=True)

    # 15. Answer Q7 Part A
    add_h1("Answer Q7 Part A - One-page migration document")
    add_p("One page: about 500 to 600 words.", italic=True)
    add_answer_box(answers["Q7_A"]["text"], is_code=False)

    # 16. Answer Q7 Part B
    add_h1("Answer Q7 Part B - Reflection")
    add_p("3 to 4 sentences.", italic=True)
    add_answer_box(answers["Q7_B"]["text"], is_code=False)

    # 17. Answer Q8 Part A
    add_h1("Answer Q8 Part A - Ambiguities and assumptions")
    add_p("Up to 250 words.", italic=True)
    add_answer_box(answers["Q8_A"]["text"], is_code=False)

    # 18. Answer Q8 Part B
    add_h1("Answer Q8 Part B - Member by member (E-01 to E-07)")
    add_p("Up to 350 words. One short paragraph or line per member.", italic=True)
    add_answer_box(answers["Q8_B"]["text"], is_code=False)

    # 19. Answer Q8 Part C
    add_h1("Answer Q8 Part C - Your questions")
    add_p("Up to five questions, most important first.", italic=True)
    add_answer_box(answers["Q8_C"]["text"], is_code=False)

    # 20. Answer Q8 Part D
    add_h1("Answer Q8 Part D - Reply to Rafi")
    add_p("Up to 200 words.", italic=True)
    add_answer_box(answers["Q8_D"]["text"], is_code=False)

    # 21. Answer Q9 Part A
    add_h1("Answer Q9 Part A - Email to Jo")
    add_p("Up to 250 words.", italic=True)
    add_answer_box(answers["Q9_A"]["text"], is_code=False)

    # 22. Answer Q9 Part B
    add_h1("Answer Q9 Part B - Working and recommendation")
    add_p("Up to 400 words. Show your calculation.", italic=True)
    add_answer_box(answers["Q9_B"]["text"], is_code=False)

    # 23. Answer Q10 Part A
    add_h1("Answer Q10 Part A - Reconciliation")
    add_p("Up to 250 words. Show each step with the numbers.", italic=True)
    add_answer_box(answers["Q10_A"]["text"], is_code=False)

    # 24. Answer Q10 Part B
    add_h1("Answer Q10 Part B - Reply to Daniel")
    add_p("Up to 200 words.", italic=True)
    add_answer_box(answers["Q10_B"]["text"], is_code=False)

    # Save document
    filename = "Karib Shams - shams321karib@gmail.com - DSE Assessment.docx"
    filepath = os.path.join(r"d:\CSE599_THESIS\Data_Solution", filename)
    doc.save(filepath)
    print(f"Document successfully created and saved to: {filepath}")

if __name__ == "__main__":
    create_document()
