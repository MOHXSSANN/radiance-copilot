"""Generate PDF from Radiance Copilot G Drive Deployment Guide."""
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable
import os

OUTPUT = r"c:\Users\Moham\OneDrive\Desktop\Dmitry Codes Extraction\Radiance_Copilot_G_Drive_Deployment_Guide.pdf"

# ── Colours ──────────────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#1B3A6B")
BLUE   = colors.HexColor("#2563EB")
LGRAY  = colors.HexColor("#F3F4F6")
MGRAY  = colors.HexColor("#D1D5DB")
DGRAY  = colors.HexColor("#374151")
WHITE  = colors.white
RED    = colors.HexColor("#DC2626")

# ── Styles ────────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

sTitle = S("sTitle",
    fontName="Helvetica-Bold", fontSize=22, textColor=WHITE,
    spaceAfter=4, alignment=TA_CENTER)
sSubtitle = S("sSubtitle",
    fontName="Helvetica", fontSize=11, textColor=colors.HexColor("#BFDBFE"),
    spaceAfter=2, alignment=TA_CENTER)
sMeta = S("sMeta",
    fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#93C5FD"),
    spaceAfter=0, alignment=TA_CENTER)

sH1 = S("sH1",
    fontName="Helvetica-Bold", fontSize=14, textColor=NAVY,
    spaceBefore=18, spaceAfter=6, borderPadding=(0,0,4,0))
sH2 = S("sH2",
    fontName="Helvetica-Bold", fontSize=11, textColor=BLUE,
    spaceBefore=12, spaceAfter=4)
sH3 = S("sH3",
    fontName="Helvetica-BoldOblique", fontSize=10, textColor=DGRAY,
    spaceBefore=8, spaceAfter=3)

sBody = S("sBody",
    fontName="Helvetica", fontSize=9.5, textColor=DGRAY,
    spaceBefore=2, spaceAfter=4, leading=14)
sBold = S("sBold",
    fontName="Helvetica-Bold", fontSize=9.5, textColor=DGRAY,
    spaceBefore=2, spaceAfter=4, leading=14)
sBullet = S("sBullet",
    fontName="Helvetica", fontSize=9.5, textColor=DGRAY,
    leftIndent=16, spaceBefore=1, spaceAfter=2, leading=13,
    bulletIndent=6)
sCode = S("sCode",
    fontName="Courier", fontSize=8, textColor=colors.HexColor("#1F2937"),
    backColor=LGRAY, leftIndent=12, rightIndent=12,
    spaceBefore=4, spaceAfter=4, leading=12,
    borderPadding=6)
sFooter = S("sFooter",
    fontName="Helvetica-Oblique", fontSize=7.5, textColor=colors.HexColor("#9CA3AF"),
    alignment=TA_CENTER)
sTOC = S("sTOC",
    fontName="Helvetica", fontSize=9.5, textColor=BLUE,
    spaceBefore=3, spaceAfter=3, leading=14)
sKeyPrinciple = S("sKeyPrinciple",
    fontName="Helvetica-Bold", fontSize=10, textColor=NAVY,
    alignment=TA_CENTER)

# ── Helper: coloured rule ─────────────────────────────────────────────────────
def rule(color=NAVY, thickness=1.5, spaceB=4, spaceA=8):
    return HRFlowable(width="100%", thickness=thickness, color=color,
                      spaceAfter=spaceA, spaceBefore=spaceB)

def spacer(h=6):
    return Spacer(1, h)

def h1(text):
    return [rule(NAVY, 1.5, spaceB=14, spaceA=0),
            Paragraph(text, sH1),
            rule(MGRAY, 0.5, spaceB=2, spaceA=6)]

def h2(text):
    return [Paragraph(text, sH2)]

def h3(text):
    return [Paragraph(text, sH3)]

def body(text):
    return Paragraph(text, sBody)

def bold(text):
    return Paragraph(text, sBold)

def bullet(text):
    return Paragraph(f"• {text}", sBullet)

def code_block(text):
    from reportlab.platypus import Preformatted
    return Preformatted(text, sCode)

def table(data, col_widths, header=True):
    t = Table(data, colWidths=col_widths)
    style = [
        ("FONTNAME",    (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 8.5),
        ("BACKGROUND",  (0,0), (-1,0),  NAVY),
        ("TEXTCOLOR",   (0,0), (-1,0),  WHITE),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LGRAY]),
        ("GRID",        (0,0), (-1,-1), 0.4, MGRAY),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING",(0,0), (-1,-1), 7),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
        ("FONTNAME",    (0,1), (-1,-1), "Helvetica"),
        ("TEXTCOLOR",   (0,1), (-1,-1), DGRAY),
    ]
    t.setStyle(TableStyle(style))
    return t

# ── Cover block (drawn as a coloured table) ───────────────────────────────────
def cover_block():
    cover_data = [[
        Paragraph("RADIANCE COPILOT", sTitle),
    ],[
        Paragraph("G: Drive - How It Works", sSubtitle),
    ],[
        Paragraph("Canada Border Services Agency (CBSA)", sMeta),
    ],[
        Paragraph("March 2026", sMeta),
    ]]
    t = Table(cover_data, colWidths=[6.5*inch])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), NAVY),
        ("TOPPADDING",   (0,0), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0), (-1,-1), 10),
        ("LEFTPADDING",  (0,0), (-1,-1), 20),
        ("RIGHTPADDING", (0,0), (-1,-1), 20),
        ("ROUNDEDCORNERS", [8]),
    ]))
    return t

# ── Key principle callout ─────────────────────────────────────────────────────
def callout(text):
    data = [[Paragraph(text, sKeyPrinciple)]]
    t = Table(data, colWidths=[6.5*inch])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), colors.HexColor("#EFF6FF")),
        ("TOPPADDING",   (0,0), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0), (-1,-1), 10),
        ("LEFTPADDING",  (0,0), (-1,-1), 14),
        ("RIGHTPADDING", (0,0), (-1,-1), 14),
        ("BOX",          (0,0), (-1,-1), 1.5, BLUE),
    ]))
    return t

# ── Build story ───────────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=letter,
        leftMargin=inch, rightMargin=inch,
        topMargin=0.85*inch, bottomMargin=0.85*inch,
        title="Radiance Copilot - G: Drive Deployment Guide",
        author="CBSA Radiance Copilot",
    )

    W = 6.5 * inch  # usable width

    story = []

    # ── Cover ─────────────────────────────────────────────────────────────────
    story += [spacer(30), cover_block(), spacer(24)]

    story.append(callout("The app runs on each officer's PC. All case files and templates live on the shared G: drive."))
    story.append(spacer(14))

    # ── G Drive Structure ─────────────────────────────────────────────────────
    story += h1("G: Drive Folder Structure")
    story.append(code_block(
        "G:\\\n"
        "  Radiance\\\n"
        "  |\n"
        "  +-- templates\\\n"
        "  |     SAISIE A FAIRE_francompact 2025 - TEMPLATE.pdf\n"
        "  |     K138 Seizure Cannabis-Stupefiant - TEMPLATE.pdf\n"
        "  |     K138 Seizure Knives-Arms - TEMPLATE.pdf\n"
        "  |     K138 Stupefiant-Others - TEMPLATE.pdf\n"
        "  |     Agenda - TEMPLATE.pdf\n"
        "  |     Saisie d'interet - TEMPLATE.pdf\n"
        "  |     k138_note_cannabis.txt\n"
        "  |     k138_note_arms.txt\n"
        "  |     k138_note_other.txt\n"
        "  |\n"
        "  +-- cases\\\n"
        "        |\n"
        "        +-- 19747 2026-02-11 INV0000604 - Cannabis\\\n"
        "        |     20260211_SAISIE_scan.pdf\n"
        "        |     Agenda_19747.pdf\n"
        "        |     K138_20260211_SAISIE_scan.pdf\n"
        "        |     Saisie_Interet_19747.pdf\n"
        "        |     barcode_19747.png\n"
        "        |     .extracted_data\\  (hidden folder)\n"
        "        |       values_latest.json\n"
        "        |\n"
        "        +-- 19800 2026-02-15 INV0000712 - Knives\\\n"
        "        +-- 19850 2026-02-18 INV0000891 - Others\\"
    ))
    story += h2("Naming Convention for Case Folders")
    story.append(body("Case folders should follow a consistent naming pattern:"))
    story.append(code_block("  <Badge#>  <Date YYYY-MM-DD>  <InventoryNumber>  -  <SeizureType>\n\n"
                            "  Example:  19747 2026-02-11 INV0000604 - Cannabis"))
    story.append(body(
        "This format makes case folders sort chronologically and lets any officer instantly "
        "identify the case type at a glance."
    ))

    # ── Templates ─────────────────────────────────────────────────────────────
    story += h1("Templates on the G: Drive")
    story.append(body(
        "All template PDFs and supporting notice text files live in <b>G:\\Radiance\\templates\\</b>. "
        "This is a <b>read-only reference area</b> - the application reads from it but never writes "
        "output files into it."
    ))
    story += h2("Why Templates Live on the G: Drive")
    story += [
        bullet("<b>Automatic updates.</b> When IT or a supervisor updates a template, every officer "
               "gets the new version immediately. No one needs to manually copy files to their PC."),
        bullet("<b>Consistency.</b> Everyone is always working from the same approved version of every form."),
        bullet("<b>No manual maintenance.</b> Officers never need to know where templates are stored "
               "or whether their copy is current."),
    ]
    story += h2("Template Files Required")
    tdata = [
        ["File", "Purpose"],
        ["SAISIE À FAIRE_... - TEMPLATE.pdf",   "Blank reference SAISIE form used for data extraction"],
        ["K138 Seizure Cannabis-... - TEMPLATE.pdf", "Blank K138 for cannabis seizures"],
        ["K138 Seizure Knives-Arms - TEMPLATE.pdf",  "Blank K138 for knives/arms seizures"],
        ["K138 Stupefiant-Others - TEMPLATE.pdf",    "Blank K138 for all other seizure types"],
        ["Agenda - TEMPLATE.pdf",                "Blank Agenda form"],
        ["Saisie d'intérêt - TEMPLATE.pdf",      "Blank Saisie d'intérêt form"],
        ["k138_note_cannabis.txt",               "Legal notice text for cannabis K138 forms"],
        ["k138_note_arms.txt",                   "Legal notice text for knives/arms K138 forms"],
        ["k138_note_other.txt",                  "Legal notice text for other K138 forms"],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td", fontName="Helvetica", fontSize=8.5, textColor=DGRAY, leading=12))
          for c in row]
         for r, row in enumerate(tdata)],
        [2.6*inch, 3.9*inch]
    ))
    story += h2("Updating Templates")
    story.append(body("To update a template, IT staff or a designated supervisor:"))
    for i, step in enumerate([
        "Opens G:\\Radiance\\templates\\ directly.",
        "Replaces the old template file with the new version, keeping the same filename.",
        "No other action is needed - all officers will automatically use the new template on their next case.",
    ], 1):
        story.append(Paragraph(f"{i}.  {step}", sBullet))
    story.append(spacer(4))
    story.append(body(
        "Notice text files (<b>.txt</b>) can be edited in Notepad at any time to update legal "
        "language without touching the PDF templates or the application itself."
    ))

    # ── What Saves Where ──────────────────────────────────────────────────────
    story += h1("What Saves Where - Local vs. G: Drive")
    story.append(body(
        "Radiance Copilot separates personal settings (which stay on the local machine) from "
        "case data (which goes to the G: drive). This separation is intentional and important."
    ))

    story += h2("What Saves LOCALLY (on each officer's PC)")
    story.append(body(
        "<b>File:</b> radiance_copilot.cfg - stored in the same folder as the application executable."
    ))
    lcl = [
        ["Setting", "Description", "Example"],
        ["profile_role",      "The officer's role",                 "BSO, Clerk, or Supervisor"],
        ["badge_number",      "Officer's CBSA badge number",        "19747"],
        ["templates_folder",  "Path to templates folder on G:",     "G:\\Radiance\\templates"],
        ["saisie_folder",     "Last case folder browsed (convenience)", "G:\\Radiance\\cases\\..."],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th2", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td2", fontName="Helvetica", fontSize=8.5, textColor=DGRAY, leading=12))
          for c in row]
         for r, row in enumerate(lcl)],
        [1.6*inch, 2.8*inch, 2.1*inch]
    ))
    story.append(body(
        "<b>This file is configured once</b> when the app is first set up on each PC. "
        "It does not contain case data or sensitive seizure information - only paths and identity settings."
    ))

    story += h2("What Saves to the G: Drive (shared, in the case folder)")
    gdrv = [
        ["File", "Description"],
        ["Completed SAISIE PDF",          "The scanned/filled SAISIE form deposited by the BSO"],
        ["Agenda_<case>.pdf",             "Filled Agenda PDF generated by the BSO"],
        ["K138_<case>.pdf",               "Filled K138 seizure notice generated by the Clerk"],
        ["Saisie_Interet_<case>.pdf",     "Filled Saisie d'intérêt form"],
        ["barcode_<case>.png",            "Barcode image generated for the case"],
        [".extracted_data\\ (hidden)",    "Contains values_latest.json - extracted SAISIE data used to pre-fill subsequent documents"],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th3", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td3", fontName="Helvetica", fontSize=8.5, textColor=DGRAY, leading=12))
          for c in row]
         for r, row in enumerate(gdrv)],
        [2.2*inch, 4.3*inch]
    ))
    story.append(body(
        "<b>values_latest.json</b> is the key handoff file between officers. When the BSO processes "
        "a SAISIE and generates an Agenda, the extracted data is saved to this file. When the Clerk "
        "later opens the same case folder to generate the K138, the application reads from this file "
        "automatically - the Clerk does not need to re-enter any information."
    ))

    # ── Role-Based Workflow ───────────────────────────────────────────────────
    story += h1("Who Does What - Role Workflows")
    story.append(body(
        "All three roles work from the same G: drive case folders. The app controls which actions "
        "each role can perform."
    ))

    for role_title, role_desc, steps in [
        (
            "BSO - Border Services Officer",
            "The BSO is the first officer to interact with a seizure case.",
            [
                "Create the case folder on the G: drive under G:\\Radiance\\cases\\, using the standard naming convention.",
                "Copy or scan the completed SAISIE PDF into the case folder.",
                "Open Radiance Copilot. The app is already configured to point to G:\\Radiance\\templates\\.",
                "Select the SAISIE PDF from the case folder on G: drive using the Browse button.",
                "Select the seizure type (Cannabis, Knives, or Others).",
                "Click \"Process PDF.\" The app extracts all data, saves values_latest.json to the case folder, and generates the Agenda PDF and barcode.",
                "The case folder on G: drive now contains everything the Clerk needs. No emailing or copying required.",
            ]
        ),
        (
            "Clerk",
            "The Clerk's role is to generate the formal K138 seizure notice after the BSO has processed the SAISIE.",
            [
                "Open Radiance Copilot - the app opens in Clerk mode.",
                "Go to the Agenda tab and click \"Select Agenda.\" Browse to the relevant case folder on G:\\Radiance\\cases\\.",
                "The application reads values_latest.json from the hidden folder automatically - no data entry needed.",
                "Select the form type if not already pre-filled.",
                "Click \"Generate K138.\" The app reads the K138 template, fills in all case fields, appends legal notice text, and saves the completed K138 PDF to the case folder on G: drive.",
                "Optionally generate the Saisie d'intérêt form from the same data.",
            ]
        ),
        (
            "Supervisor",
            "The Supervisor role has full access to all application tabs and can open any case folder on the G: drive.",
            [
                "Reviewing completed case folders (Agenda, K138, Saisie d'intérêt) before signing off.",
                "Re-generating any document if a correction is needed.",
                "Browsing across multiple case folders to monitor workload.",
                "Updating notice text files in G:\\Radiance\\templates\\ if legal language changes.",
            ]
        ),
    ]:
        story += h2(role_title)
        story.append(body(role_desc))
        for i, step in enumerate(steps, 1):
            story.append(Paragraph(f"{i}.  {step}", sBullet))
        story.append(spacer(6))

    # ── Installation ──────────────────────────────────────────────────────────
    story += h1("Installation")
    story += h2("What Gets Installed on Each PC")
    story.append(body(
        "Radiance Copilot is distributed as a single self-contained Windows executable "
        "(<b>R-Copilot.exe</b>), built with PyInstaller. No Python installation is required "
        "on officer workstations."
    ))

    story += h2("Option A - Single Executable (Recommended)")
    for i, step in enumerate([
        "IT places R-Copilot.exe in a stable local folder on each PC, e.g. C:\\Program Files\\RadianceCopilot\\R-Copilot.exe",
        "A shortcut is created on the officer's Desktop pointing to that executable.",
        "On first launch, the officer enters their badge number, selects their role, and clicks Browse to navigate to G:\\Radiance\\templates\\. The app saves these settings to radiance_copilot.cfg.",
        "From that point on, the officer simply double-clicks the shortcut - no further configuration is needed.",
    ], 1):
        story.append(Paragraph(f"{i}.  {step}", sBullet))

    story += h2("Option B - Batch Script Launcher")
    story.append(body(
        "For environments where distributing an executable is not preferred, IT can place the "
        "Python source files on a shared location and provide each officer with a .bat launcher:"
    ))
    story.append(code_block(
        "@echo off\n"
        "cd /d \"G:\\Radiance\\app\\\"\n"
        "python saisie_a_faire_extractor.py\n"
        "pause"
    ))
    story.append(body(
        "This requires Python and required libraries (PyMuPDF, reportlab, tkinterdnd2) to be "
        "installed on each workstation."
    ))

    story += h2("First-Time Configuration Per Machine")
    cfg = [
        ["Setting",          "Value to Enter"],
        ["Templates Folder", "G:\\Radiance\\templates"],
        ["Role",             "Select appropriate role for that officer"],
        ["Badge Number",     "Officer's CBSA badge number"],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th4", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td4", fontName="Helvetica", fontSize=8.5, textColor=DGRAY))
          for c in row]
         for r, row in enumerate(cfg)],
        [2.2*inch, 4.3*inch]
    ))
    story += h2("Updating the Application")
    story.append(body(
        "When a new version of R-Copilot.exe is released, IT replaces the executable on each machine. "
        "The radiance_copilot.cfg file is preserved - officers do not need to reconfigure anything after an update."
    ))

    # ── Practical Considerations ──────────────────────────────────────────────
    story += h1("Practical Considerations")

    story += h2("File Locking - One Person Per Case at a Time")
    story.append(body(
        "Windows does not prevent two people from opening the same folder simultaneously, but "
        "writing to the same PDF at the same time can cause file corruption. The following "
        "practice should be established:"
    ))
    story += [
        bullet("Only one officer should actively process a given case at any one time. "
               "In practice this is natural: the BSO finishes their steps before the Clerk begins."),
        bullet("If two officers write to the same case folder simultaneously, the second officer will "
               "receive a file-in-use error. The fix: close any open PDFs and try again."),
        bullet("For high case volume, a simple whiteboard or shared log noting which cases are "
               "\"in progress\" is sufficient. No additional software locking is required."),
    ]

    story += h2("Offline Access - What Happens If the G: Drive Is Unavailable")
    story += [
        bullet("The application will <b>launch normally</b>, since it runs from the local PC."),
        bullet("Officers <b>will not be able to open or save case files</b> - all case data lives on the G: drive."),
        bullet("<b>No data is lost.</b> All previously saved work remains intact on the G: drive and will be accessible once the connection is restored."),
        bullet("<b>Recommended:</b> If an officer knows they will be working offline, they should copy the relevant case folder to their local desktop before disconnecting, and copy it back when reconnected."),
    ]

    story += h2("Keeping Templates Up to Date")
    story += [
        bullet("IT or a designated supervisor is responsible for G:\\Radiance\\templates\\."),
        bullet("When a form is revised, the old template PDF is replaced with the new one, keeping the same filename."),
        bullet("No officer intervention is needed - the next case processed automatically uses the new template."),
        bullet("Notice text (.txt) files can be edited in any text editor. Changes take effect immediately on the next case."),
    ]

    story += h2("Permissions Summary")
    perms = [
        ["Folder",                          "Officers",     "IT / Supervisor"],
        ["G:\\Radiance\\templates\\",        "Read-only",    "Read / Write"],
        ["G:\\Radiance\\cases\\",            "Read / Write", "Read / Write"],
        ["Local radiance_copilot.cfg",       "Read / Write (own PC)", "Read (for support)"],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th5", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td5", fontName="Helvetica", fontSize=8.5, textColor=DGRAY))
          for c in row]
         for r, row in enumerate(perms)],
        [3.0*inch, 1.75*inch, 1.75*inch]
    ))

    story += h2("Backup")
    story += [
        bullet("G:\\Radiance\\cases\\ should be included in the standard network drive backup schedule."),
        bullet("G:\\Radiance\\templates\\ should also be backed up, with version history retained if possible."),
        bullet("Local radiance_copilot.cfg files do not need to be backed up - they can be re-entered in under two minutes if a PC is replaced."),
    ]

    # ── Quick Reference ───────────────────────────────────────────────────────
    story += h1("Quick Reference")
    qr = [
        ["Question",                                        "Answer"],
        ["Where does the app run?",                         "Locally on each officer's Windows PC"],
        ["Where do case files go?",                         "G:\\Radiance\\cases\\<case-folder>\\"],
        ["Where are templates stored?",                     "G:\\Radiance\\templates\\"],
        ["Who updates templates?",                          "IT or designated supervisor (replaces files on G: drive)"],
        ["What is saved locally?",                          "Only radiance_copilot.cfg (role, badge, paths)"],
        ["How do I set up a new PC?",                       "Install exe, launch, enter badge/role, point to G:\\Radiance\\templates\\"],
        ["What if G: drive is down?",                       "App launches but cannot open/save cases; no data is lost"],
        ["Can two officers edit the same case?",            "Not simultaneously - one at a time to avoid file conflicts"],
        ["How does the Clerk get the BSO's data?",          "Automatically, via values_latest.json in the case folder"],
    ]
    story.append(table(
        [[Paragraph(c, ParagraphStyle("th6", fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE))
          if r == 0 else Paragraph(c, ParagraphStyle("td6", fontName="Helvetica", fontSize=8.5, textColor=DGRAY, leading=12))
          for c in row]
         for r, row in enumerate(qr)],
        [3.2*inch, 3.3*inch]
    ))

    # ── Footer callback ───────────────────────────────────────────────────────
    story.append(spacer(20))
    story.append(rule(MGRAY, 0.5))
    story.append(Paragraph(
        "Document prepared for internal CBSA operational deployment. "
        "For technical support, contact your regional IT desk.",
        sFooter
    ))

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#9CA3AF"))
        canvas.drawCentredString(
            letter[0] / 2,
            0.45 * inch,
            f"Radiance Copilot - G: Drive   ·   Page {doc.page}"
        )
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f"PDF saved: {OUTPUT}")

if __name__ == "__main__":
    build()
