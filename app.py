"""SOP Generator for Training Providers.

A generic, placeholder-driven SOP generator: enter company details manually or
upload supporting documents (ACRA business profile, etc.) to auto-extract them,
choose whether to show a logo, and generate a customised
SOP_<Company Name>.docx from the standard template.
"""
from datetime import date
from pathlib import Path

import streamlit as st

import sop_generator as G
from extraction import extract_company_info

ASSETS = Path(__file__).parent / "assets"
TRAINING_RECORD_XLSX = ASSETS / "Training_Record_Template.xlsx"

st.set_page_config(page_title="SOP Generator for Training Providers",
                   page_icon="📄", layout="wide")

# ----------------------------- styling --------------------------------
st.markdown(
    """
    <style>
      .app-footer {position: fixed; left: 0; bottom: 0; width: 100%;
        background: #0f1b2d; color: #cfe2f3; text-align: center;
        padding: 6px 0; font-size: 0.8rem; z-index: 999;}
      .app-footer a {color: #7fb2e6; text-decoration: none; font-weight: 600;}
      .block-container {padding-bottom: 3rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

FIELDS = ["company_name", "uen", "address", "website", "owner_name",
          "contact_name", "contact_tel", "contact_email"]
for k in FIELDS:
    st.session_state.setdefault(k, "")


# ============================ PAGE: GENERATOR =========================
def page_generate():
    st.title("📄 SOP Generator for Training Providers")
    st.caption("Generate a customised Standard Operating Procedure (Word document) "
               "from the standard template — fully branded with your company's "
               "name, logo and contact details.")

    st.subheader("1. Company Details")
    mode = st.radio("How would you like to provide company details?",
                    ["✍️ Enter manually", "📎 Upload supporting documents (auto-extract)"],
                    horizontal=True)

    if mode.startswith("📎"):
        files = st.file_uploader(
            "Upload supporting documents (ACRA business profile, etc.)",
            type=["pdf", "docx"], accept_multiple_files=True)
        if st.button("🔍 Extract details", disabled=not files):
            info, notes = extract_company_info(files)
            for k, v in info.items():
                if k in FIELDS and v:
                    st.session_state[k] = v
            for n in notes:
                st.info(n)
            if info:
                st.success("Extracted: " + ", ".join(info.keys()) +
                           ". Review and edit below before generating.")
            else:
                st.warning("Could not extract details automatically — please fill "
                           "in the fields below manually.")

    with st.form("sop_form"):
        c1, c2 = st.columns(2)
        with c1:
            company_name = st.text_input("Company Name (as registered with ACRA/ROC) *",
                                         value=st.session_state["company_name"],
                                         placeholder="Acme Training Pte. Ltd.")
            uen = st.text_input("UEN", value=st.session_state["uen"],
                                placeholder="20231234A")
            address = st.text_area("Registered Address *",
                                   value=st.session_state["address"],
                                   placeholder="123 Example Road, Singapore 123456")
            website = st.text_input("Website (optional)",
                                    value=st.session_state["website"],
                                    placeholder="www.example.com")
        with c2:
            owner_name = st.text_input("Owner / Managing Director *",
                                       value=st.session_state["owner_name"],
                                       placeholder="Jane Doe")
            contact_name = st.text_input("Support Contact Name *",
                                         value=st.session_state["contact_name"]
                                         or st.session_state["owner_name"],
                                         placeholder="Jane Doe")
            contact_tel = st.text_input("Support Tel *",
                                        value=st.session_state["contact_tel"],
                                        placeholder="+65 1234 5678")
            contact_email = st.text_input("Support Email *",
                                          value=st.session_state["contact_email"],
                                          placeholder="support@example.com")

        st.subheader("2. Logo")
        lc1, lc2 = st.columns([2, 1])
        with lc1:
            logo_file = st.file_uploader("Company logo (optional)",
                                         type=["png", "jpg", "jpeg"])
        with lc2:
            show_logo = st.toggle("Show logo on cover page", value=True)

        submitted = st.form_submit_button("⚙️ Generate SOP", type="primary")

    if submitted:
        required = {"Company Name": company_name, "Registered Address": address,
                    "Owner / Managing Director": owner_name,
                    "Support Contact Name": contact_name,
                    "Support Tel": contact_tel, "Support Email": contact_email}
        missing = [k for k, v in required.items() if not (v or "").strip()]
        if missing:
            st.error("Please fill in: " + ", ".join(missing))
            return
        info = dict(
            company_name=company_name.strip(), uen=uen.strip(),
            address=address.strip(), website=website.strip(),
            owner_name=owner_name.strip(), contact_name=contact_name.strip(),
            contact_tel=contact_tel.strip(), contact_email=contact_email.strip(),
            date=date.today().strftime("%d %b %Y"),
        )
        logo_bytes = logo_file.read() if logo_file else None
        with st.spinner("Generating SOP…"):
            buf = G.generate(info, logo_bytes=logo_bytes, show_logo=show_logo)
        fname = G.safe_filename(company_name)
        st.success(f"SOP generated: {fname}")
        st.info("ℹ️ The Table of Contents builds automatically when you open the file "
                "in Word (or right-click the contents → *Update Field*).")
        st.download_button("⬇️ Download SOP (.docx)", data=buf, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument."
                                "wordprocessingml.document")


# ====================== PAGE: SSG RTP GUIDE ==========================
def page_rtp_guide():
    st.title("📋 SSG RTP — Organisation Registration Guide")
    st.caption("Summary for first-time Training Providers applying for Organisation "
               "Registration (OR) to become an SSG-funded Registered Training Partner.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Application Fee", "S$545", "GST incl • non-refundable")
    m2.metric("Onsite Assessment", "Half-day", "by SSG officers")
    m3.metric("Track Record", "≥ 1 year", "active each quarter")
    m4.metric("Approval Needed", "OR + CA", "both must pass")

    st.divider()
    st.subheader("✅ Eligibility & Core Requirements")
    cols = st.columns(2)
    with cols[0]:
        with st.expander("🏢 Legal Entity Status", expanded=True):
            st.markdown("- Legal entity registered in Singapore (ACRA / ROS, etc.)\n"
                        "- Entity name must match official registration\n"
                        "- Cannot use proscribed terms: *National, University, "
                        "Singapore, Ministry*")
        with st.expander("📝 Declarations (5-year lookback)"):
            st.markdown("- No contractual breaches under the SSG Act 2016, "
                        "Skills Development Levy Act 1979, Private Education Act 2009\n"
                        "- No criminal convictions for dishonesty, fraud or abuse\n"
                        "- Able to adopt mandatory **Singpass e-attendance**")
    with cols[1]:
        with st.expander("💰 Financial Health", expanded=True):
            st.markdown("- Latest **IRAS Notice of Assessment (NOA)** showing "
                        "positive trade income\n"
                        "- LLPs: profit/loss allocation + partners' individual NOAs")
        with st.expander("📈 Training Track Record & Target Group"):
            st.markdown("- Minimum **1 year** of training, active **at least once "
                        "each quarter**\n"
                        "- Evidence: invoices, attendance records, trainee comms\n"
                        "- Choose target group: **Public**, **In-House**, or **Both**")

    st.divider()
    st.subheader("📎 Supporting Documents")
    s1, s2 = st.columns(2)
    with s1:
        st.markdown("**Stage 1**")
        st.markdown("- Legal registration proof (ACRA/ROS certificate)\n"
                    "- Management committee list (if a society)\n"
                    "- Completed OR declaration form\n"
                    "- Notice of Assessment (NOA)\n"
                    "- Training track-record write-up (SSG template)\n"
                    "- Evidence of training (invoices, attendance, recordings)")
    with s2:
        st.markdown("**Stage 2**")
        st.markdown("- Sample brochure / mock-up website with required disclosures\n"
                    "- Facilities write-up + photos of training/assessment rooms\n"
                    "- Proof of premises (lease / rental invoices)\n"
                    "- Sample learner registration form / contract\n"
                    "- **Policies & Operations Manual** (Course Admin + Outcomes)\n"
                    "- Requirements Specifications + Document Preparation List")
    st.success("💡 The SOP generated by this app helps fulfil the **Policies & "
               "Operations Manual** requirement under Stage 2.")

    st.divider()
    st.subheader("🛠️ Application Process")
    rtp_flow = ASSETS / "rtp_application.png"
    if rtp_flow.exists():
        fcol = st.columns([1, 2, 1])[1]
        fcol.image(str(rtp_flow), use_container_width=True,
                   caption="SSG RTP Organisation Registration — application flow")
    steps = [
        ("Set up Corppass", "Register a Corppass account (corppass.gov.sg / +65 6335 3530)."),
        ("Set up PayNow", "Establish a corporate PayNow account."),
        ("Compile documents", "Prepare all OR Stage 1 and Stage 2 documents."),
        ("Submit via TPGateway", "Submit OR application with a concurrent Course Application (CA)."),
        ("Pay OR fee", "Pay the non-refundable S$545 (GST incl.) application fee."),
        ("Onsite assessment", "SSG conducts a half-day onsite assessment of policies vs operations."),
        ("Approval", "Both OR and CA must be approved to register as an SSG-funded TP."),
    ]
    for i, (title, desc) in enumerate(steps, 1):
        st.markdown(f"**{i}. {title}** — {desc}")

    st.divider()
    st.subheader("📂 On-Site Assessment — Document Preparation (Jan 2026)")
    st.caption("Before the on-site assessment, prepare a **Policy & Operations Manual / "
               "SOPs / flowcharts** (organised by the OR requirement numbering) plus "
               "**samples of evidence** for each item below.")

    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Criterion 1 — Course Administration & Corporate Governance**")
        with st.expander("1.1  Communications & Management of Learners"):
            st.markdown(
                "- Pre-course advisory & screening: course content & relevance, fee/grant "
                "breakdown & payment modes, refund policy, attendance & admission "
                "pre-requisites, completion requirements\n"
                "- Timely responses to course queries; pre-requisite eligibility checks\n"
                "- Completed registration forms & supporting docs (CV, certificates)\n"
                "- Course-confirmation comms; vetting/approval of course info (schedule, "
                "venue, assessment, technical requirements)\n"
                "- Closed-loop comms: enquiries, feedback/complaints, withdrawals/refunds, "
                "appeals; learning-support & post-course advisory")
        with st.expander("1.2  Administration Systems"):
            st.markdown(
                "- Verification/vetting of assessment results for accuracy\n"
                "- Documented process to submit results into TPGateway; training & "
                "assessment records\n"
                "- E-attendance taking & monitoring (with intervention actions)")
        with st.expander("1.3  Corporate Governance"):
            st.markdown(
                "- Goal setting & management review: KPIs and targets for learner "
                "experience & outcomes; measurement, analysis & follow-up owners/timelines\n"
                "- Financial management: vetting of financial data & fee payments; regular "
                "monitoring of financial statements\n"
                "- Internal review: regular reviews of systems/processes & improvement plans")
        with st.expander("1.4  Management of Marketing Activities"):
            st.markdown(
                "- Vetting/approval of marketing materials (SSG Marketing Guidelines & Code "
                "of Practice)\n"
                "- Code of Conduct communicated & acknowledged by marketing reps; training "
                "on critical info\n"
                "- Management approval of promotions; monitoring & intervention actions")
        with st.expander("1.5  Management of Adult Educators (AEs)"):
            st.markdown(
                "- AE credentials & subject-matter expertise (qualifications, certs, CVs, "
                "National AE registry)\n"
                "- Induction plans; Code of Conduct & acknowledgement\n"
                "- Delivery monitoring (observations, survey analysis), appraisals & "
                "intervention actions\n"
                "- AE training records, deployment documents & annual training calendar")
    with cc2:
        st.markdown("**Criterion 2 — Outcomes**")
        with st.expander("2.1  Training Outcomes", expanded=True):
            st.markdown(
                "- Documented processes & records tracking training outcomes "
                "(e.g. **TRAQOM** survey)\n"
                "- How outcome results are used to improve training quality")
        st.markdown("**📤 Online Submission (via TPGateway)**")
        st.markdown(
            "- Submit **3 zip folders**: *Criterion 1*, *Criterion 2*, and "
            "*Policy & Operations Manual*\n"
            "- Each zip **≤ 20 MB**, no `.exe`, no encrypted files\n"
            "- Allowed formats — **Docs:** doc, docx, pdf, ppt(x), pps(x), rtf, txt, "
            "xls(x), msg • **Images:** bmp, gif, jpg/jpeg, png, tif(f)")
        st.success("💡 Tip: organise your SOP/flowcharts by the OR requirement numbering, "
                   "or include a mapping table to the OR requirements.")

    st.divider()
    st.markdown("📌 **Post-approval:** comply with SSG Terms for Training Providers, "
                "undergo ongoing Training Provider Quality Assessment (TPQA), and keep "
                "your OR profile updated (target group, PayNow, contacts, key personnel).")
    st.link_button("🔗 Official SSG TPGateway — Apply for Organisation Registration",
                   "https://www.tpgateway.gov.sg/plan-courses/organisation-registration-"
                   "for-first-time-training-provider/apply-for-organisation-registration")


# ================== PAGE: CHECKLIST & TEMPLATES ======================
DOC_CHECKLIST = {
    "1. ACRA": {
        "desc": "Proof of legal entity registration in Singapore.",
        "items": ["ACRA / BizFile Business Profile (Company)",
                  "Entity name matches official registration",
                  "Management committee list (if a registered society)"],
    },
    "2. Finance": {
        "desc": "Evidence of financial health.",
        "items": ["Latest IRAS Notice of Assessment (NOA) — positive trade income",
                  "LLP: profit/loss allocation + partners' individual NOAs"],
    },
    "3. Training Facilities": {
        "desc": "Proof of training premises and resources.",
        "items": ["Write-up on training resources & facilities",
                  "Photos of training and assessment rooms",
                  "Proof of premises (lease agreement / rental invoices)"],
    },
    "4. Training Records": {
        "desc": "Min. 1 year track record, active each quarter.",
        "items": ["Training track-record write-up (use the template below)",
                  "Invoices for past training",
                  "Attendance records / session recordings",
                  "Communications with trainees"],
    },
    "5. SOP": {
        "desc": "Policies & Operations Manual (Stage 2).",
        "items": ["Standard Operating Procedure — generate it in the *Generate SOP* page",
                  "Covers Course Administration, Training Quality & Outcomes",
                  "Sample learner registration form / contract"],
    },
}


def page_checklist():
    st.title("✅ Supporting Documents Checklist")
    st.caption("Track the five document sets required for SSG Organisation "
               "Registration. Tick items as you prepare them.")

    total = sum(len(v["items"]) for v in DOC_CHECKLIST.values())
    done = 0
    cols = st.columns(2)
    for idx, (cat, data) in enumerate(DOC_CHECKLIST.items()):
        with cols[idx % 2]:
            with st.container(border=True):
                st.markdown(f"### {cat}")
                st.caption(data["desc"])
                cat_done = 0
                for j, item in enumerate(data["items"]):
                    checked = st.checkbox(item, key=f"chk_{idx}_{j}")
                    done += checked
                    cat_done += checked
                st.progress(cat_done / len(data["items"]),
                            text=f"{cat_done}/{len(data['items'])} ready")

    st.divider()
    pct = int(done / total * 100) if total else 0
    st.subheader(f"Overall readiness: {done}/{total} ({pct}%)")
    st.progress(done / total if total else 0)
    if done == total:
        st.success("🎉 All supporting documents are ready for your OR application!")

    st.divider()
    st.subheader("📥 Downloadable Templates")
    t1, t2 = st.columns(2)
    with t1:
        st.markdown("**Training Records Write-Up (SSG)**")
        st.caption("Template to document your training track record for OR submission.")
        if TRAINING_RECORD_XLSX.exists():
            st.download_button(
                "⬇️ Download Training Record Template (.xlsx)",
                data=TRAINING_RECORD_XLSX.read_bytes(),
                file_name="Training_Record_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet")
        else:
            st.warning("Template file not found.")
    with t2:
        st.markdown("**Standard Operating Procedure (SOP)**")
        st.caption("Generate your branded SOP in the *Generate SOP* page.")
        st.info("→ Use the **Generate SOP** menu item.")


# ----------------------------- nav -----------------------------------
st.sidebar.title("🗂️ Menu")
page = st.sidebar.radio("Go to", ["📄 Generate SOP", "✅ Documents Checklist",
                                  "📋 SSG RTP Registration Guide"])
st.sidebar.divider()
st.sidebar.caption("Generate a Standard Operating Procedure for SkillsFuture "
                   "Registered Training Partners, and learn how to register with SSG.")

if page.startswith("📄"):
    page_generate()
elif page.startswith("✅"):
    page_checklist()
else:
    page_rtp_guide()

# ----------------------------- footer --------------------------------
st.markdown(
    '<div class="app-footer">Powered by '
    '<a href="https://www.tertiaryinfotech.com/" target="_blank">'
    'Tertiary Infotech Academy Pte Ltd</a></div>',
    unsafe_allow_html=True,
)
