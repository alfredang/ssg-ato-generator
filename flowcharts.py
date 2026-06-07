"""Generate clean SOP diagrams (org charts + process flowcharts) as PNGs.

Diagrams are recreated from the InnoHat and HighSpark source SOPs and are fully
parameterised: the owner/company name is injected at generation time, so the
shipped SOP template contains only placeholders and never any real names.

Process boxes are rounded, decisions are diamonds, terminals are ovals, with
consistent Helvetica typography and a navy/blue/green palette.
"""
import os
from graphviz import Digraph

os.environ["PATH"] = "/opt/homebrew/bin:" + os.environ.get("PATH", "")

# --- palette ---
NAVY = "#1F3A5F"
BLUE = "#2E75B6"
LIGHT = "#EAF2FB"
DECISION = "#CFE2F3"
GREEN = "#4F9D3A"
TERM = "#DDE3EA"
FONT = "Helvetica"

PLACEHOLDER_OWNER = "[Owner / Managing Director]"
STAFF = "[Name]"


def _base(name):
    g = Digraph(name, format="png")
    g.attr(rankdir="TB", bgcolor="white", nodesep="0.35", ranksep="0.55")
    g.attr("node", fontname=FONT, fontsize="11", color=NAVY, penwidth="1.3")
    g.attr("edge", fontname=FONT, fontsize="10", color=NAVY, penwidth="1.2",
           arrowsize="0.8")
    g.attr(dpi="200")
    return g


def _proc(g, nid, text, fill=LIGHT):
    g.node(nid, text, shape="box", style="rounded,filled", fillcolor=fill)


def _decision(g, nid, text):
    g.node(nid, text, shape="diamond", style="filled", fillcolor=DECISION,
           margin="0.05")


def _terminal(g, nid, text):
    g.node(nid, text, shape="oval", style="filled", fillcolor=TERM,
           fontname=FONT + " Bold")


def _render(g, name, out_dir):
    return g.render(filename=name, directory=out_dir, cleanup=True)


# =====================================================================
# Org charts
# =====================================================================
def org_management(out_dir, owner):
    g = _base("org_management")
    g.attr(label="Organisation Chart of Management Team", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    g.node("md", f"<<b>{owner}</b><br/><font point-size='10'>Managing Director / Owner</font>>",
           shape="box", style="filled", fillcolor=NAVY, fontcolor="white")
    depts = ["HR &amp; Admin", "Operations", "Finance", "Sales", "Digital Marketing"]
    for i, d in enumerate(depts):
        did = f"d{i}"
        g.node(did, f"<<b>{d}</b>>", shape="box", style="filled",
               fillcolor=BLUE, fontcolor="white")
        g.edge("md", did, arrowhead="none")
        prev = did
        for j, role in enumerate(["Manager", "Assistant"]):
            sid = f"{did}s{j}"
            g.node(sid, f"<<b>{role}</b><br/><font point-size='9'>{STAFF}</font>>",
                   shape="box", style="filled", fillcolor=GREEN, fontcolor="white")
            g.edge(prev, sid, arrowhead="none")
            prev = sid
    return _render(g, "org_management", out_dir)


def org_training(out_dir, owner):
    g = _base("org_training")
    g.attr(label="Organisation Chart of Training Team", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    g.node("tm", f"<<b>Training Manager &amp;<br/>Management Representative</b>"
                 f"<br/><font point-size='10'>{owner}</font>>",
           shape="box", style="filled", fillcolor=NAVY, fontcolor="white")
    g.node("fin", f"<<b>Finance Manager</b><br/><font point-size='9'>{STAFF}</font>>",
           shape="box", style="filled", fillcolor=BLUE, fontcolor="white")
    g.edge("tm", "fin", arrowhead="none")
    roles = ["HR &amp; Admin Manager", "Training Co-ordinator", "Trainers / Assessors",
             "Course Developers", "Curriculum Developer<br/>(Associate)",
             "Subject Matter Expert<br/>(Associate)"]
    for i, r in enumerate(roles):
        rid = f"r{i}"
        g.node(rid, f"<<b>{r}</b><br/><font point-size='9'>{STAFF}</font>>",
               shape="box", style="filled", fillcolor=BLUE, fontcolor="white")
        g.edge("tm", rid, arrowhead="none")
    return _render(g, "org_training", out_dir)


# =====================================================================
# Flowcharts
# =====================================================================
def refund(out_dir, owner=None):
    g = _base("refund")
    g.attr(label="Refund Process", labelloc="t", fontname=FONT + " Bold", fontsize="14")
    _proc(g, "a", "Receive Refund\nRequest in Writing")
    _proc(g, "b", "Investigate Reason\nfor Refund")
    _decision(g, "c", "Received 3 days\nbefore Course\nStart Date?")
    _decision(g, "d", "Approved?")
    _proc(g, "e", "Submit to Training\nManager for Approval")
    _proc(g, "f", "Issue Refund &\nUpdate Customer\nwithin 30 Days")
    _proc(g, "gg", "Reject Request &\nUpdate Customer\nwithin 5 Days")
    _proc(g, "h", "Inform Customer\nReason for Rejection\nwithin 5 Days")
    _terminal(g, "end1", "End")
    _terminal(g, "end2", "End")
    g.edge("a", "b"); g.edge("b", "c")
    g.edge("c", "e", label="Yes"); g.edge("c", "gg", label="No")
    g.edge("e", "d"); g.edge("d", "f", label="Yes"); g.edge("d", "h", label="No")
    g.edge("f", "end1"); g.edge("gg", "end2"); g.edge("h", "end2")
    return _render(g, "refund", out_dir)


def enquiry(out_dir, owner=None):
    g = _base("enquiry")
    g.attr(label="Pre-course Enquiry & Registration", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    _proc(g, "a", "Enquiry received via\nEmail or Course-page Webform")
    _proc(g, "b", "Send automated acknowledgement\n(reply within 1-2 working days)")
    _proc(g, "c", "Learning consultant gives pre-course\nadvisory: course, dates, fees, funding,\neligibility & student contract")
    _decision(g, "d", "Trainee confirms\ninterest & signs\nstudent contract?")
    _proc(g, "e", "Send brochure: course title, objectives,\ntrainer profile, schedule, venue/virtual\nlink, payment & refund policy, contact")
    _proc(g, "f", "Admin cc-ed to issue invoice &\nSkillsFuture registration guide")
    _decision(g, "gg", "Trainee makes\ndeposit / payment?")
    _proc(g, "h", "Send goodbye email &\nrequest for feedback")
    _proc(g, "i", "Send reminder")
    _terminal(g, "reg", "Enrolment\nConfirmed")
    g.edge("a", "b"); g.edge("b", "c"); g.edge("c", "d")
    g.edge("d", "h", label="No"); g.edge("d", "e", label="Yes")
    g.edge("e", "f"); g.edge("f", "gg")
    g.edge("gg", "i", label="No"); g.edge("i", "gg", style="dashed")
    g.edge("gg", "reg", label="Yes")
    return _render(g, "enquiry", out_dir)


def funding(out_dir, owner=None):
    g = _base("funding")
    g.attr(label="Course Funding & Enrolment", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    _proc(g, "a", "Upon successful enrolment, admin liaises\nwith trainee on sponsorship &\nSkillsFuture credit usage")
    _decision(g, "b", "Self-sponsored or\nemployer-sponsored?")
    _proc(g, "c", "Self-funded:\nGuide steps to claim via MySkillsFuture\nPortal (within 60 days before start);\nrequest contact information")
    _proc(g, "d", "Employer-sponsored:\nClient pays net fee (after SSG grant);\nrequest sponsoring employer's details\nfor the trainees")
    _terminal(g, "e", "Proceed to\nCourse Run")
    g.edge("a", "b"); g.edge("b", "c", label="Self"); g.edge("b", "d", label="Employer")
    g.edge("c", "e"); g.edge("d", "e")
    return _render(g, "funding", out_dir)


def course_confirmation(out_dir, owner=None):
    g = _base("course_confirmation")
    g.attr(label="Course Run Confirmation & Reminders", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    _proc(g, "a", "2 weeks before\ncourse start date")
    _decision(g, "b", "Course run\nconfirmed?")
    _proc(g, "c", "Send 1st reminder email. Include:\nvenue & directions, reporting time,\nbring NRIC & Singpass, items to bring")
    _proc(g, "d", "2 days before\ncourse start date")
    _proc(g, "e", "Send 2nd reminder email. Include:\nworksheet & digital exercise files,\nvenue, reporting time, NRIC & Singpass")
    _proc(g, "f", "Notify enrolled trainees of cancellation;\nrefund deposit & course fee (5 working\ndays); SkillsFuture credits not deducted")
    _terminal(g, "g1", "Proceed to\nCourse Delivery")
    _terminal(g, "g2", "End")
    g.edge("a", "b")
    g.edge("b", "c", label="Yes"); g.edge("b", "f", label="No / Cancelled")
    g.edge("c", "d"); g.edge("d", "e"); g.edge("e", "g1"); g.edge("f", "g2")
    return _render(g, "course_confirmation", out_dir)


def trainer_performance(out_dir, owner=None):
    g = _base("trainer_performance")
    g.attr(label="Trainer Performance Management", labelloc="t",
           fontname=FONT + " Bold", fontsize="14")
    _proc(g, "a", "Identify Performance Issue\n(e.g. satisfaction score < 70%)")
    _proc(g, "b", "Document Performance Concerns")
    _decision(g, "c", "Does Trainer\nAcknowledge Issue?")
    _proc(g, "d", "Provide coaching / support &\nset improvement plan")
    _decision(g, "e", "Is Performance\nSatisfactory?")
    _proc(g, "f", "Termination of Trainer\nif Necessary")
    _terminal(g, "g1", "Retain Trainer")
    _terminal(g, "g2", "End")
    g.edge("a", "b"); g.edge("b", "c")
    g.edge("c", "d", label="Yes"); g.edge("c", "f", label="No")
    g.edge("d", "e"); g.edge("e", "g1", label="Yes"); g.edge("e", "f", label="No")
    g.edge("f", "g2")
    return _render(g, "trainer_performance", out_dir)


def rtp_application(out_dir, owner=None):
    """SSG RTP Organisation Registration application flow (static guide diagram)."""
    g = _base("rtp_application")
    g.attr(label="SSG RTP — Organisation Registration Application Flow",
           labelloc="t", fontname=FONT + " Bold", fontsize="14")
    _terminal(g, "start", "First-time\nTraining Provider")
    _proc(g, "a", "Set up Corppass account")
    _proc(g, "b", "Set up corporate PayNow")
    _proc(g, "c", "Compile OR Stage 1 & 2 documents\n(incl. Policy & Operations Manual / SOP)")
    _proc(g, "d", "Submit OR + concurrent Course\nApplication (CA) via TPGateway")
    _proc(g, "e", "Pay OR fee: S$545\n(GST incl., non-refundable)")
    _proc(g, "f", "SSG half-day on-site assessment\n(policies vs actual operations)")
    _decision(g, "gg", "Both OR & CA\napproved?")
    _terminal(g, "reg", "Registered SSG-funded\nTraining Partner")
    _proc(g, "fail", "Application Unsuccessful\n(reapply after addressing gaps)", fill="#F8D7DA")
    g.edge("start", "a"); g.edge("a", "b"); g.edge("b", "c"); g.edge("c", "d")
    g.edge("d", "e"); g.edge("e", "f"); g.edge("f", "gg")
    g.edge("gg", "reg", label="Yes"); g.edge("gg", "fail", label="No")
    return _render(g, "rtp_application", out_dir)


# marker name -> builder
DIAGRAMS = {
    "org_management": org_management,
    "org_training": org_training,
    "enquiry": enquiry,
    "funding": funding,
    "course_confirmation": course_confirmation,
    "refund": refund,
    "trainer_performance": trainer_performance,
}


def build_all(out_dir, owner=PLACEHOLDER_OWNER):
    """Render every diagram with the given owner name; return {key: png_path}."""
    os.makedirs(out_dir, exist_ok=True)
    return {key: fn(out_dir, owner) for key, fn in DIAGRAMS.items()}


if __name__ == "__main__":
    import sys
    d = sys.argv[1] if len(sys.argv) > 1 else "/tmp/diagrams"
    for k, p in build_all(d).items():
        print(k, "->", p)
