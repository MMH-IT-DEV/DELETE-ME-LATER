import copy
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def w(tag, attrs=None):
    return ET.Element(qn(W_NS, tag), attrs or {})


def wt(tag, text=None, attrs=None):
    node = w(tag, attrs)
    if text is not None:
        node.text = text
    return node


URLS = {
    "SECURITY_TRACKER_SHEET": "https://docs.google.com/spreadsheets/d/1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ/edit",
    "HEALTH_MONITOR_SHEET": "https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7splmRS7HYJTR77GpfI5w/edit",
    "SECURITY_TRACKER_SCRIPT": "https://script.google.com/u/0/home/projects/1Fp0ooeKm028-0XYu5X5CCFrxzILbDEPxwX6Si_I7c3A3GcwmEnANNb4k/edit",
    "HEALTH_MONITOR_SCRIPT": "https://script.google.com/u/0/home/projects/1TBkee_JgNKnHxxeCbWSp3uJyh5Wg5o_GLowbp0k51DP98GjzzDWN4eS5/edit",
    "IT008_DOC": "https://docs.google.com/document/d/1KvqBxBo0bpkqGtFrwPWpYTWdb8TQdnY1TTuqT5iZ6Go/edit",
    "IT010_DOC": "https://docs.google.com/document/d/13PHfqd7Z9Ued0jdjTiTS1A2UTTS3K-h0AMX6ujFkAZg/edit",
}


class RelationshipManager:
    def __init__(self, rels_xml):
        self.root = ET.fromstring(rels_xml)
        self.existing = {}
        max_id = 0
        for rel in self.root.findall(qn(PKG_REL_NS, "Relationship")):
            rel_id = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            mode = rel.attrib.get("TargetMode", "")
            if target and mode == "External":
                self.existing[target] = rel_id
            match = re.match(r"rId(\d+)", rel_id)
            if match:
                max_id = max(max_id, int(match.group(1)))
        self.next_id = max_id + 1

    def external(self, target):
        if target in self.existing:
            return self.existing[target]
        rel_id = f"rId{self.next_id}"
        self.next_id += 1
        ET.SubElement(
            self.root,
            qn(PKG_REL_NS, "Relationship"),
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                "Target": target,
                "TargetMode": "External",
            },
        )
        self.existing[target] = rel_id
        return rel_id

    def to_xml(self):
        return ET.tostring(self.root, encoding="utf-8", xml_declaration=True)


class DocBuilder:
    def __init__(self, sect_pr, rels):
        self.sect_pr = copy.deepcopy(sect_pr)
        self.rels = rels
        self.body_items = []

    def blank(self, after="120"):
        p = w("p")
        ppr = ET.SubElement(p, qn(W_NS, "pPr"))
        ET.SubElement(ppr, qn(W_NS, "spacing"), {qn(W_NS, "after"): after})
        self.body_items.append(p)

    def title(self, text):
        self.body_items.append(
            self._paragraph(
                [{"text": text, "bold": True, "size": 52}],
                align="center",
                after="160",
            )
        )

    def meta(self, label, value):
        self.body_items.append(
            self._paragraph(
                [
                    {"text": label, "bold": True, "size": 22},
                    {"text": value, "underline": True, "size": 22},
                ],
                align="right",
                after="80",
            )
        )

    def meta_link(self, label, text, url):
        self.body_items.append(
            self._paragraph(
                [
                    {"text": label, "bold": True, "size": 22},
                    {"text": text, "url": url, "size": 22},
                ],
                align="right",
                after="80",
            )
        )

    def heading(self, text):
        self.body_items.append(
            self._paragraph(
                [{"text": text, "bold": True, "size": 28, "color": "434343"}],
                after="120",
            )
        )

    def subheading(self, text):
        self.body_items.append(
            self._paragraph([{"text": text, "bold": True, "size": 22}], after="80")
        )

    def para(self, text):
        self.body_items.append(self._paragraph([{"text": text, "size": 22}]))

    def bullet(self, text):
        self.body_items.append(self._paragraph([{"text": "- " + text, "size": 22}]))

    def link_para(self, text, url):
        self.body_items.append(self._paragraph([{"text": text, "url": url, "size": 22}]))

    def table(self, rows, widths=None, header=True):
        self.body_items.append(self._table(rows, widths=widths, header=header))

    def summary_table(self, rows):
        table_rows = []
        for label, paragraphs in rows:
            table_rows.append([label, paragraphs])
        self.body_items.append(self._table(table_rows, widths=[1900, 7460], header=False, summary=True))

    def document_xml(self):
        root = w("document")
        body = ET.SubElement(root, qn(W_NS, "body"))
        for item in self.body_items:
            body.append(item)
        body.append(copy.deepcopy(self.sect_pr))
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _paragraph(self, parts, align="left", after="100", line="276"):
        p = w("p")
        ppr = ET.SubElement(p, qn(W_NS, "pPr"))
        ET.SubElement(
            ppr,
            qn(W_NS, "spacing"),
            {
                qn(W_NS, "after"): after,
                qn(W_NS, "line"): line,
                qn(W_NS, "lineRule"): "auto",
            },
        )
        if align != "left":
            ET.SubElement(ppr, qn(W_NS, "jc"), {qn(W_NS, "val"): align})
        for part in parts:
            if part.get("url"):
                rel_id = self.rels.external(part["url"])
                hyperlink = ET.SubElement(
                    p,
                    qn(W_NS, "hyperlink"),
                    {qn(R_NS, "id"): rel_id, qn(W_NS, "history"): "1"},
                )
                hyperlink.append(self._run(part))
            else:
                p.append(self._run(part))
        return p

    def _run(self, part):
        r = w("r")
        rpr = ET.SubElement(r, qn(W_NS, "rPr"))
        ET.SubElement(
            rpr,
            qn(W_NS, "rFonts"),
            {
                qn(W_NS, "ascii"): "Arial",
                qn(W_NS, "hAnsi"): "Arial",
                qn(W_NS, "cs"): "Arial",
            },
        )
        if part.get("bold"):
            ET.SubElement(rpr, qn(W_NS, "b"), {qn(W_NS, "val"): "1"})
        if part.get("italic"):
            ET.SubElement(rpr, qn(W_NS, "i"), {qn(W_NS, "val"): "1"})
        if part.get("underline"):
            ET.SubElement(rpr, qn(W_NS, "u"), {qn(W_NS, "val"): "single"})
        if part.get("url"):
            ET.SubElement(rpr, qn(W_NS, "u"), {qn(W_NS, "val"): "single"})
            ET.SubElement(rpr, qn(W_NS, "color"), {qn(W_NS, "val"): "1155CC"})
        elif part.get("color"):
            ET.SubElement(rpr, qn(W_NS, "color"), {qn(W_NS, "val"): part["color"]})
        ET.SubElement(rpr, qn(W_NS, "sz"), {qn(W_NS, "val"): str(part.get("size", 22))})
        ET.SubElement(rpr, qn(W_NS, "szCs"), {qn(W_NS, "val"): str(part.get("size", 22))})
        text = part.get("text", "")
        t = wt("t", text)
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        return r

    def _cell_paragraph(self, item, header=False):
        if isinstance(item, str):
            parts = [{"text": item, "size": 22, "bold": header}]
            return self._paragraph(parts, align="center" if header else "left")
        if isinstance(item, dict):
            return self._paragraph(item["parts"], align=item.get("align", "left"))
        return self._paragraph([{"text": str(item), "size": 22}])

    def _table(self, rows, widths=None, header=True, summary=False):
        tbl = w("tbl")
        tbl_pr = ET.SubElement(tbl, qn(W_NS, "tblPr"))
        ET.SubElement(tbl_pr, qn(W_NS, "tblW"), {qn(W_NS, "w"): "9360", qn(W_NS, "type"): "dxa"})
        ET.SubElement(tbl_pr, qn(W_NS, "jc"), {qn(W_NS, "val"): "left"})
        borders = ET.SubElement(tbl_pr, qn(W_NS, "tblBorders"))
        border_size = "8" if summary else "4"
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            ET.SubElement(
                borders,
                qn(W_NS, edge),
                {
                    qn(W_NS, "val"): "single",
                    qn(W_NS, "sz"): border_size,
                    qn(W_NS, "space"): "0",
                    qn(W_NS, "color"): "000000",
                },
            )
        grid = ET.SubElement(tbl, qn(W_NS, "tblGrid"))
        if widths:
            for width in widths:
                ET.SubElement(grid, qn(W_NS, "gridCol"), {qn(W_NS, "w"): str(width)})
        for row_index, row in enumerate(rows):
            tr = ET.SubElement(tbl, qn(W_NS, "tr"))
            for col_index, cell in enumerate(row):
                tc = ET.SubElement(tr, qn(W_NS, "tc"))
                tc_pr = ET.SubElement(tc, qn(W_NS, "tcPr"))
                if widths:
                    ET.SubElement(
                        tc_pr,
                        qn(W_NS, "tcW"),
                        {qn(W_NS, "w"): str(widths[col_index]), qn(W_NS, "type"): "dxa"},
                    )
                tc_mar = ET.SubElement(tc_pr, qn(W_NS, "tcMar"))
                for edge in ("top", "left", "bottom", "right"):
                    ET.SubElement(tc_mar, qn(W_NS, edge), {qn(W_NS, "w"): "100", qn(W_NS, "type"): "dxa"})
                ET.SubElement(tc_pr, qn(W_NS, "vAlign"), {qn(W_NS, "val"): "top"})
                paragraphs = cell if isinstance(cell, list) else [cell]
                for para in paragraphs:
                    tc.append(self._cell_paragraph(para, header=header and row_index == 0))
        return tbl


def extract_sect_pr(doc_xml):
    root = ET.fromstring(doc_xml)
    body = root.find(qn(W_NS, "body"))
    return copy.deepcopy(body.find(qn(W_NS, "sectPr")))


def hyperlink_item(text, url):
    return {"parts": [{"text": text, "url": url, "size": 22}]}


def text_item(text):
    return {"parts": [{"text": text, "size": 22}]}


def build_it014(builder):
    builder.title("IT-014: Automation Maintenance")
    builder.blank()
    builder.meta("Prepared by: ", "Erik Demchuk")
    builder.meta("Reviewed by: ", "_______________________________")
    builder.meta("QA Approval: ", "_______________________________")
    builder.meta("Effective Date: ", "March 11, 2026")
    builder.blank()
    builder.summary_table(
        [
            ("Department:", [text_item("Information Technology")]),
            ("SOP ID:", [text_item("IT-014")]),
            ("Version:", [text_item("1.5")]),
            (
                "Resources:",
                [
                    hyperlink_item("Systems Health Monitor spreadsheet", URLS["HEALTH_MONITOR_SHEET"]),
                    hyperlink_item("Systems Health Monitor Apps Script", URLS["HEALTH_MONITOR_SCRIPT"]),
                    hyperlink_item("Security & GMP Connection Tracker spreadsheet", URLS["SECURITY_TRACKER_SHEET"]),
                    hyperlink_item("Security & GMP Connection Tracker Apps Script", URLS["SECURITY_TRACKER_SCRIPT"]),
                ],
            ),
            (
                "Important Notes:",
                [
                    text_item("- Use the System Registry tab as the source of truth for active monitored systems."),
                    text_item("- Do not store secrets in the workbook. Record only type, location, and expiry information."),
                    text_item("- Keep the Maintenance Guide link current for every active system row."),
                ],
            ),
        ]
    )
    builder.blank()

    builder.heading("1. Purpose")
    builder.para("To define procedures for maintaining, monitoring, and troubleshooting all registered systems in the Systems Health Monitor, including automations, dashboards, scripts, and bots tracked by IT.")
    builder.para("This SOP covers registry maintenance, incident logging, credential expiry review, and periodic validation for GMP-critical systems.")
    builder.blank()

    builder.heading("2. Scope")
    builder.para("This SOP applies to IT personnel and system owners who maintain the Systems Health Monitor and respond to issues involving registered automations or dashboards.")
    builder.bullet("Managing entries in the System Registry tab")
    builder.bullet("Responding to health warnings and failures")
    builder.bullet("Renewing expiring credentials and authentication tokens")
    builder.bullet("Logging and following up on incidents")
    builder.bullet("Validating GMP-critical systems on the defined review cycle")
    builder.blank()

    builder.heading("3. Responsibilities")
    builder.table(
        [
            ["Role", "Responsibility"],
            ["IT Administrator", "Maintains the Systems Health Monitor, updates system records, reviews health warnings, and records maintenance actions."],
            ["System Owner", "Confirms system purpose, run frequency, current status, and maintenance requirements for assigned systems."],
            ["IT Manager", "Approves significant changes to GMP-critical systems, reviews escalations, and confirms the register remains accurate and usable."],
            ["QA Manager", "Approves SOP revisions and supports validation expectations for GMP-critical automations and dashboards."],
        ],
        widths=[2200, 7160],
    )
    builder.blank()

    builder.heading("4. Definitions")
    builder.table(
        [
            ["Term", "Definition"],
            ["Systems Health Monitor", "Workbook used to register active automations, log incidents, and retain heartbeat or maintenance history."],
            ["System Registry", "Primary tab listing each automation, bot, script, or dashboard that must be tracked by IT."],
            ["Heartbeat Log", "Operational log used to record successful runs, warnings, failures, recoveries, and maintenance notes."],
            ["Maintenance Guide", "Live SOP or maintenance document link used to maintain a specific automation or system."],
            ["GMP Critical", "Flag showing whether the system supports GMP-relevant operations and therefore requires tighter review and validation."],
        ],
        widths=[2400, 6960],
    )
    builder.blank()

    builder.heading("5. Procedure")
    builder.subheading("5.1  Daily Health Check Response")
    builder.para("Review the Systems Health Monitor each working day and confirm warning or failure conditions are assigned and being acted on.")
    builder.bullet("Review the System Registry for systems showing warnings, overdue heartbeats, failed runs, or credentials approaching expiry.")
    builder.bullet("Check open incidents and confirm each row has an owner, current status, and documented next action.")
    builder.bullet("Review Heartbeat Log entries when you need to confirm whether a system has been running normally or recently changed.")
    builder.blank()

    builder.subheading("5.2  Add or Update a System Record")
    builder.bullet("Add a new row or update the existing row in System Registry whenever a system is introduced, retired, materially changed, or reassigned.")
    builder.bullet("Complete the core fields: System Name, Description, Type, Platform, GMP Critical, Run Frequency, Maintenance Guide, Auth Type, Expiry Date, Owner, Validated, Status, Last Heartbeat, and Last Run OK.")
    builder.bullet("Use support fields such as Auth Location, Last Validated, Heartbeat Method, and Notes whenever needed for maintainability.")
    builder.bullet("The Maintenance Guide field must contain the live SOP or maintenance document link for that system.")
    builder.blank()

    builder.subheading("5.3  Routine Maintenance")
    builder.bullet("Renew credentials before the Expiry Date and update the related tracker fields after renewal is complete.")
    builder.bullet("After a planned maintenance change, confirm the system returns to normal operation and update Status, Last Heartbeat, Last Run OK, or Notes as needed.")
    builder.bullet("Record significant maintenance actions, recoveries, failures, or verification outcomes in Heartbeat Log so there is a usable operating history.")
    builder.blank()

    builder.subheading("5.4  Log and Manage Incidents")
    builder.bullet("When a system issue is identified, add or update a row in Incident Reports or the live incident intake used by the workbook.")
    builder.bullet("Record the event time, affected system, severity, current status, assigned owner, and resolution notes.")
    builder.bullet("If a system is Down, update both the incident row and the related System Registry status so the workbook remains consistent.")
    builder.blank()

    builder.subheading("5.5  Review GMP-Critical Systems")
    builder.bullet("Review every GMP-critical system on the defined review cycle and whenever a significant change occurs.")
    builder.bullet("Confirm the owner, maintenance guide link, authentication tracking, run frequency, validation state, and most recent operating evidence are current.")
    builder.bullet("Update Validated and Last Validated when the review is completed.")
    builder.blank()

    builder.subheading("5.6  Retire or Remove a System")
    builder.bullet("Confirm the system is no longer needed and that dependent maintenance guides or support documentation are updated first.")
    builder.bullet("Remove the system from active use in the System Registry or clearly mark it as retired according to the operating decision for that system.")
    builder.bullet("Record the retirement action in Heartbeat Log and close any related open incident items.")
    builder.blank()

    builder.subheading("5.7  Keep the Tracker Usable")
    builder.bullet("Keep the workbook links, system links, ownership, and maintenance fields current so the tracker remains actionable.")
    builder.bullet("Avoid removing detail that is still needed for maintenance, review, or auditability.")
    builder.bullet("If a field is no longer used operationally, remove or hide it only after the replacement process and recordkeeping approach are defined.")
    builder.blank()

    builder.heading("6. References")
    builder.link_para("Systems Health Monitor spreadsheet", URLS["HEALTH_MONITOR_SHEET"])
    builder.link_para("Systems Health Monitor Apps Script", URLS["HEALTH_MONITOR_SCRIPT"])
    builder.link_para("Security & GMP Connection Tracker spreadsheet", URLS["SECURITY_TRACKER_SHEET"])
    builder.link_para("Security & GMP Connection Tracker Apps Script", URLS["SECURITY_TRACKER_SCRIPT"])
    builder.link_para("FDA 21 CFR Part 11 - Electronic Records; Electronic Signatures", "https://www.ecfr.gov/current/title-21/chapter-I/subchapter-A/part-11")
    builder.blank()

    builder.heading("7. Attachments")
    builder.table(
        [
            ["Attachment", "Summary"],
            ["Attachment A", "System Registry core operational fields and support fields."],
            ["Attachment B", "Incident and heartbeat log field definitions."],
            ["Attachment C", "Status response guide."],
        ],
        widths=[2000, 7360],
    )
    builder.blank()
    builder.subheading("Attachment A - System Registry Core Operational Fields")
    builder.table(
        [
            ["Field", "Use"],
            ["System Name", "Unique name used to identify the automation or monitored system."],
            ["Description", "Plain-language description of what the system does."],
            ["Type", "Operational grouping such as FLOW, SCRIPT, BOT, or DASHBOARD."],
            ["Platform", "Primary platform or integration surface, such as Shopify, Google, Katana, or FedEx."],
            ["GMP Critical", "Yes or No indicator used to identify systems requiring tighter review and control."],
            ["Run Frequency", "Expected cadence, event trigger, or schedule for the system."],
            ["Maintenance Guide", "Link to the system-specific SOP or procedure used to maintain that system."],
            ["Auth Type", "Authentication method such as None, API Key, Bearer Token, or other credential class."],
            ["Expiry Date", "Credential expiry date when applicable."],
            ["Owner", "Person accountable for business or IT ownership of the system."],
            ["Validated", "Current validation state for the system."],
            ["Status", "Current operating status, such as Healthy, Degraded, Down, or Unknown."],
            ["Last Heartbeat", "Most recent timestamp showing the system was confirmed to have run or checked in."],
            ["Last Run OK", "Whether the most recent run completed successfully."],
        ],
        widths=[2500, 6860],
    )
    builder.blank()
    builder.subheading("Attachment B - Incident and Heartbeat Log Fields")
    builder.table(
        [
            ["Field", "Use"],
            ["Timestamp", "Date and time the event or incident was recorded."],
            ["System", "System or automation associated with the event."],
            ["Severity or Status", "Operational priority and current state of the event or incident."],
            ["Assigned To", "Person responsible for investigation or follow-up."],
            ["Action", "Short description of the action that occurred or was taken."],
            ["Details", "Additional context describing the event, issue, or maintenance outcome."],
        ],
        widths=[2500, 6860],
    )
    builder.blank()
    builder.subheading("Attachment C - Status Response Guide")
    builder.table(
        [
            ["Status", "Meaning / Response"],
            ["Healthy", "System is operating normally. Continue routine monitoring."],
            ["Degraded", "System is running with a warning condition. Review promptly and correct before it becomes an outage."],
            ["Down", "System is not operating as expected. Open or update an incident and assign ownership immediately."],
            ["Unknown", "The tracker does not currently have enough evidence to confirm health. Investigate until status is clarified."],
        ],
        widths=[1800, 7560],
    )
    builder.blank()

    builder.heading("8. Revision History")
    builder.table(
        [
            ["Version", "Effective Date", "Description of Change", "Change Control #"],
            ["1.5", "March 11, 2026", "Rewrote the local Word version, corrected live links, and aligned the content to the current Systems Health Monitor workflow.", "Pending"],
            ["1.4", "March 6, 2026", "Reformatted the SOP to the corporate template and aligned the procedure to the live Systems Health Tracker.", "Pending"],
            ["1.3", "March 4, 2026", "Full SOP rewrite for the automation maintenance process.", "n/a"],
        ],
        widths=[1100, 1700, 4300, 2260],
    )


def build_it015(builder):
    builder.title("IT-015: IT Security Tracker Maintenance")
    builder.blank()
    builder.meta("Prepared by: ", "Erik Demchuk")
    builder.meta("Reviewed by: ", "_______________________________")
    builder.meta("QA Approval: ", "_______________________________")
    builder.meta("Effective Date: ", "March 11, 2026")
    builder.blank()
    builder.summary_table(
        [
            ("Department:", [text_item("Information Technology")]),
            ("SOP ID:", [text_item("IT-015")]),
            ("Version:", [text_item("1.5")]),
            (
                "Resources:",
                [
                    hyperlink_item("Security & GMP Connection Tracker spreadsheet", URLS["SECURITY_TRACKER_SHEET"]),
                    hyperlink_item("Security Tracker Apps Script", URLS["SECURITY_TRACKER_SCRIPT"]),
                    hyperlink_item("Systems Health Monitor spreadsheet", URLS["HEALTH_MONITOR_SHEET"]),
                    text_item("- Slack request commands: /gmp-access, /gmp-request-access, /gmp-issue"),
                ],
            ),
            (
                "Important Notes:",
                [
                    text_item("- Passwords are stored only in 1Password; never in the tracker."),
                    text_item("- Presence Status shows whether the account exists; Activity Status shows whether use is evidenced."),
                    text_item("- Account categories must remain grouped as Katana, Wasp, Shopify, ShipStation, 1Password, and Other Systems."),
                ],
            ),
        ]
    )
    builder.blank()

    builder.heading("1. Purpose")
    builder.para("To define the operating process for receiving, approving, provisioning, verifying, reviewing, and closing GMP-related system access in the Security & GMP Connection Tracker.")
    builder.blank()

    builder.heading("2. Scope")
    builder.para("This SOP applies to IT, managers, and QA personnel responsible for maintaining access records, review checks, and incidents for GMP-connected systems managed by MyMagicHealer.")
    builder.para("It covers the current workbook structure: Access Control, Review Checks, and Incidents, plus the supporting Slack intake and tracker automation.")
    builder.blank()

    builder.heading("3. Responsibilities")
    builder.table(
        [
            ["Role", "Responsibility"],
            ["IT Administrator", "Monitors Slack intake, reviews access requests, maintains the account register, and triages incidents."],
            ["Provisioning Owner", "Creates, changes, or removes access in the target system and records the real account identifier in the tracker."],
            ["Manager / System Owner", "Approves or denies business access where required and supports review decisions for unusual or privileged access."],
            ["IT Manager", "Owns escalations, approves workflow changes, and ensures the tracker remains operational."],
            ["QA Manager", "Approves SOP revisions and confirms the process remains aligned with GMP expectations."],
        ],
        widths=[2400, 6960],
    )
    builder.blank()

    builder.heading("4. Definitions")
    builder.table(
        [
            ["Term", "Definition"],
            ["Access Control", "Primary operating tab containing ACCESS REQUESTS and ACTIVE GMP ACCOUNTS."],
            ["Review Checks", "Action queue used to review due, missing, unmanaged, or no-evidence accounts."],
            ["Incidents", "Tab used to record and manage system issues and access-related incidents."],
            ["Presence Status", "Whether the account currently exists in the system: Provisioning, Present, Missing, Unmanaged, or Revoked."],
            ["Activity Status", "Whether there is real usage evidence: Verified Active, Some Evidence, No Evidence, or Unknown."],
            ["120-Day Review Cycle", "Standard review cadence for non-revoked GMP accounts unless an earlier exception review is required."],
            ["GMP Security menu", "Custom spreadsheet menu used to set up the workbook, install triggers, sync approved requests, refresh evidence, and test Slack."],
        ],
        widths=[2500, 6860],
    )
    builder.blank()

    builder.heading("5. Procedure")
    builder.subheading("5.1  Receive and Review the Request")
    builder.bullet("Open the Access Control tab and review the new request row in ACCESS REQUESTS.")
    builder.bullet("Confirm the target user, company email, department, system, access level, reason, manager, and GMP impact are complete.")
    builder.bullet("Enter the Provisioning Owner who will complete the access setup if the request is approved.")
    builder.bullet("If the request is unusual, privileged, or unclear, confirm approval with the manager or system owner before moving forward.")
    builder.blank()

    builder.subheading("5.2  Approve or Deny the Request")
    builder.bullet("Set Approval to Approved or Denied.")
    builder.bullet("If denied, document the reason in Notes and stop processing.")
    builder.bullet("If approved, confirm the request can move to provisioning.")
    builder.blank()

    builder.subheading("5.3  Provision Access and Sync the Active Account")
    builder.bullet("Create or update the account in the target system.")
    builder.bullet("Record the real account identifier in Access ID. Do not store passwords in the tracker.")
    builder.bullet("Set Provisioning to Provisioned when the account is ready for use.")
    builder.bullet("Run GMP Security -> Sync Approved Requests.")
    builder.bullet("Confirm the related account appears in ACTIVE GMP ACCOUNTS under the correct category: Katana, Wasp, Shopify, ShipStation, 1Password, or Other Systems.")
    builder.blank()

    builder.subheading("5.4  Maintain ACTIVE GMP ACCOUNTS")
    builder.bullet("Use Presence Status to track whether the account currently exists in the system: Provisioning, Present, Missing, Unmanaged, or Revoked.")
    builder.bullet("Use Activity Status only when there is real usage evidence. If no reliable activity source exists, leave Activity Status as Unknown.")
    builder.bullet("Keep Platform Account, Access Level, Dept, Next Review Due, Owner, and Notes current.")
    builder.bullet("Do not manually drag rows between category headers. Re-run Sync Approved Requests if placement looks wrong.")
    builder.blank()

    builder.subheading("5.5  Refresh Evidence and Run Daily Maintenance")
    builder.table(
        [
            ["Menu Action", "Purpose"],
            ["Refresh Katana Accounts", "Pull Katana account presence data and update account rows."],
            ["Refresh Activity Signals", "Recalculate generic activity bands and review indicators from current evidence."],
            ["Refresh Shopify Activity", "Update Shopify-specific activity evidence and then refresh broader activity signals."],
            ["Generate Review Checks", "Create or update review rows based on active-account risk and staleness."],
            ["Run Daily Maintenance", "Run the normal daily sequence for account refresh and review generation."],
        ],
        widths=[2600, 6760],
    )
    builder.blank()

    builder.subheading("5.6  Complete Review Checks")
    builder.bullet("Work only from OPEN REVIEW QUEUE. Completed items remain below in COMPLETED REVIEWS for history.")
    builder.bullet("Read Why In Review, Presence Status, Activity Status, System, Person, and Access Level before making a decision.")
    builder.bullet("Select Reviewer Action: Keep, Reduce, Remove, or Need Info.")
    builder.bullet("Set Decision Status to Done when complete, or Waiting when follow-up is required.")
    builder.blank()

    builder.subheading("5.7  Log and Manage Incidents")
    builder.bullet("Receive issues through /gmp-issue or record them manually in Incidents if Slack is unavailable.")
    builder.bullet("Confirm Severity, System, Reported By, Summary, Assigned To, and Status are complete.")
    builder.bullet("Use Status values Open, Investigating, and Resolved.")
    builder.bullet("Keep investigation notes and current fix status in the same incident row and Slack thread.")
    builder.bullet("If an incident affects an access record, link the related Access ID or Request ID.")
    builder.blank()

    builder.subheading("5.8  Verify Triggers and Integrations")
    builder.bullet("Use GMP Security -> Setup Triggers if triggers are missing or were deleted.")
    builder.bullet("Confirm these triggers exist: onEditInstallable, runDailySecurityMaintenance, keepSecurityCommandWarm_, and processPendingSecurityIntake_.")
    builder.bullet("Run GMP Security -> Test Slack Connection after webhook or bot changes.")
    builder.bullet("In Script Properties, confirm SLACK_WEBHOOK_URL, SLACK_BOT_TOKEN, and HEALTH_MONITOR_URL are populated when troubleshooting integrations.")
    builder.blank()

    builder.heading("6. References")
    builder.link_para("Security & GMP Connection Tracker spreadsheet", URLS["SECURITY_TRACKER_SHEET"])
    builder.link_para("Security Tracker Apps Script", URLS["SECURITY_TRACKER_SCRIPT"])
    builder.link_para("Systems Health Monitor spreadsheet", URLS["HEALTH_MONITOR_SHEET"])
    builder.link_para("IT-008 - IT Security & Access Control", URLS["IT008_DOC"])
    builder.link_para("IT-010 - Cyber Security Incident Response", URLS["IT010_DOC"])
    builder.link_para("FDA 21 CFR Part 11 - Electronic Records; Electronic Signatures", "https://www.ecfr.gov/current/title-21/chapter-I/subchapter-A/part-11")
    builder.blank()

    builder.heading("7. Attachments")
    builder.table(
        [
            ["Attachment", "Summary"],
            ["Attachment A", "ACCESS REQUESTS key fields."],
            ["Attachment B", "ACTIVE GMP ACCOUNTS key fields."],
            ["Attachment C", "REVIEW CHECKS key fields."],
            ["Attachment D", "INCIDENTS key fields."],
        ],
        widths=[2000, 7360],
    )
    builder.blank()
    builder.subheading("Attachment A - ACCESS REQUESTS Key Fields")
    builder.table(
        [
            ["Field", "Description / Values"],
            ["Request ID", "Stable request identifier generated by the workbook."],
            ["Queue Priority", "High / Medium / Low."],
            ["Request Type", "New Access / Change Access / Remove Access."],
            ["Target User", "Person whose account is being created, changed, or removed."],
            ["Company Email", "Primary matching key for fallback sync logic."],
            ["GMP System", "System name used for sync and category placement."],
            ["Access Level", "Admin / Full Access / Read-Write / Read-Only / Limited."],
            ["Approval", "Submitted / Approved / Denied."],
            ["Provisioning", "Open / Provisioning / Provisioned / Removal Pending / Closed."],
            ["Access ID", "Populated after sync or provisioning when an account record exists."],
            ["Slack Thread / Notes", "Links the request to Slack follow-up and audit comments."],
        ],
        widths=[2500, 6860],
    )
    builder.blank()
    builder.subheading("Attachment B - ACTIVE GMP ACCOUNTS Key Fields")
    builder.table(
        [
            ["Field", "Description / Values"],
            ["Access ID", "Primary account identifier in the tracker."],
            ["GMP System", "Canonical category-driving system value."],
            ["Person / Company Email", "Human identity attached to the account row."],
            ["Platform Account", "Actual account or login identifier when known."],
            ["Access Level", "Current granted access level."],
            ["Presence Status", "Provisioning / Present / Missing / Unmanaged / Revoked."],
            ["Activity Status", "Verified Active / Some Evidence / No Evidence / Unknown."],
            ["Last Activity Evidence", "Latest login or usage evidence used for account review."],
            ["Next Review Due", "Date driving review queue generation."],
            ["Review Status", "Open / Waiting / Done."],
            ["Owner", "Internal owner responsible for maintaining the access."],
            ["Notes", "Audit trail of sync, provisioning, and review actions."],
        ],
        widths=[2500, 6860],
    )
    builder.blank()

    builder.subheading("Attachment C - REVIEW CHECKS Key Fields")
    builder.table(
        [
            ["Field", "Description / Values"],
            ["Review ID", "Stable review identifier."],
            ["Review Cycle / Priority", "Scheduling and priority controls."],
            ["Person / System / Access Level", "The account under review."],
            ["Presence Status / Activity Status", "Signals used to decide if access is still justified."],
            ["Why In Review", "Reason the row entered the queue."],
            ["Reviewer Action", "Keep / Reduce / Remove / Need Info."],
            ["Decision Status", "Open / Waiting / Done."],
            ["Completed At / Next Review Due", "Review completion and forward scheduling."],
            ["Linked Request / Slack Thread", "Cross-reference to access changes and communication."],
        ],
        widths=[2800, 6560],
    )
    builder.blank()

    builder.subheading("Attachment D - INCIDENTS Key Fields")
    builder.table(
        [
            ["Field", "Description / Values"],
            ["Incident ID", "Stable incident identifier."],
            ["Reported At / Severity", "Event date and High / Medium / Low impact level."],
            ["System / Summary", "Affected system and concise description of the issue."],
            ["Assigned To / Status", "Current owner and Open / Investigating / Resolved state."],
            ["Response Target", "Internal escalation or response destination."],
            ["Resolution / Root Cause", "Documented fix and underlying cause."],
            ["Linked Access ID / Linked Request ID", "Connection to related access-control records."],
            ["Slack Thread / Notes", "Conversation trace and supporting audit notes."],
        ],
        widths=[2800, 6560],
    )
    builder.blank()

    builder.heading("8. Revision History")
    builder.table(
        [
            ["Version", "Effective Date", "Description of Change", "Change Control #"],
            ["1.5", "March 11, 2026", "Rewrote the local Word version, corrected links, and aligned the content to the live 3-tab tracker, slash commands, categories, and current trigger set.", "Pending"],
            ["1.4", "March 6, 2026", "Removed Google Apps Script function detail and tightened the procedure flow for operations use.", "Pending"],
            ["1.3", "March 6, 2026", "Added embedded tracker screenshots to the formatted SOP attachments.", "Pending"],
            ["1.2", "March 6, 2026", "Reformatted the SOP to the corporate template and aligned the document to the 3-tab workbook model.", "Pending"],
            ["1.0", "January 30, 2026", "Initial release based on the earlier tracker layout.", "n/a"],
        ],
        widths=[1100, 1700, 4300, 2260],
    )


def rewrite_docx(template_path, output_path, build_fn):
    with zipfile.ZipFile(template_path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    sect_pr = extract_sect_pr(entries["word/document.xml"])
    rels = RelationshipManager(entries["word/_rels/document.xml.rels"])
    builder = DocBuilder(sect_pr, rels)
    build_fn(builder)

    entries["word/document.xml"] = builder.document_xml()
    entries["word/_rels/document.xml.rels"] = rels.to_xml()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


def main():
    base = Path(__file__).resolve().parents[2] / "sops" / "working-docx"
    auto_src = base / "2026_SOP-009-Automation-Maintenance-v1.4.docx"
    tracker_src = base / "2026_SOP-009-IT-Security-Tracker-Maintenance-v1.4.docx"
    auto_out = base / "2026_SOP-009-Automation-Maintenance-v1.5.docx"
    tracker_out = base / "2026_SOP-009-IT-Security-Tracker-Maintenance-v1.5.docx"

    rewrite_docx(auto_src, auto_out, build_it014)
    rewrite_docx(tracker_src, tracker_out, build_it015)

    print(auto_out)
    print(tracker_out)


if __name__ == "__main__":
    main()
