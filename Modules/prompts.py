
# def get_executive_summary_and_objective_prompt(reference_text, condensed_rfp, num_interfaces=None):
#     """
#     Generates Crave-style Executive Summary and Objective based on the RFP.
#     """
#     interface_info = (
#         f"The ~{num_interfaces} interfaces represent a scope of approximately {num_interfaces} ICOs to migrate."
#         if num_interfaces
#         else "The project involves migration from the current SAP PI/PO integration platform to SAP Integration Suite."
#     )

#     return f"""
# You are an expert SAP RFP proposal writer for **Crave InfoTech**.

# Your writing must strictly follow **the tone, structure, and style of the reference text below**.
# Do NOT use generic openings like "honored" or "delighted".
# ALWAYS start with:  
# **"Crave InfoTech is pleased to submit proposal for the PI/PO Integration Migration..."**

# ---

# ### 🔹 TASK:
# Generate **two labeled sections** only:
# 1️⃣ **Executive Summary** – 300–350 words.  
#    Follow Crave’s tone exactly:
#    - Start with “Crave InfoTech is pleased to submit proposal for…”
#    - Mention Crave’s SAP partnership (since 2007), global presence, and expertise.
#    - Include bullet points for key SAP competencies (use the same list from the reference if not provided in RFP).
#    - Mention ISO 9000 quality assurance.
#    - Maintain formal, client-centric language.
#    - Keep paragraphs structured and professional (avoid sales tone).

# 2️⃣ **Objective** – a short 100-word paragraph + the below table:

# | No. | Migration of ICOs from SAP PI/PO to SAP Integration Suite as per details below |
# |------|--------------------------------------------------------------------------------|
# | 1 | No of Interfaces to be migrated from SAP PI/PO to SAP Integration Suite: {num_interfaces} |

# Then add:
# **Interfaces Configuration Objects (ICOs) are listed in the Appendix.**

# ---

# ### 🔹 PROJECT CONTEXT:
# {interface_info}

# ### 🔹 USE THIS STYLE AND TONE AS STRICT REFERENCE:
# {reference_text}

# ### 🔹 CONDENSED RFP CONTENT:
# {condensed_rfp:6000}
# """



# # def get_scope_prereq_assumptions_prompt(reference_text, condensed_rfp, num_interfaces=None):
# #     """
# #     Returns a compact yet adaptive prompt for generating:
# #     In Scope, Migration Project Prerequisites, Assumptions, and Out of Scope.
# #     """
# #     interface_info = (
# #         f"Approximately {num_interfaces} interfaces are expected to be migrated."
# #         if num_interfaces
# #         else "The project involves migration of interfaces from SAP PI/PO to SAP Integration Suite."
# #     )

# #     return f"""
# # You are an expert SAP proposal writer at Crave InfoTech.

# # You will use:
# # 1. Crave’s internal reference material (past proposals and templates).
# # 2. Condensed RFP content (specific to the uploaded client).

# # Your task:
# # Generate a professional, **client-specific** section containing:
# # - In Scope
# # - Migration Project Prerequisites
# # - Assumptions
# # - Out of Scope

# # Guidelines:
# # • Analyze the uploaded RFP text to infer the client name (e.g., Procter & Gamble) and use it naturally in sentences.  
# # • If the client name is not explicitly stated, use “the client” instead of any specific past name like “Haceb.”  
# # • Rewrite and adapt any content derived from references — do not copy it literally.  
# # • Mention {interface_info} appropriately within scope.  
# # • Use professional, concise bullets (e.g., “•”).
# # • Maintain the tone and clarity of enterprise-grade SAP migration proposals.
# # • Keep output within 350–450 words total.
# # • Output headings in this exact format (Markdown-style, no extra numbering):
# #   ### In Scope
# #   ### Migration Project Prerequisites
# #   ### Assumptions
# #   ### Out of Scope

# # Reference Material:
# # {reference_text}

# # Condensed RFP:
# # {condensed_rfp}
# # """

# def get_scope_prereq_assumptions_prompt(reference_text, condensed_rfp, num_interfaces=None):
#     """Compact, focused prompt to generate a concise 'Scope and Out of Scope' section."""
#     interface_info = (
#         f"Migration of approximately {num_interfaces} interfaces from SAP PI/PO to SAP Integration Suite."
#         if num_interfaces
#         else "Migration of interfaces from SAP PI/PO to SAP Integration Suite."
#     )

#     return f"""
# You are an SAP proposal writer at Crave InfoTech.

# Generate a concise, professional section covering:
# - In Scope
# - Migration Project Prerequisites
# - Assumptions
# - Out of Scope

# Tone and structure should match a real client proposal — crisp, business-like, and easy to read.  
# Each section should have 4–8 bullet points, each one line or two at most.  
# Mention {interface_info} in the scope.  
# Use 'the client' instead of any past customer name.  
# Avoid unnecessary descriptions or closing summaries.

# Reference Text:
# {reference_text}

# Condensed RFP:
# {condensed_rfp:6000}
# """


# # def get_resource_schedule_and_commercial_prompt(reference_text, condensed_rfp):
# #     """
# #     Prompt for generating the Resource Schedule and Commercials section.
# #     """
# #     return f"""
# # You are a senior SAP proposal writer at Crave InfoTech.

# # You will generate a concise yet professional section titled **Resource Schedule and Commercials**.
# # This should describe:
# # - Resource planning and role distribution across project phases
# # - High-level effort estimation (without pricing)
# # - Commercial and delivery model (e.g., Fixed Bid, T&M)
# # - Key value propositions around efficiency and transparency

# # Tone: enterprise-grade, suitable for RFP responses for global clients.

# # Output format:
# # ### Resource Schedule
# # • Brief introduction (2–3 lines)
# # • Table (with these columns):
# #   | Crave Infotech Resources | Location | Allocation | Resource Count |
# #   |-------|--------|-----------|----------------|
# #   | Project Manager | Onshore | Fulltime | 1 |
# #   | Integration Developer | Onshore | Fulltime | 1 |
# #   | Integration Developer  | Offshore | Fulltime | Based on ICO count |
# #   | Business Analyst  | Offshore | Fulltime | 1 |

# # • Closing summary paragraph highlighting Crave InfoTech’s flexibility and scalability in delivery.

# # Reference Material:
# # {reference_text}

# # Condensed RFP:
# # {condensed_rfp}
# # """

# def get_resource_schedule_and_commercial_prompt(reference_text, condensed_rfp):
#     """
#     Concise prompt for generating the highly structured Resource Schedule and Commercials section.
#     """
#     return f"""
# You are a senior SAP proposal writer at Crave InfoTech.

# Generate the section titled **Resource Schedule and Commercials**. The output MUST strictly follow this exact structure using Markdown.

# **STRUCTURE REQUIRED (Strict Replication):**

# ### Resource Schedule

# 1.  **Crave Resources:**
#     * Start with the intro sentence: "Crave Infotech proposes to deploy the following team with their indicative loading based on current understanding –"
#     * Follow with a 4-column Markdown table (`Crave Infotech Resources`, `Location`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager (Onshore, 1), Integration Developer (Onshore, 1), Integration Developer (Offshore, 3), Business Analyst (Offshore, 1).
# 2.  **Client Resources:**
#     * Start with the intro sentence: "Recommended team from [Client Name, use 'the client'] who need to be available during the project execution –"
#     * Follow with a 3-column Markdown table (`Client Resources`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager, SAP IT/Basis, Solution Architect, Integration Specialist, Business Analysts/Functional SME, ABAP (If required).

# ### Commercials

# 3.  **Model & Cost:**
#     * Start with the paragraph: "We propose to execute this project on T&M basis. Following is the resource estimation and indicative of total cost:"
#     * Include the fixed line: "Cost: $ (17 Weeks)"
# 4.  **Change Request:** Include the bolded paragraph: "**Any new enhancements or changes identified during the project phase will be considered a change request and will be estimated separately**"
# 5.  **Notes:** Use header `Note:` followed by a bulleted list of the two specified points (resource/fee estimates and onsite billing details).
# 6.  **Payment Terms:** Use header `Timesheet, Invoices and Payment Terms` followed by a bulleted list of the four specified payment/invoicing terms.

# Reference Material:
# {reference_text}

# Condensed RFP:
# {condensed_rfp:5000}
# """

# def get_communication_plan_prompt(reference_text, condensed_rfp):
#     """
#     Generates the Communication Plan section with structured subsections:
#     - Communication overview
#     - Issue management and escalation procedures
#     - Exhibits (tables) for interaction, issue classification, and escalation
#     """
#     return f"""
# You are an expert SAP proposal writer at Crave InfoTech.

# You will create a structured section titled **Communication Plan** suitable for enterprise RFP responses.
# This section describes how Crave InfoTech manages communication, meetings, and issue resolution during SAP migration projects.

# The section must include:
# 1. **Introduction paragraph** — briefly describe the communication objectives and reporting cadence.
# 2. **Table: Exhibit – Daily Interaction**
#    Columns: Activity | Communication Mode | Report Recipient/s | Frequency | Comments
#    Include entries for weekly status reports and project effort tracking.
# 3. **Subsection: Issue Resolution and Escalation Procedure**
#    • Paragraph describing Crave InfoTech’s issue management and escalation approach.  
#    • Table: Issue Management  
#      Columns: Task | Timescale | Responsibility  
#    • Bulleted list of issue reporting guidelines.
# 4. **Table: Issue Classification**
#    Columns: Problem Type | Definition | Reporting Process | Solution Responsible
#    Include rows for Low, Serious, and Critical issues.
# 5. **Table: Escalation Process**
#    Columns: Issue Type | Escalation Point | Escalation Criteria | Governance Role (Project Core Group)

# Formatting rules:
# - Use markdown-style headers for subsections (### Header)
# - Preserve table structure exactly
# - Keep tone formal, structured, and professional
# - Use “the client” when specific name not mentioned
# - Keep length within ~700–900 words

# Reference Material:
# {reference_text}

# Condensed RFP:
# {condensed_rfp:6000}
# """

def get_communication_plan_prompt(reference_text, condensed_rfp):
    """
    Generates the Communication Plan section with structured subsections
    that strictly follow the client's required format and content for governance.
    """
    return f"""
You are an expert SAP proposal writer at Crave InfoTech.

You will create a structured section titled **Communication Plan**. The output MUST strictly follow the content and tables outlined below.

**STRICT STRUCTURE AND CONTENT REQUIRED:**

### Communication Plan

1.  **Communication Overview:** Start with the sentence: "Following table represents different reports that will be generated periodically and circulated to various stakeholders."

2.  **Table 1: Exhibit: Daily Interaction:**
    * **Content:** 5-column table (`Activity`, `Communication Mode`, `Report Recipient/s`, `Frequency`, `Comments`) covering **Weekly Status Reports** and **Project Effort and Schedule Variance**. Include Crave PM and [Client Name] PM as recipients.
    * **Closing:** Conclude with the fixed sentence: "Additional meetings will be decided during project planning phase as necessary."

### Issue Resolution and Escalation Procedure

3.  **Introductory Paragraphs:**
    * Paragraph 1: Describe the Crave InfoTech Project Manager's role in identifying, capturing, resolving, tracking, monitoring, and deciding on appropriate escalations.
    * Paragraph 2: State: "The responsibilities / timescales for executing these procedures are:"

4.  **Table 2: Issue Management:**
    * **Content:** 3-column table (`Task`, `Timescale`, `Responsibility`) detailing the 3 core tasks: **Raise Issue Report**, **Assessment/Management**, and **Reporting**. Ensure Crave PM is responsible for the latter two.

5.  **Issue Reporting Guidelines:**
    * **Content:** Generate the complete 7-point bulleted list detailing the full issue reporting lifecycle (from issue report development, documentation, classification, action definition, assignment, monitoring, and closure by the Crave PM).

6.  **Classification Intro:** Add the line: "Following exhibit explains the process of issue classification:"

7.  **Table 3: Issue Classification:**
    * **Content:** 4-column table (`Problem Type`, `Definition`, `Reporting process`, `Solution Responsible`) classifying issues into **Low, Serious, and Critical**. Ensure the process and responsibility reflect escalation to the Client PM for Serious/Critical issues.

8.  **Escalation Intro:** Add the line: "Following escalation process will be followed as part of the issue management:"

9.  **Table 4: Exhibit: Escalation Process:**
    * **Content:** 4-column table detailing two escalation levels: **Project Manager** (for delivery/quality issues, 48hr criteria, Weekly checkpoints) and **Account Executive** (for contract/program management issues, 48hr criteria, Monthly Review).

10. **Closing Sentence:** Add the final line: "The Project Workgroup will be comprised of representatives from Crave InfoTech and [Client Name/Haceb] key project stakeholder."

Formatting rules:
- Use markdown headers (# or ##) only as shown in the output structure.
- Use "Haceb" or "the client" based on context.
- Keep tone formal, structured, and professional.

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""


def get_executive_summary_and_objective_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """
    Generates Crave-style Executive Summary and Objective based on the RFP.
    """
    interface_info = (
        f"The ~{num_interfaces} interfaces represent a scope of approximately {num_interfaces} ICOs to migrate."
        if num_interfaces
        else "The project involves migration from the current SAP PI/PO integration platform to SAP Integration Suite."
    )

    return f"""
You are an expert SAP RFP proposal writer for **Crave InfoTech**.

Your writing must strictly follow **the tone, structure, and style of the reference text below**.
Do NOT use generic openings like "honored" or "delighted".
ALWAYS start with:  
**"Crave InfoTech is pleased to submit proposal for the PI/PO Integration Migration..."**

---

### 🔹 TASK:
Generate **two labeled sections** only:
1️⃣ **Executive Summary** – 300–350 words.  
   Follow Crave’s tone exactly:
   - Start with “Crave InfoTech is pleased to submit proposal for…”
   - Mention Crave’s SAP partnership (since 2007), global presence, and expertise.
   - Include bullet points for key SAP competencies (use the same list from the reference if not provided in RFP).
   - Mention ISO 9000 quality assurance.
   - Maintain formal, client-centric language.
   - Keep paragraphs structured and professional (avoid sales tone).

2️⃣ **Objective** – a short 100-word paragraph + the below table:

| No. | Migration of ICOs from SAP PI/PO to SAP Integration Suite as per details below |
|------|--------------------------------------------------------------------------------|
| 1 | No of Interfaces to be migrated from SAP PI/PO to SAP Integration Suite: {num_interfaces} |

Then add:
**Interfaces Configuration Objects (ICOs) are listed in the Appendix.**

---

### 🔹 PROJECT CONTEXT:
{interface_info}

### 🔹 USE THIS STYLE AND TONE AS STRICT REFERENCE:
{reference_text}

### 🔹 CONDENSED RFP CONTENT:
{condensed_rfp}
"""

def get_scope_prereq_assumptions_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """Compact, focused prompt to generate a concise 'Scope and Out of Scope' section."""
    interface_info = (
        f"Migration of approximately {num_interfaces} interfaces from SAP PI/PO to SAP Integration Suite."
        if num_interfaces
        else "Migration of interfaces from SAP PI/PO to SAP Integration Suite."
    )

    return f"""
You are an SAP proposal writer at Crave InfoTech.

Generate a concise, professional section covering:
- In Scope
- Migration Project Prerequisites
- Assumptions
- Out of Scope

Tone and structure should match a real client proposal — crisp, business-like, and easy to read.  
Each section should have 4–8 bullet points, each one line or two at most.  
Mention {interface_info} in the scope.  
Use 'the client' instead of any past customer name.  
Avoid unnecessary descriptions or closing summaries.

Reference Text:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""

def get_resource_schedule_and_commercial_prompt(reference_text, condensed_rfp):
    """
    Concise prompt for generating the highly structured Resource Schedule and Commercials section.
    """
    return f"""
You are a senior SAP proposal writer at Crave InfoTech.

Generate the section titled **Resource Schedule and Commercials**. The output MUST strictly follow this exact structure using Markdown.

**STRUCTURE REQUIRED (Strict Replication):**

### Resource Schedule

    * Start with the intro sentence: "Crave Infotech proposes to deploy the following team with their indicative loading based on current understanding –"
    * Follow with a 4-column Markdown table (`Crave Infotech Resources`, `Location`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager (Onshore, Fulltime), Integration Developer (Onshore, Fulltime), Integration Developer (Offshore, Fulltime), Business Analyst (Offshore, Fulltime).

    * Start with the intro sentence: "Recommended team from [Client Name, use 'the client'] who need to be available during the project execution –"
    * Follow with a 3-column Markdown table (`Client Resources`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager, SAP IT/Basis, Solution Architect, Integration Specialist, Business Analysts/Functional SME, ABAP (If required).

### Commercials

    * Start with the paragraph: "We propose to execute this project on T&M basis. Following is the resource estimation and indicative of total cost:"
    * Include the fixed line: "Cost: $ (17 Weeks)"
4.  **Change Request:** Include the bolded paragraph: "**Any new enhancements or changes identified during the project phase will be considered a change request and will be estimated separately**"
5.  **Notes:** Use header `Note:` followed by a bulleted list of the two specified points (resource/fee estimates and onsite billing details).
6.  **Payment Terms:** Use header `Timesheet, Invoices and Payment Terms` followed by a bulleted list of the four specified payment/invoicing terms.

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""


def get_communication_plan_prompt(reference_text, condensed_rfp):
    """
    Generates a concise, structured Communication Plan section prompt with Crave/client role clarity.
    """
    return f"""
You are an expert SAP proposal writer at Crave InfoTech.

Write a formal, client-ready **Communication Plan** section for an SAP migration or implementation RFP.  
Describe how Crave InfoTech and the client (use actual client name from context if available, e.g., Haceb) will manage communication, meetings, reporting, and escalation during the project.

The section must include:

1. 2–3 lines on the importance of clear and consistent communication for project success, stakeholder alignment, and timely decision-making.

2. **Exhibit: Daily Interaction** – Table with columns:  
   Activity | Communication Mode | Report Recipient/s | Frequency | Comments  
   Include a row for Weekly Status Report and others like Kick-off Meeting, Daily Stand-up, Steering Committee, etc.  
   Clearly mention which roles belong to **Crave InfoTech** and which belong to **the client** (e.g., “Crave InfoTech Project Manager”, “Haceb Project Manager”).

3. **Issue Resolution and Escalation Procedure** – Short paragraph describing the structured approach for logging, managing, and escalating issues.  
   Then include:
   - **Table: Issue Management** (Task | Timescale | Responsibility)  
   - **Bulleted list:** Issue reporting guidelines that mention role responsibilities on both Crave and client sides.

4. **Table: Issue Classification** – Columns: Problem Type | Definition | Reporting Process | Solution Responsible  
   Include rows for Low, Serious, and Critical.  
   Indicate clearly who is responsible on Crave side and on the client side for each case.

5. **Table: Escalation Process** – Columns: Issue Type | Escalation Point | Escalation Criteria | Governance Role (Project Core Group)  
   Include rows for Project Delivery, Contract, Unresolved Delivery Issue, and Program Management Issue.  
   Specify which escalation points or governance roles belong to Crave vs. the client.

6. Close with 1–2 lines summarizing how this structured plan ensures transparency, timely updates, and strong collaboration between Crave InfoTech and the client.

**Formatting rules:**
- Use markdown headers and tables exactly.
- Keep tone formal, enterprise-level, and realistic.
- Replace “the client” with actual client name (from context or reference) wherever possible.
- Keep total length around 700–900 words.
- Ensure all roles are clearly marked as Crave-side or Client-side.

### 🔹 USE THIS STYLE AND TONE AS STRICT REFERENCE:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""

