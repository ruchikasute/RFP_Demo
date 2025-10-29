
def get_executive_summary_and_objective_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """
    Unified prompt that generates:
    - Executive Summary (narrative)
    - Objective (structured, with table)
    """

    interface_info = (
        f"The project involves the migration of approximately {num_interfaces} interfaces from SAP PI/PO to SAP Integration Suite."
        if num_interfaces
        else "The project involves migration from the current SAP PI/PO integration platform to SAP Integration Suite."
    )

    return f"""
You are an expert SAP RFP proposal writer for Crave InfoTech.

You will receive:
1. **Reference Material** – previous successful RFP responses from Crave InfoTech (tone, structure, writing style).
2. **Condensed RFP** – summary of the current client's requirements.
3. **Project Context** – {interface_info}

Your task:
Generate **two distinct sections** clearly labeled as:
1️⃣ **Executive Summary** – a persuasive, client-centric narrative (250–350 words) about Crave InfoTech’s capability, value proposition, and approach to enable the client's integration modernization.
2️⃣ **Objective** – a concise section (around 80–100 words) describing the migration goal, followed by a **table** exactly in the following structure:

| No. | Migration of ICOs from SAP PI/PO to SAP Integration Suite as per details below |
|------|--------------------------------------------------------------------------------|
| 1 | No of Interfaces to be migrated from SAP PI/PO to SAP Integration Suite: {num_interfaces} |

Add a closing note under the table:
**Interfaces Configuration Objects (ICOs) are listed in the Appendix.**

Formatting rules:
- Label sections clearly using “**Executive Summary**” and “**Objective**”.
- Use professional, enterprise-grade tone suitable for clients like Procter & Gamble.
- Preserve the table format exactly as shown.
- Avoid adding extra lists or headings beyond those sections.

Reference Material:
{reference_text}

Condensed RFP Content:
{condensed_rfp}
"""

def get_scope_prereq_assumptions_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """
    Returns a compact yet adaptive prompt for generating:
    In Scope, Migration Project Prerequisites, Assumptions, and Out of Scope.
    """
    interface_info = (
        f"Approximately {num_interfaces} interfaces are expected to be migrated."
        if num_interfaces
        else "The project involves migration of interfaces from SAP PI/PO to SAP Integration Suite."
    )

    return f"""
You are an expert SAP proposal writer at Crave InfoTech.

You will use:
1. Crave’s internal reference material (past proposals and templates).
2. Condensed RFP content (specific to the uploaded client).

Your task:
Generate a professional, **client-specific** section containing:
- In Scope
- Migration Project Prerequisites
- Assumptions
- Out of Scope

Guidelines:
• Analyze the uploaded RFP text to infer the client name (e.g., Procter & Gamble) and use it naturally in sentences.  
• If the client name is not explicitly stated, use “the client” instead of any specific past name like “Haceb.”  
• Rewrite and adapt any content derived from references — do not copy it literally.  
• Mention {interface_info} appropriately within scope.  
• Use professional, concise bullets (e.g., “•”).
• Maintain the tone and clarity of enterprise-grade SAP migration proposals.
• Keep output within 350–450 words total.
• Output headings in this exact format (Markdown-style, no extra numbering):
  ### In Scope
  ### Migration Project Prerequisites
  ### Assumptions
  ### Out of Scope

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""

def get_resource_schedule_and_commercial_prompt(reference_text, condensed_rfp):
    """
    Prompt for generating the Resource Schedule and Commercials section.
    """
    return f"""
You are a senior SAP proposal writer at Crave InfoTech.

You will generate a concise yet professional section titled **Resource Schedule and Commercials**.
This should describe:
- Resource planning and role distribution across project phases
- High-level effort estimation (without pricing)
- Commercial and delivery model (e.g., Fixed Bid, T&M)
- Key value propositions around efficiency and transparency

Tone: enterprise-grade, suitable for RFP responses for global clients.

Output format:
### Resource Schedule and Commercials
• Brief introduction (2–3 lines)
• Table (with these columns):
  | Role | Phase | Duration | Allocation (%) |
  |-------|--------|-----------|----------------|
  | Project Manager | Overall | 12 weeks | 15% |
  | Integration Consultant | Build Phase | 10 weeks | 40% |
  | QA & Testing | Testing | 4 weeks | 20% |
  | Cutover Lead | Deployment | 2 weeks | 10% |
  | Support | Hypercare | 2 weeks | 15% |
• Closing summary paragraph highlighting Crave InfoTech’s flexibility and scalability in delivery.

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""
def get_communication_plan_prompt(reference_text, condensed_rfp):
    """
    Generates the Communication Plan section with structured subsections:
    - Communication overview
    - Issue management and escalation procedures
    - Exhibits (tables) for interaction, issue classification, and escalation
    """
    return f"""
You are an expert SAP proposal writer at Crave InfoTech.

You will create a structured section titled **Communication Plan** suitable for enterprise RFP responses.
This section describes how Crave InfoTech manages communication, meetings, and issue resolution during SAP migration projects.

The section must include:
1. **Introduction paragraph** — briefly describe the communication objectives and reporting cadence.
2. **Table: Exhibit – Daily Interaction**
   Columns: Activity | Communication Mode | Report Recipient/s | Frequency | Comments
   Include entries for weekly status reports and project effort tracking.
3. **Subsection: Issue Resolution and Escalation Procedure**
   • Paragraph describing Crave InfoTech’s issue management and escalation approach.  
   • Table: Issue Management  
     Columns: Task | Timescale | Responsibility  
   • Bulleted list of issue reporting guidelines.
4. **Table: Issue Classification**
   Columns: Problem Type | Definition | Reporting Process | Solution Responsible
   Include rows for Low, Serious, and Critical issues.
5. **Table: Escalation Process**
   Columns: Issue Type | Escalation Point | Escalation Criteria | Governance Role (Project Core Group)

Formatting rules:
- Use markdown-style headers for subsections (### Header)
- Preserve table structure exactly
- Keep tone formal, structured, and professional
- Use “the client” when specific name not mentioned
- Keep length within ~700–900 words

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""
