# Refined Solution Architecture: PPTX → DOCX Design Document Converter

> This document combines and refines ideas from both the initial architecture proposal and Gemini Pro's agentic architecture review, grounded by an actual analysis of the sample PPTX (27 slides) and DOCX (383 paragraphs, 116 sections, 22 tables, ~24 images, 7 appendices).

---

## 1. Problem Statement

Pre-sales engineers create design documents as PowerPoint slides (PPTX) first — these are presented directly to the customer. Afterward, a much more detailed Word document (DOCX) must be produced covering the same topics in formal, standalone prose.

This manual conversion is:
- **Time-consuming** — the DOCX is typically 3–5× the content length of the PPTX.
- **Repetitive** — the same structural patterns appear across projects.
- **Error-prone** — details get lost, images get missed, and governance sections are forgotten.

The goal is an **AI-powered pipeline** that automates this conversion.

---

## 2. Document Analysis: PPTX vs DOCX Differences

We analyzed the real sample pair (`UMS_PowerBI_Fabric_Design_Workshop_v1.1_final.pptx` → `UMS_PowerBI_Fabric_Planning_Design_v1.1.docx`) and identified **7 critical difference categories** that the AI pipeline must handle:

### 2.1 Content Comparison Summary

| Dimension | PPTX (Slides) | DOCX (Document) |
|---|---|---|
| **Volume** | 27 slides | 383 paragraphs, 116 heading sections |
| **Text style** | Bullet points, short statements | Full paragraphs, formal prose |
| **Tables** | ~10 inline tables | 22 tables (same data + new governance tables) |
| **Images** | ~5 diagrams/screenshots | ~24 images (same + new explanatory figures) |
| **Speaker Notes** | Present on some slides (e.g., Slide 7, 14, 19) | N/A — notes are absorbed into body text |
| **Structure depth** | Flat (slide titles only) | 4-level heading hierarchy (H1–H4) |
| **Governance content** | Not present | Extensive (Resource Org, Naming, Tags, Cost Mgmt) |

### 2.2 The 7 Critical Difference Categories

#### Category 1: Boilerplate & Governance Front Matter
**The DOCX contains 15+ pages of front matter that do NOT exist in the PPTX:**
- Cover page (project title, date, company branding)
- Abstract
- Prepared By / Distribution List / Change Record tables
- Terms of Use / Disclaimer / Trademarks sections
- Design Document Sign-Off block (with signature placeholders)
- Table of Contents, Table of Figures, Table of Tables

> **AI Strategy:** These must NOT be AI-generated. They should come from a **pre-built corporate DOCX template** loaded by `python-docx`. The AI only fills in dynamic fields (project name, date, authors).

#### Category 2: Executive Narratives (Not in Slides)
**The DOCX has sections that are synthesized from the holistic slide context:**
- **Executive Summary** — a full paragraph summarizing the entire project scope
- **Project Overview** — scope statement rewritten for a document audience
- **Document Purpose** — explains the context of the workshop/design review
- **Document Audience** — describes who should read this document

> **AI Strategy:** The **Orchestrator Agent** must generate these by reading ALL slide content as context and synthesizing a narrative. This is a "document-level" generation task, not a "slide-level" task.

#### Category 3: Expanding the "Why" (RAG Focus)
**Slides state decisions; the DOCX explains the reasoning and technology:**

| PPTX (Slide 7) | DOCX Section |
|---|---|
| "MFA has been enforced at the tenant level" | Full explanation of Conditional Access, what it is, how it works with Entra ID, why MFA was chosen, purpose bullets, architecture diagrams, and configuration recommendations (~15 paragraphs) |

| PPTX (Slide 9) | DOCX Section |
|---|---|
| Architecture diagram only | Detailed breakdown of each component in the diagram: Data Factory, Bronze/Silver/Gold Lakehouse, Dataflow Gen2, AutoML, Semantic Model, Power BI Desktop connection — each with 2-3 sentences |

> **AI Strategy:** The **Researcher Agent (RAG)** queries Microsoft Learn to expand each technology mentioned in the slides. The **Writer Agent** uses both the slide bullet point AND the RAG context to produce formal prose explaining what, why, and how.

#### Category 4: Injecting Standard Governance Sections
**The DOCX includes governance sections that are only briefly implied in slides (or not present at all):**
- **Resource Organization Design** (Management Groups, Subscriptions, Resource Groups)
- **Naming and Tagging** (Azure naming conventions, Fabric naming conventions, Application acronyms, Suggested tags)
- **Governance Consideration** (Azure Governance Discipline, Cost Management Tools, Violation Triggers and Actions)
- **Multi-Factor Authentication**
- **Identity and Access Management** (RBAC patterns, Identity Mapping)

> **AI Strategy:** The **Orchestrator Agent** must recognize when a project involves Azure/Fabric infrastructure and automatically add these standard sections to the document outline. The **Writer Agent** generates content using RAG from Microsoft Learn (Cloud Adoption Framework, Well-Architected Framework). **Critical guard-rail:** These injected sections should be flagged as "auto-generated" suggestions for the pre-sales engineer to review, preventing hallucinated requirements.

#### Category 5: Image and Diagram Handling
**Both documents share common architecture diagrams, but the DOCX adds more:**
- PPTX diagrams are transferred as-is (binary image copies)
- DOCX adds formal figure captions (e.g., "Figure 1 Data Gateway Architecture Overview")
- DOCX includes additional explanatory screenshots (e.g., "Figure 2 Admin Portal Button")
- Complex diagrams get text descriptions in the DOCX (data flow walkthrough)

> **AI Strategy:** 
> - Transfer images from PPTX as binary blobs → insert into DOCX via `python-docx`
> - Generate figure captions automatically using slide context
> - Use **GPT-4o Vision** for architecture diagrams to generate text-based data flow descriptions
> - Additional explanatory screenshots should be handled via a library of common reference images or flagged as placeholders

#### Category 6: Tables — Preserved + Expanded
**Tables from the PPTX are kept in the DOCX with the same data, but the DOCX adds new tables:**
- PPTX tables (design decisions, workspace roles, naming conventions, security rules, VM specs) → copied directly
- DOCX adds: Fabric Administrator Capabilities, Workspace-level Role Capabilities, Cost Management Tools, Violation Triggers, Outbound Port Summary — these are reference tables from Microsoft Learn

> **AI Strategy:** 
> - Copy tables from PPTX as-is (preserve structure)
> - Use RAG to generate additional reference tables from Microsoft Learn content
> - Apply consistent table formatting via the DOCX template

#### Category 7: Appendices (Reference Links)
**The DOCX ends with 7 appendices that are pure Microsoft Learn link collections:**
- Appendix 1: Computing (intro, limits, VM sizes)
- Appendix 2: Network (VNet, VPN Gateway, limits)
- Appendix 3: Identity & Security (NSG, Security Center, WAF)
- Appendix 4: Logging & Monitoring (Azure Monitor, Log Analytics)
- Appendix 5: Cloud Governance (Blueprint, ISO 27001)
- Appendix 6: On-Premise Data Gateway (firewall, communication settings)
- Appendix 7: Machine Learning (Data Science, tutorials)

> **AI Strategy:** The Orchestrator should detect which Azure services are mentioned and auto-generate relevant appendices with Microsoft Learn links. These can be partially template-driven (common appendix structures) + AI-generated based on detected services.

---

## 3. Refined Architecture: Weighed Comparison & Best Approach

### 3.1 Comparison of Proposed Approaches

| Aspect | My Initial Proposal | Gemini Pro's Proposal | **Refined Recommendation** |
|---|---|---|---|
| **Extraction** | python-pptx + Content Understanding | python-pptx + GPT-4o Vision for diagrams only | **python-pptx as primary** + GPT-4o Vision only for complex architecture diagrams. Content Understanding is overkill for PPTX parsing since python-pptx handles structured extraction well |
| **AI Pattern** | Single GPT-4.1 per slide | Multi-agent (Orchestrator → Researcher → Writer) | **Multi-agent pattern** — essential because the DOCX requires both document-level synthesis (Executive Summary) and slide-level expansion. A single prompt cannot do both |
| **RAG** | Azure AI Search + Microsoft Learn | Azure AI Search + Microsoft Learn | **Agreed** — essential for expanding the "why" and generating governance sections |
| **Image Handling** | Binary transfer + GPT-4o for understanding | Binary transfer + GPT-4o for diagrams only | **Agreed** — binary transfer for all images, GPT-4o Vision only for diagrams that need text descriptions |
| **Template** | python-docx template | python-docx corporate template with boilerplate pre-loaded | **Corporate template approach** — load a DOCX template with all fixed front matter, sign-off blocks, and formatting pre-configured |
| **Slide Grouping** | GPT-4.1 groups related slides | Orchestrator Agent designs ToC | **Orchestrator Agent** — better terminology and clearer responsibility separation |
| **Missing Sections** | Not addressed | Auto-inject governance sections | **Yes, but with guard-rails** — auto-inject but flag as suggestions to prevent hallucinated requirements |
| **Speaker Notes** | Mentioned as "gold" content | Extracted and used as primary context | **Agreed** — speaker notes are the #1 supplementary context source |

### 3.2 Recommended Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    EXTRACTION PHASE                      │
│                                                          │
│  ┌──────────────┐    ┌───────────────┐                  │
│  │  python-pptx │    │  GPT-4o       │                  │
│  │  • Text      │    │  Vision       │                  │
│  │  • Notes     │───▶│  • Diagram    │                  │
│  │  • Images    │    │    desc.      │                  │
│  │  • Tables    │    └───────┬───────┘                  │
│  └──────┬───────┘            │                           │
│         │                    │                           │
│         ▼                    ▼                           │
│  ┌──────────────────────────────────────┐               │
│  │      Structured JSON Payload         │               │
│  │  (per-slide: title, text, notes,     │               │
│  │   images, tables, diagram_desc)      │               │
│  └──────────────────┬───────────────────┘               │
└─────────────────────┼───────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────┐
│              AGENTIC GENERATION PHASE                    │
│              (Azure AI Foundry)                          │
│                                                          │
│  ┌────────────────────────────────────┐                  │
│  │        ORCHESTRATOR AGENT          │                  │
│  │  • Reads full JSON payload         │                  │
│  │  • Groups slides into sections     │                  │
│  │  • Designs DOCX Table of Contents  │                  │
│  │  • Detects missing standard        │                  │
│  │    sections and injects them       │                  │
│  │  • Dispatches to Writer Agent      │                  │
│  └───────────┬────────────────────────┘                  │
│              │                                           │
│   ┌──────────┴──────────┐                                │
│   ▼                     ▼                                │
│  ┌──────────────┐  ┌──────────────────┐                  │
│  │  RESEARCHER  │  │  WRITER AGENT    │                  │
│  │  AGENT (RAG) │  │  (GPT-4.1)      │                  │
│  │              │  │                  │                  │
│  │  Azure AI    │  │  • Expand bullets│                  │
│  │  Search +    │◀▶│  • Write exec    │                  │
│  │  MS Learn    │  │    summaries     │                  │
│  │  Index       │  │  • Generate     │                  │
│  │              │  │    governance   │                  │
│  └──────────────┘  └────────┬─────────┘                  │
│                             │                            │
└─────────────────────────────┼────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────┐
│                  ASSEMBLY PHASE                          │
│                  (python-docx)                           │
│                                                          │
│  ┌──────────────────┐  ┌──────────────────────────┐     │
│  │  Corporate DOCX  │  │  AI-Generated Content    │     │
│  │  Template        │  │  + Original Images       │     │
│  │  • Cover Page    │  │  + Copied Tables         │     │
│  │  • ToU/Disclaimer│  │  + Figure Captions       │     │
│  │  • Sign-Off Block│  │  + Appendix Links        │     │
│  │  • Formatting    │  │                          │     │
│  └────────┬─────────┘  └───────────┬──────────────┘     │
│           │                        │                     │
│           ▼                        ▼                     │
│  ┌─────────────────────────────────────────────────┐    │
│  │              Final DOCX Output                   │    │
│  │  • Corporate template + AI-generated prose       │    │
│  │  • Original images with captions                 │    │
│  │  • Tables (copied + RAG-generated)               │    │
│  │  • Auto-generated appendices                     │    │
│  └─────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────┘
```

---

## 4. Proposed JSON Schema (Extraction Output)

The `python-pptx` extractor should produce a structured JSON payload like this:

```json
{
  "metadata": {
    "filename": "UMS_PowerBI_Fabric_Design_Workshop_v1.1_final.pptx",
    "total_slides": 27,
    "extracted_at": "2025-11-06T10:00:00Z",
    "project_name": "UMS Power BI & Fabric"
  },
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title_slide",
      "layout_name": "Title Slide 3",
      "title": "UMS Power BI & Fabric",
      "text_content": ["Design Review", "6 November 2025"],
      "speaker_notes": null,
      "images": [],
      "tables": [],
      "is_section_divider": true
    },
    {
      "slide_number": 9,
      "slide_type": "content_slide",
      "layout_name": "Title_Content_WHITE_C",
      "title": "Data Workflow Design",
      "text_content": [],
      "speaker_notes": null,
      "images": [
        {
          "image_id": "img_009_001",
          "filename": "slide_9_image_1.png",
          "binary_path": "/tmp/extracted_images/slide_9_image_1.png",
          "is_architecture_diagram": true,
          "ai_description": "Architecture diagram showing data flow from On-Premises SQL Server through OPDG to Microsoft Fabric, with Bronze/Silver/Gold Lakehouse layers and Power BI visualization."
        }
      ],
      "tables": [],
      "is_section_divider": false
    },
    {
      "slide_number": 10,
      "slide_type": "content_slide",
      "layout_name": "Title_Content_WHITE_C",
      "title": "Data Workflow Design Consideration",
      "text_content": [],
      "speaker_notes": null,
      "images": [],
      "tables": [
        {
          "table_id": "tbl_010_001",
          "headers": ["ID", "Description", "Workloads Type"],
          "rows": [
            ["DC01", "Data Pipeline will be used to ingest data from Microsoft SQL Server.", "Data Factory"],
            ["DC02", "...", "..."]
          ]
        }
      ],
      "is_section_divider": false
    }
  ]
}
```

---

## 5. Key Architectural Risks and Mitigations

### Risk 1: Incoherent Narrative Across 40+ Pages
**Problem:** If each section is generated independently, the document may read as disconnected paragraphs rather than a cohesive narrative.

**Mitigation:**
- The Orchestrator Agent generates a **document outline with section summaries** first.
- The Writer Agent receives the outline context with every section request, maintaining awareness of how each section fits into the whole.
- A final **coherence review pass** using GPT-4.1 checks cross-references and narrative flow.

### Risk 2: Hallucinated Requirements in Auto-Injected Sections
**Problem:** The "Injecting Missing Sections" feature (e.g., governance, cost management) could introduce requirements the pre-sales engineer did not intend.

**Mitigation:**
- Auto-injected sections are tagged with `[AI-SUGGESTED]` markers in the generated DOCX.
- The AI uses **only factual, generic best-practice content** from Microsoft Learn — no project-specific claims.
- The pre-sales engineer reviews and approves/removes these sections before delivery.
- The system prompt explicitly instructs: *"Write these sections as general recommendations, not project commitments."*

### Risk 3: Image Quality and Placement
**Problem:** Images extracted from PPTX may lose quality or be placed incorrectly in the DOCX.

**Mitigation:**
- Extract images as high-resolution binary blobs (PNG/EMF) from PPTX.
- Use `python-docx` InlineShape insertion with explicit width/height settings.
- Map each image to its DOCX section using the slide-to-section mapping from the Orchestrator.

### Risk 4: Speaker Notes Inconsistency
**Problem:** Some slides have detailed speaker notes (e.g., Slide 7 has reasoning about Conditional Access), while most have no notes at all.

**Mitigation:**
- When notes exist, they are treated as **primary context** (higher priority than bullet text).
- When notes are absent, the Writer Agent relies on the RAG Researcher to fill in the "why."
- The system prompt adjusts behavior based on whether notes are available.

---

## 6. Tech Stack (Final)

| Component | Service / Tool | Justification |
|---|---|---|
| **Backend API** | Python FastAPI | Lightweight, async-capable, easy to deploy |
| **Slide Parsing** | `python-pptx` | Best structured extraction for PPTX natively |
| **Diagram Understanding** | Azure OpenAI GPT-4o (Vision) | Only for architecture diagrams, not all images |
| **Orchestrator + Writer** | Azure OpenAI GPT-4.1 | Cost-effective, best for long-form text generation |
| **Knowledge Grounding** | Azure AI Search + Microsoft Learn Index | RAG for expanding technology descriptions |
| **Document Assembly** | `python-docx` | Programmatic DOCX creation with template support |
| **Agentic Framework** | Azure AI Foundry Agent Service / Semantic Kernel | Orchestrator→Worker pattern management |
| **Frontend** (optional) | Web UI (upload + preview + review) | For the pre-sales engineer to review before export |
| **Hosting** | Azure Container Apps | Scalable, cost-effective for event-driven workloads |

---

## 7. Feasibility POC Scope

For the initial proof-of-concept, implement a **simplified 3-step pipeline** using the UMS sample documents:

### POC Phase 1: Extraction
- `python-pptx` extracts all 27 slides into structured JSON.
- Images exported as binary files.
- Speaker notes captured.

### POC Phase 2: Generation
- **Single GPT-4.1 call** to generate the document outline (ToC) from the full JSON.
- **Per-section GPT-4.1 calls** to expand slide content into paragraphs.
- **GPT-4o Vision** used for the Data Workflow Architecture diagram (Slide 9) and AutoML Architecture (Slide 21) to generate text descriptions.
- **No RAG in POC** — hardcode a few Microsoft Learn references to demonstrate the pattern.

### POC Phase 3: Assembly
- Load a simple DOCX template with Enfrasys branding.
- Inject AI-generated sections with heading hierarchy.
- Insert original images with figure captions.
- Copy tables from PPTX into DOCX.

### POC Success Criteria
- Generated DOCX should have the same section structure as the sample DOCX.
- Images should appear in correct sections.
- AI-expanded text should be technically accurate and formal in tone.
- Total generation time < 5 minutes for a 27-slide deck.

---

## 8. Beyond POC: Production Roadmap

| Phase | Scope |
|---|---|
| **POC** | Single pipeline, hardcoded RAG, UMS sample only |
| **Phase 1** | Multi-agent with Azure AI Foundry, real RAG with Azure AI Search, corporate template library |
| **Phase 2** | Web UI for upload/review/edit, appendix auto-generation, multiple project types |
| **Phase 3** | Bi-directional (DOCX → PPTX), multi-language support, feedback learning loop |
