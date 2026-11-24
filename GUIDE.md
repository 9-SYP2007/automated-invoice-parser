# Automated Invoice & Receipt Parser for Accounting — User & Deployment Guide

## 1. System Architecture & IT Workflow

This project is a **100% browser-based, serverless AI accounting automation system** built with HTML5, CSS3, and Vanilla JavaScript. Because all file encoding, AI inference requests, mathematical auditing, and Excel/JSON generation occur directly inside the user's browser, there is zero server maintenance cost and strong data privacy.

### End-to-End 5-Stage IT Pipeline

```text
[1. Multi-Format Upload]  ──▶  [2. Multimodal AI Extraction]  ──▶  [3. Math & 7% VAT Audit]
   (.png, .jpg, .pdf)           (Gemini 2.5 Flash REST API)         (Line-Item & Tax Check)
                                                                              │
[5. Accounting Export]    ◀──  [4. Interactive Inspector & Editor] ◀──────────┘
   (.xlsx & UTF-8 .json)        (Side-by-Side Preview & Inline Edit)
