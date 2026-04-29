# 🔍 DFIR Investigation: M57-JO Case

A comprehensive digital forensics and incident response investigation into a corporate data exfiltration incident at M57.biz.

## 🌐 Live Forensic Portal
You can view the interactive investigation dashboard here:
👉 **[Launch Interactive Forensic Portal](https://lafandoor.github.io/DFIR-Investigation-M57/portal/)**

---

## 📋 Case Overview
- **Case Title**: M57.Biz — Corporate Data Exfiltration
- **Case Number**: DFIR-2024-012
- **Evidence Image**: `jo-2009-11-16.E01` (Windows XP SP3)
- **Primary Subject**: Jo (Employee)
- **Verdict**: **Highly Probable Exfiltration**

## 🛠️ Tools & Methodology
- **Autopsy 4.23**: Primary forensic analysis and artifact extraction.
- **FTK Imager**: Image verification and mounting.
- **Velociraptor**: Threat hunting simulation and VQL-based artifact review.
- **Custom Report Generator**: A Node.js script (`java.js`) using the `docx` library to generate formal documentation.

## 📂 Project Structure
- `/portal`: The interactive web dashboard.
- `/csv's`: Exported metadata and artifact logs from Autopsy.
- `/ss`: Screenshot evidence gallery.
- `java.js`: The forensic report generation script.
- `DFIR_FINAL_v2.docx`: The official final investigation report.

## 📑 Key Findings
1. **ZIPPER.EXE**: Evidence of file compression found in unallocated Prefetch space (now deleted).
2. **USB Connection**: A LaCie Rugged HDD was connected just before system imaging.
3. **CD Burning**: IMAPI service logs confirm disc burning activity on the acquisition day.
4. **Counter-Forensics**: Execution of HELIX.EXE followed by an immediate system shutdown.

---
*Prepared by Youssef Moataz — AASTMT Cybersecurity, Class of 2026*
