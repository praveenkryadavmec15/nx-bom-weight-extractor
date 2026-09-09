# Contributing to Siemens NX BOM & Weight Extractor

Thank you for your interest in improving this project! We welcome contributions, bug reports, feature suggestions, and documentation improvements from the engineering and CAD automation community.

---

## ⚠️ Core Design Principle: Zero External Dependencies

Before proposing code changes or opening pull requests, please keep in mind the fundamental constraint of this project:

> **Corporate IT Lockdowns**: Most enterprise CAD workstations strictly prohibit `pip install` and prevent installing third-party Python packages.
> 
> **Strict Rule**: All code must run purely on the **native Python standard library** bundled with Siemens NX (`os`, `datetime`, `xml.etree.ElementTree`, `re`, `math`, etc.) alongside `NXOpen` and `NXOpen.UF`.
> 
> **Do NOT introduce dependencies** such as `pandas`, `openpyxl`, `xlsxwriter`, `numpy`, or `scipy`.

---

## 🐛 Reporting Bugs

When submitting a bug report, please provide:
1. **NX Version**: (e.g., NX 12, NX 1980, NX 2212, NX 2406).
2. **Environment**: Teamcenter-managed or Native NX mode.
3. **Behavior**: What happened vs. what you expected.
4. **NX Listing Window Output**: Copy and paste any logs or error traces printed in the NX Information Window.
5. **Sample Part/Assembly Structure**: (Without sharing any proprietary CAD geometry or corporate confidential data).

---

## 💡 Feature Suggestions

We love ideas that make CAD automation smoother for engineers! Typical areas of improvement:
- Additional component classification rules (e.g. thermal pads, sealants, harness clips).
- Additional default materials and densities.
- Support for alternate enterprise reference set conventions.
- Enhanced Excel styling or summary rollups.

Please open a GitHub Issue with the prefix `[Feature Request]` explaining your use case and why it would benefit other engineering teams.

---

## 🛠️ Pull Request Process

1. **Fork the repository** on GitHub.
2. **Create a descriptive feature branch**:
   ```bash
   git checkout -b feature/fastener-classification-update
   ```
3. **Make your changes**:
   - Maintain PEP 8 style guidelines.
   - Keep comments and docstrings clear.
   - Ensure backward compatibility across NX versions (avoid APIs introduced only in the most recent patch unless guarded with `hasattr` or `try...except`).
4. **Verify Your Changes**:
   - Test in an active Siemens NX session with complex multi-level assemblies.
   - Test re-running over existing Excel files to verify that override persistence remains intact.
5. **Commit and Push**:
   ```bash
   git commit -m "feat: enhance classification for thermal interface materials"
   git push origin feature/fastener-classification-update
   ```
6. **Open a Pull Request** against `main`. Describe your changes and mention any related issues.

---

## 📜 Code of Conduct

Please keep discussions respectful, constructive, and focused on helping fellow engineers solve CAD automation challenges.

