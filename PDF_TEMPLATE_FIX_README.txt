PDF TEMPLATE FIX
================

This version fixes the Download PDF generation:
1. Uses report_templates/College_Template.pdf as the official page background.
2. Applies the college template to every generated PDF page.
3. Keeps the CO/PO/PEO report content in the foreground.
4. Matches the generated PDF to the college template's portrait page size.
5. Automatically scales wide report tables to the available page width.
6. Fixes an existing ReportLab style bug that made table body text white/invisible.
7. Adds pypdf to requirements.txt.

Run:
    pip install -r requirements.txt
    python app.py

Then use the application's Download PDF button.


IMPORTANT REPEAT-PAGE FIX
The PDF template merge was corrected to create a fresh blank PDF page for every generated report page. The old deepcopy(template_page) approach could reuse pypdf indirect objects and make every output page show the first report page. The corrected method merges the college template first and the unique report content second.

The report includes: Subject Details, CO + Bloom table, PO table, PEO table, CO-PO matrix, CO-PEO matrix, Graphical Analysis, CO-PO Justification, CO-PEO Justification, Mapping Scale, and Faculty/HOD validation.
