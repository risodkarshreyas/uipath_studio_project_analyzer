# uipath_studio_project_analyzer


On macOS, the command is python3. Try:
bashpython3 uipath_analyzer.py ./MyProject --output all --report-dir ./reports

Also, make sure you have the required libraries installed first:
bashpip3 install openpyxl pandas tabulate

If you don't have pip3 either, you may need to install Python first. What does this return?
bashpython3 --version


How to run
# Scans any folder depth, finds every project.json automatically
python3 uipath_analyzer_v2.py ./MyRepo --repo --output all --report-dir ./reports

# Still works as single-project too (no --repo flag needed)
python3 uipath_analyzer_v2.py ./MyProject --output all




✅ Fully Airgapped Compatible
The tool has zero network dependencies at runtime. Here's exactly what it does and doesn't do:
ComponentNetwork?DetailsXAML / XML parsing❌ None
Python's built-in xml.etree.ElementTree — stdlib onlyActivity classification❌ None
Pure dictionary + regex lookups hardcoded in the scriptSelector deep-parsing❌ NoneLocal string/XML operationsReusability scoring❌ None
Rule-based arithmetic — no model involvedExcel export❌ Noneopenpyxl writes files locallyCSV / JSON export❌ None
pandas + json — all localDescriptions & recommendations❌ None
Static lookup tables in WORKFLOW_PATTERNS and ACTIVITY_DESCRIPTIONS dictionaries




📊 Reusability Matrix — two levels
Workflow-level — scored per .xaml file across all projects:
SignalWeightHow many projects contain this workflowup to +20 ptsHow many times it's invoked by othersup to +15 ptsWorkflow name pattern match (e.g. SendEmail, ErrorHandler, Login)up to +20 ptsHigh-value activity categories present+3 pts each
Activity Category-level — scored per category (Excel, Orchestrator, etc.):
ScoreLabelMeaning80–100🟢 HighStrong shared library candidate60–79🟡 MediumGood candidate, needs parameterisation40–59🟠 Low-MediumProject-specific for now<40🔴 LowToo selector/context-dependent
