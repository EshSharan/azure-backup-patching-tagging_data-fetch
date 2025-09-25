⸻
Azure Automation Scripts
This repository contains three Python scripts for managing, auditing, and reporting Azure resources:
1. Backup Script – Fetches Azure VM backup policies and protected items.
2. Tagging Script – Retrieves Azure resources and their tags.
3. Patching script - Fetches Azure VM AUM Data.
4. Main - Orchestration script that calls other 3 script.
⸻
Prerequisites
Before running the scripts, ensure the following:
1. Python Environment
• Python 3.9+ is recommended.
• Install required packages using pip:

pip install azure-identity azure-mgmt-subscription azure-mgmt-compute azure-mgmt-recoveryservices azure-mgmt-recoveryservicesbackup azure-mgmt-resource pandas openpyxl requests