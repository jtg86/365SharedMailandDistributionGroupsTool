Microsoft 365 Shared Mailbox & Distribution Group Management Tool

A PowerShell-based GUI tool designed to simplify common Microsoft 365 administrative tasks for support staff who do not work directly with PowerShell.

The tool provides a graphical interface for managing shared mailboxes, distribution groups, and calendar permissions, while still relying on official Microsoft modules and APIs under the hood.

Purpose

In many IT support teams, routine Microsoft 365 administration tasks are repetitive but sensitive:

granting access to shared mailboxes

managing distribution group membership

adjusting calendar permissions

These tasks are often handled by a small number of PowerShell-capable administrators, even though the operations themselves are straightforward.

This tool was created to:

reduce manual PowerShell usage for first-line support

standardize common changes

lower the risk of syntax errors

log all actions for traceability

It is intended as an operational support tool, not a replacement for proper RBAC design or administrative oversight.

Features

GUI-based interface (WPF)

Manage shared mailbox permissions

FullAccess

SendAs

Manage distribution group membership

Manage calendar permissions

Bulk input support (multiple users at once)

Action logging to local log files

Uses official Microsoft 365 PowerShell modules

Example Workflows
Grant FullAccess to a Shared Mailbox

Enter the shared mailbox address

Paste one or more user UPNs (semicolon-separated)

Select Grant FullAccess

Execute the action

Result is applied and logged

Add Users to a Distribution Group

Enter the distribution group address

Paste user UPNs

Choose Add members

Execute

Modify Calendar Permissions

Select mailbox calendar

Choose permission level

Apply changes

All actions are logged for later review.

Prerequisites

PowerShell 7.x

Microsoft 365 tenant access

Required PowerShell modules:

ExchangeOnlineManagement

Appropriate Exchange RBAC roles, depending on the operation:

Exchange Administrator (or equivalent scoped role)

Internet connectivity to Microsoft 365 services

Security & Safety Notes

The tool does not store credentials

Authentication is handled via standard Microsoft sign-in

Actions are executed in the context of the signed-in user

No secrets or tokens are written to disk

All changes should be tested in a non-production environment before broad use

Recommended practice:
Use least-privileged roles and restrict access to this tool to trusted support staff.

Logging

All actions are logged locally to the logs/ directory

Each run generates a timestamped log file

Logs are intended for operational traceability and troubleshooting

⚠️ The logs/ directory should not be committed to version control.

Intended Audience

IT support staff

Microsoft 365 administrators

Operations teams in need of safe, repeatable workflows

This tool is not intended for end users.

Limitations

No role enforcement inside the tool itself (relies on tenant RBAC)

No approval workflow

No undo/rollback functionality

Designed for interactive use, not automation pipelines

Screenshots

<img width="1691" height="996" alt="image" src="https://github.com/user-attachments/assets/e6103e6b-0fa4-4d2c-8f24-78e74e856f42" />


Disclaimer

This tool is provided as-is.
Always validate changes in a test environment before applying them in production.

License

MIT Licensepo
