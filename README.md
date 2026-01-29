# M365 Mail & Groups Drift Toolkit

A lightweight **desktop GUI tool built with PowerShell 7 and WPF** for managing Microsoft 365 / Exchange Online objects in day-to-day IT operations.

This tool is designed to simplify and standardize common administrative tasks while remaining **tenant-friendly**, performant, and safe to use in large environments.

The project is part of my personal portfolio and demonstrates practical automation, Exchange Online knowledge, and operational tooling.

---

## Overview

The M365 Mail & Groups Drift Toolkit provides a graphical interface for managing:

- Mailbox permissions
- Resource calendar permissions
- Distribution groups
- Mail-enabled security groups
- Dynamic distribution groups (read-only)

All data is loaded **on demand** using **server-side filtering**, avoiding full tenant enumeration and reducing load on Exchange Online.

---

## Key Features

### Search (tenant-friendly)
Search Exchange Online using server-side filters:
- Shared mailboxes
- Room mailboxes
- Equipment mailboxes
- Distribution groups
- Mail-enabled security groups
- Dynamic distribution groups

Results are limited per object type to ensure good performance even in large tenants.

---

### Mailbox permissions
Supported mailbox types:
- Shared mailboxes
- Room mailboxes
- Equipment mailboxes

View and manage:
- FullAccess
- SendAs

Capabilities:
- Grant permissions to multiple users or groups in one operation
- Remove selected permissions
- Supports both user and group trustees
- Changes are logged locally

---

### Resource calendar permissions (Room / Equipment)
For room and equipment mailboxes, the tool can display:
- Calendar folder permissions
- Access levels (Author, Editor, LimitedDetails, AvailabilityOnly, etc.)
- Special principals such as Default and Anonymous

Navigation feature:
- Double-click a group shown in calendar permissions to open that group directly in the Group view.

---

### Group management
Supported group types:
- Distribution Groups
- Mail-enabled Security Groups

Capabilities:
- View group members
- Add members in bulk
- Remove selected members

Input accepts:
- Email address
- User Principal Name (UPN)
- Mail nickname / alias

---

### Dynamic distribution groups
Read-only view:
- Displays RecipientFilter
- Displays RecipientContainer

Dynamic group membership cannot be modified manually by design.

---

## Bulk input format

When adding users or groups, multiple identities can be pasted at once.

Accepted formats:
- Email address
- UPN
- Alias (mail nickname, without @)

Supported separators:
- Comma
- Semicolon
- Whitespace
- New lines

Example:
