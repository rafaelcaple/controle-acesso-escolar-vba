# School Access Control System

A QR code-based student access control system built in VBA for Microsoft Excel, developed between June 2022 and August 2023 and **deployed in production at CIE Miécimo da Silva, Rio de Janeiro**.

## Background

The school had no reliable way to track cafeteria entries, which led to duplicate records and inconsistencies in attendance data. The system was built to solve that problem: students scan the QR code on their badge, the entry is logged automatically, and duplicate access with the same student ID is blocked at the source.

Due to its success, the system was later **expanded from cafeteria control to full school entry and exit tracking**, giving the administrative staff complete visibility into student attendance throughout the day.

## My Contributions

This was a team project with 4 members. I was responsible for the implementation of the system, including:

- Entry logging via QR code badge: a physical scanner reads the badge and passes the student ID to the system, which validates and registers the entry automatically
- Duplicate access prevention: the same student ID cannot be registered twice in the same session
- Student photo display at the moment of check-in, allowing staff to visually verify the student's identity against the badge being scanned
- Manual search fallback for students without their badge
- Data saving and automatic backup

## Impact

- Eliminated duplicate and inconsistent cafeteria records
- Enabled visual identity verification at entry points
- Improved security and traceability for student entry/exit throughout the school
- Adopted daily by the administrative staff

## Features

- Main dashboard to access all system functionalities
- Entry and Exit Logging: Register student attendance when a physical QR code scanner reads the student's badge and submits the ID to the system
- Student Photo Display: Shows the student's photo at check-in for identity verification
- Manual Search: Look up and register students by name or ID when no badge is available
- Duplicate Prevention: Blocks multiple entries with the same student ID
- Save Spreadsheet: Save current system data
- Spreadsheet Backup: Create backup copies of spreadsheets
- Record Cleanup: Remove outdated records

## Prerequisites

- Microsoft Excel 2010 or higher
- Macros enabled: Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros**

## How to run

1. Download the `planilhas/` folder
2. Open the desired spreadsheet in Excel
3. Use the existing button or create a new one and link it to: `TelaPrincipal.Show`

## Configuration

- Backup directory: Folder where automatic spreadsheet backups will be saved
- Photo directory: Folder where student photos are stored
- Access control messages: Customize messages displayed when logging entries
- Spreadsheet password: Set a password to protect the system and spreadsheet data

## Modules and Forms

### Modules

- `var.bas` — Global system variables
- `SalvarLimpar.bas` — Backup and cleanup functions
- `NumDeRegistrados.bas` — Record counting
- `Relogio.bas` — Date/time functions
- `ScrollMouse.bas` — Enhanced scroll in forms

### Forms

- `TelaPrincipal` — Main interface for accessing system functionalities
- `Verificador / VerificadorSaida / Verificador2 / Verificador3` — Forms responsible for entry and exit control
- `Pesquisa` — Form for searching existing records by name, ID, or other criteria

## Code Structure

The VBA code is organized inside `src/`, divided by functionality:

- `entrada-saida/` — modules and forms for entry/exit control
- `refeitorio/` — modules and forms for cafeteria control
