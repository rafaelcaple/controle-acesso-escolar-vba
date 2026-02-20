# School Access Control System

## About

Access control system developed in VBA for Microsoft Excel between June 2022 and August 2023, **applied in a real environment at CIE Miécimo da Silva**.

### Original Problem

The school faced inconsistencies in cafeteria access control, with no reliable way to track entries and prevent duplicate records.

### Solution

Badge-based system that automatically logged access entries, implementing validations to prevent duplicates and ensure full traceability.

### Evolution

Due to the success of the solution, the system was later **adapted for general school entry and exit control**, significantly improving security and organization.

### Impact

- Eliminated inconsistencies in cafeteria access control
- Prevented duplicate entries
- Improved security in student entry/exit control
- System used daily by the administrative staff

## Features

- Main dashboard to access all system functionalities
- Entry and Exit Logging: Register student attendance points
- Search by Record: Look up existing records by name, ID, or other criteria
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
- Photo directory: Folder where photos will be stored
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
