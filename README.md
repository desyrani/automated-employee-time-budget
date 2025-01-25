**Overview**

This repository contains the codebase for the Employee Time Budget system, designed to help organizations manage and monitor employee time allocation effectively. The system integrates with tools like Google Sheets and Google Calendar to automate tracking, analysis, and reporting of employee time budgets.

**Key Features**

- Automated Time Tracking: Integrates with Google Calendar to sync and analyze time allocation.
- Customizable Workflows: Features tailored functions to accommodate organization-specific requirements.
- Interactive UI: Provides an easy-to-use interface for managing and visualizing time budgets.

**Repository Structure**

**1. Code**

This folder contains the core logic for the Employee Time Budget system, including scripts for:
    - Fetching calendar events and syncing them with Google Sheets.
    - Performing calculations for weekly and monthly time budgets.
    - Automating the exclusion of holidays, breaks, and out-of-office activities.

**Key Functions:**

- Synchronization with Google Calendar.
- Time allocation calculations.
- Exclusion of specific activities (e.g., holidays, breaks).


**2. Custom**

This folder includes custom scripts and configurations to adapt the system to specific organizational needs.

**Highlights:**

- Custom rules for time budget calculations.
- Dynamic exclusion of activities based on input from Google Sheets.
- Integration with public holiday and collective leave datasets.

**3. UI**

This folder contains code for the user interface, which enhances user interaction with the system.

**Features:**

- Custom menus for Google Sheets.
- Search functionality to query activities or employees by name or email.
- Tools for real-time updates and reporting.
