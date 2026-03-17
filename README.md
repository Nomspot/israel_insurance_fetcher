# Government Insurance Data Extractor

## Overview

A Python-based automation tool designed to extract insurance data from a legacy government system that does not provide a public API.

The system automates authentication, handles session persistence, processes user input from Excel, and extracts structured insurance data (Life & Health) through browser automation.

---

## Key Features

* Automated login to a government system (including MFA flow)
* Session persistence using cookies to avoid repeated authentication
* Intelligent handling of dynamic UI elements (dropdowns, modals, async content)
* Extraction of Life and Health insurance data
* Aggregation and normalization of insurance records
* Excel-compatible output format
* GUI interface for easy data input and output

---

## Technologies Used

* **Python**
* **Selenium / undetected-chromedriver** – browser automation & bot-detection bypass
* **Tkinter** – GUI for user interaction
* **SQLite / Pickle** – session and local data persistence
* **psutil** – process cleanup and resource management

---

## How It Works

1. User pastes client data (copied from Excel) into the GUI
2. The system ensures an active authenticated session:

   * Loads cookies if available
   * Otherwise performs login + waits for SMS (MFA) confirmation
3. The script navigates the government system
4. Inputs client details (ID, birth date, issue date, etc.)
5. Extracts:

   * Life insurance data
   * Health insurance data
6. Aggregates results by company and insurance type
7. Outputs formatted results (ready to paste back into Excel)

---

## Example Workflow

```text
Excel Input → GUI Input → Automation → Data Extraction → Aggregation → Excel Output
```

---

## Project Structure

```text
main.py              # Main application (GUI + automation logic)
site_cookies.pkl     # Stored session cookies
```

---

## Output Format

The system returns a structured line containing:

* Client Name
* ID
* Issue Date
* Birth Date
* Phone
* Insurance Company
* Life Insurance Summary
* Health Insurance Summary
* Additional Categories (Critical illness, accidents, nursing)
* Reason

---

## Key Technical Highlights

### 🔹 Session Reuse

* Stores cookies locally to avoid repeated logins
* Automatically detects expired sessions and re-authenticates

### 🔹 MFA Handling

* Uses a GUI-based waiting mechanism for SMS verification
* Synchronizes automation flow with user confirmation

### 🔹 Dynamic UI Handling

* Custom retry logic for dropdown selections
* Handles stale elements and asynchronous page loads

### 🔹 Data Aggregation

* Groups insurance policies by company and type
* Calculates total premiums
* Produces clean, human-readable output

---

## Limitations

* Requires valid credentials for the government system
* Depends on the current structure of the target website (may break if UI changes)
* Not publicly deployable due to sensitive data access

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure credentials

Edit the script and set:

```python
SING_IN_ID = "your_id"
SIGN_IN_PASSWORD = "your_password"
```

### 3. Run

```bash
python main.py
```

---

## Notes

* This project interacts with restricted systems and sensitive data
* Intended for internal or educational purposes only

---
