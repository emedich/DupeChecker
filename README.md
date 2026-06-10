# Lead Checker - Calder Capital

A high-performance desktop tool for cross-referencing business leads against a Keap CRM contact database.

## Features

- **Tiered Matching Logic**: Uses a sophisticated same-row matching algorithm to identify duplicates while minimizing false positives.
- **High Performance**: Optimized with `pandas` vectorization to process 50,000+ leads against 400,000+ contacts in seconds.
- **Smart Data Cleaning**:
  - **Websites**: Strips protocols, `www`, and subdirectories to match core domains.
  - **Companies**: Removes common business suffixes (Inc, LLC, Corp) and handles semicolon-separated multiple company names.
- **User-Friendly GUI**: Simple dark-themed interface for non-technical users.
- **Direct Integration**: Inserts results directly into the original leads CSV at the far left for immediate visibility.

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/lead-checker.git
   cd lead-checker
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script to open the graphical interface:
```bash
python lead_checker.py
```

1. Select your **Leads CSV** (e.g., an export from Sourcescrub).
2. Select your **Keap Contacts CSV** (your master CRM export).
3. Click **Run Check**.

The results will be saved directly into your original Leads CSV file.

## Matching Rules

A lead is marked as a duplicate (`Remove Duplicate = YES`) if:
1. **Email matches** any of the 3 email fields in Keap.
2. **OR** both **Name + Company** match on the same row in Keap.
3. **OR** both **Name + Website** match on the same row in Keap.
4. **OR** any **3 or more** of the 5 categories match on the same row in Keap.

## Building a Standalone Executable

To create a single-file `.exe` for Windows users:
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole lead_checker.py
```

## Support

For issues or questions, please contact [leads@caldergr.com](mailto:leads@caldergr.com).

---
*Created for Calder Capital*
