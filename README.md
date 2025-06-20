# Mendix security scanner

This script analyzes data exposure in a Mendix application by identifying all entities accessible to a given user session. It collects object and attribute data from exposed entities and generates a comprehensive Excel report.

It is primarily designed for **security auditing and compliance reviews**, helping identify if session-based access controls are too permissive.

---

## Why This Matters

This scanner helps you answer:

> _"Can this session access or modify data it shouldn't be allowed to?"_

How:

- If an entity is listed, the session can **read it**.
- If any field is not readonly, the session can **write to it**.
- **Bolded fields and red columns** indicate insecure write access.
- Console output of microflows shows possible **indirect manipulation** paths.

---

## Requirements

- Python **3.7+**
- `openpyxl` package
- `requests` package
  
Install dependencies with:

```bash
pip install requests openpyxl
```

Usage
```bash
python mendix_scan.py --url https://my-app.mendixcloud.com/ --cookie "abc123" "-l 500"
```
## Arguments


| Argument           | Description                                                                 |
|--------------------|-----------------------------------------------------------------------------|
| `--url`            | **Required.** Mendix app base URL (with or without `/xas/`)                |
| `--cookie`         | **Required.** Raw session ID (e.g. just `XASSESSIONID`) |
| `--output`         | Optional. Path to output Excel file. Default: `SecurityScan_Report_<timestamp>.xlsx` |
| `--limit`, `-l`    | Optional. Max objects to retrieve per entity. Default: `100`               |
| `--microflow`, `-m`| Optional. Print microflow access details to console                        |
| `--proxy`, `-p`    | Optional. Proxy URL (e.g. `http://127.0.0.1:8080` for Burp/ZAP)  

> **Note:** All cookies from `Set-Cookie` are handled automatically.  
> The script will fail gracefully with `Application requires authentication` if anonymous uses is disabled.

---

## Example

```bash
python mendix_scan.py --url https://my-app.mendixcloud.com/ --cookie "abc123" -l 500
```
## Output

An Excel report containing:

- A **Summary** tab with object counts and writable field counts per entity.
- **Individual tabs** for each entity:
  - Up to 100 records with attributes and editability info.
  - Writable fields in **bold** and **highlighted in red**.
- Metadata includes each object's **GUID**.

Default filename example:
`SecurityScan_Report_2025-06-20_15-30.xlsx`

---
## License

MIT License — use at your own risk. For internal audits and testing purposes only.
