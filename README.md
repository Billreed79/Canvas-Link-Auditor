# Canvas External Link & APA Citation Auditor

A Python utility for auditing external links in a Canvas LMS course and generating an Excel report containing link locations, HTTP status information, external resource metadata, and automatically generated APA 7-style citations.

Designed for instructors and course developers who want to review external resources before a new semester, identify broken or redirected links, and maintain a bibliography of resources used throughout a Canvas course.

> **Disclaimer:** This project is an independent community-developed tool and is not affiliated with, sponsored by, or endorsed by my employer, Instructure or Canvas LMS.

---

## Features

* Scans Canvas course content for external URLs
* Identifies links in:

  * Modules
  * Module items
  * Pages
  * Assignments
  * Announcements
  * Discussions
  * Quizzes
  * External tools
  * Embedded resources
  * Canvas Files
* Associates links with their Canvas module when possible
* Excludes Canvas-internal links
* Detects duplicate external URLs
* Checks external URLs for HTTP status
* Identifies redirects
* Retrieves metadata from external webpages
* Generates approximate APA 7-style citations
* Identifies common resource types such as:

  * Webpages
  * PDFs
  * YouTube videos
  * GitHub repositories
* Assigns a citation confidence level
* Generates a formatted Excel workbook
* Creates a separate citation-review report for citations requiring manual verification

---

## Excel Report

The script generates:

```text
canvas_external_links.xlsx
```

The workbook contains several worksheets.

### Summary

Provides an overview of the audit:

* Canvas URL
* Course ID
* Report generation date
* Total external link occurrences
* Unique external URLs
* Unique external domains
* Working URLs
* Problem URLs
* APA citation confidence statistics
* Information grouped by Canvas Module

### All External Links

Contains every discovered external link and its context.

Example:

| Module   | Content Type | Content      | Link Text            | URL         | Status |
| -------- | ------------ | ------------ | -------------------- | ----------- | ------ |
| Module 1 | Page         | Introduction | Fedora Documentation | https://... | OK     |
| Module 4 | Assignment   | Hyper-V Lab  | Microsoft Learn      | https://... | OK     |

### Unique URLs

Consolidates duplicate URLs and shows where each resource is used.

### APA Citations

Provides one citation for each unique external resource.

Example:

```text
Microsoft. (2025). Hyper-V virtualization on Windows. Microsoft Learn. https://...
```

### Citation Review

Lists resources where the automatically generated citation may require manual review.

### Domains

Summarizes external resources by domain.

### Broken Links

Lists links that return HTTP errors, timeouts, SSL errors, or other problems.

### By Module

Organizes the complete link inventory by Canvas module.

---

## Requirements

* Windows, macOS, or Linux
* Python 3.9+
* A Canvas LMS account with API access
* A Canvas API access token

Python packages:

```text
requests
pandas
openpyxl
beautifulsoup4
```

---

## Installation

Clone the repository:

```bash
git clone https://github.com/Billreed79/Canvas-Link-Auditor.git
cd Canvas-Link-Auditor
```

Install the required Python packages:

```bash
python -m pip install requests pandas openpyxl beautifulsoup4 pypdf python-docx python-pptx
```

On Windows, you can alternatively use:

```powershell
py -m pip install requests pandas openpyxl beautifulsoup4 pypdf python-docx python-pptx
```

---

## Canvas API Token

The script requires a Canvas API access token.

Create a personal access token through your Canvas account's settings. The exact location may vary depending on your institution's Canvas configuration.
Usually it's located in Account → Settings → Approved Integrations → + New Access Token

Give it a description such as:

__External Link Audit__

Then copy the token. Don't put the token into the Python file!

**Never commit your Canvas API token to GitHub.**

The script reads the token from the `CANVAS_API_TOKEN` environment variable.
In Windows PowerShell, run:
### Windows PowerShell

```powershell
$env:CANVAS_API_TOKEN="your-token-here"
```

Verify that it is available:

```powershell
$env:CANVAS_API_TOKEN
```

### macOS/Linux

```bash
export CANVAS_API_TOKEN="your-token-here"
```

---

## Configuration

Open:

```text
canvas_link_auditor_apa.py
```

Change:

```python
CANVAS_URL = "https://YOUR-SCHOOL.instructure.com"
COURSE_ID = "YOUR_COURSE_ID"
```

For example, if your Canvas course URL is:

```text
https://example.instructure.com/courses/123456
```

configure:

```python
CANVAS_URL = "https://example.instructure.com"
COURSE_ID = "123456"
```

The API token does not need to be placed in the Python source code.

---

## Running the Auditor

Run:

```powershell
python canvas_link_auditor_apa.py
```

The script will:

1. Connect to Canvas
2. Retrieve course modules
3. Scan module items
4. Scan course pages
5. Scan assignments
6. Scan discussions
7. Scan quizzes
8. Extract external URLs
9. Test external URLs
10. Retrieve webpage metadata
11. Generate APA-style citations
12. Create the Excel report

The final report will be:

```text
canvas_external_links.xlsx
```

---

## APA Citation Generation

The script attempts to identify bibliographic information from metadata exposed by the external webpage.

It looks for information such as:

* Author
* Organization
* Publication date
* Page title
* Website name
* Resource type

It uses common metadata fields including:

```text
citation_title
citation_author
citation_publication_date
dc.title
dc.creator
dc.date
og:title
og:site_name
article:published_time
```

When an individual author cannot be identified, the website or organization may be used as a group author.

### Citation Confidence

Each citation receives a confidence rating.

| Confidence | Meaning                                                       |
| ---------- | ------------------------------------------------------------- |
| High       | Author, title, and publication year identified                |
| Medium     | Author and title identified, but publication date unavailable |
| Low        | Limited metadata available                                    |
| Very Low   | Insufficient metadata; manual review strongly recommended     |

**Automatically generated citations should always be reviewed before being submitted as an authoritative bibliography or reference list.**

---

## Privacy and Security

This application communicates with two types of services:

### Your Canvas instance

The Canvas API is used to retrieve course content and module information.

### External websites

When link checking and APA metadata extraction are enabled, the script makes HTTP requests to external URLs found in the course.

The script does not intentionally collect student submissions, grades, or personal student information.

However, users should review their institution's policies regarding:

* Canvas API access
* Course data
* External HTTP requests
* API tokens
* Institutional security requirements

### Protect your API token

Do **not** put your token directly in the Python source code.

Do **not** commit a token to GitHub.

Do **not** include a token in screenshots or bug reports.

If a token is accidentally published, revoke it immediately through Canvas and create a new one.

---

## Limitations

The auditor cannot guarantee that every external URL in a Canvas course will be discovered.

Canvas courses can contain content in many forms, including:

* Rich HTML
* External tools
* Embedded applications
* Third-party LTI integrations
* Files
* Custom JavaScript
* Dynamic content
* Content loaded after the page is rendered

Some external services also prevent automated requests or require authentication.

As a result:

> **The report should be considered an audit and maintenance aid rather than a definitive inventory of every URL associated with a course.**

---

## APA Limitations

APA citation generation is based on metadata available from external websites.

Websites frequently have incomplete, inconsistent, or incorrect metadata. The generated citations therefore require instructor review.

In particular, manually verify:

* Author
* Organization name
* Publication date
* Page title
* Website name
* Resource type
* URL
* Whether a retrieval date is appropriate

The script does not attempt to replace a formal APA reference-management system.

---

## Suggested Course Maintenance Workflow

A useful workflow for instructors is to run the auditor:

### Before each semester

Run the report and review:

```text
Broken Links
Citation Review
Unique URLs
By Module
```

### During course development

Use the **APA Citations** worksheet to identify resources that need to be added to:

* Course references
* Reading lists
* Module resources
* Assignment instructions
* Course bibliography

### After updating links

Run the auditor again and compare the results with the previous report.

---

## Example Use Case

An instructor has a 13-module System Administration course containing:

* Microsoft documentation
* Fedora documentation
* LabEx labs
* Azure resources
* Linux documentation
* YouTube instructional videos
* GitHub repositories
* Vendor documentation

Running the auditor produces an Excel workbook showing where each resource appears in the course and provides an automatically generated APA-style citation.

For example:

```text
Module 4 – Virtualization
    Assignment: Hyper-V Lab
    Link: Microsoft Learn
    URL: https://...
    Status: OK

    APA:
    Microsoft. (2025). Hyper-V virtualization on Windows.
    Microsoft Learn. https://...
```

This makes it easier to perform a resource and citation review before publishing a course.

---

## License

This project is released under the **MIT License**.

See [LICENSE](LICENSE) for the complete license text.

Copyright (c) 2026 William Reed.

---

## Contributing

Contributions, bug reports, and improvements are welcome.

Before submitting a pull request:

1. Test the changes against a Canvas course.
2. Do not include Canvas API tokens.
3. Do not include student information or other sensitive course data.
4. Document significant changes.
5. Verify that existing functionality continues to work.

---

## Disclaimer

This project is provided "as is" without warranty.

It is an independent tool and is not affiliated with, sponsored by, or endorsed by Instructure, Canvas LMS, or any external services accessed by the application.

Users are responsible for ensuring that their use of the application complies with their institution's policies, Canvas terms, applicable laws, and the terms of external websites accessed by the auditor.
