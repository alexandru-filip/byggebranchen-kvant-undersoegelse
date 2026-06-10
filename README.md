# Byggebranchen kvanitative undersøgelse

This repo is a collection of Python scripts designed to assist in a quantitative market research for single-family house building companies in Jylland. 

The results are used as a part of a 1st internal semester project on "Markedsforingsøkonm" education, Dania Erhvervsakademi. 

No actual data is published in this repo. 

Needs a .csv list of companies as input (get them from virk.dk) for the scripts to be computational. 


# Website Keyword Scan Script

## Script

`website_keyword_scan.py`

## What It Does

This script takes the company list from `companies_with_public_financials.csv`, finds each company's website, and scans the website for selected keywords.

It is meant to help identify companies that mention relevant terms such as:

- `nybygget`
- `nybyggeri`
- `nybyg`
- `familiehus`
- `typehus`
- `parcelhus`
- `boligbyggeri`

Only companies with keyword hits are saved by default.

## Why It Exists

The financial enrichment script can identify company size, but it does not explain what kind of construction work each company focuses on.

This website scanner adds a qualitative layer: it checks whether a company appears to work with new-build homes, family houses, type houses, parcel houses, or related housing construction topics.

That makes it easier to filter the company list by both:

- financial size
- relevant business focus

## Input

By default, the script reads:

```text
companies_with_public_financials.csv
```

The input should contain company names and preferably CVR numbers. The existing enriched CSV already has both:

- `input_company_name`
- `resolved_cvr`

## Output

By default, the script writes:

```text
website_keyword_hits.csv
```

Important output columns:

- `input_company_name`
- `cvr_number`
- `website`
- `matched_keywords`
- `keyword_count`
- `matched_pages`
- `pages_checked`
- `status`
- `notes`

`keyword_count` is the number of distinct keywords found on the scanned website pages.

If a company website is found but `robots.txt` does not allow crawling, the script still saves the company in the CSV. In that case:

```text
website = the website that was found
keyword_count = 0
status = blocked_by_robots
```

This makes it clear that the company had a website, but the site was not scanned.

## How Website Lookup Works

The script first tries to find a website through `apicvr.dk` using the company's CVR number or company name.

If no website is returned, it tries likely `.dk` domains based on the company name. For example:

```text
DML HUSE A/S -> dmlhuse.dk / dml-huse.dk
```

The guessed website is only accepted if the page content appears to match the company name.

## How Keyword Scanning Works

For each website, the script:

1. Checks whether `robots.txt` allows access.
2. Opens the front page.
3. Finds internal links on the same website.
4. Prioritizes links that look relevant, such as pages containing words like `nybyg`, `bolig`, `huse`, `projekter`, or `referencer`.
5. Scans up to 12 pages per website by default.
6. Saves the company if any keyword is found.

## Editing Keywords

Keywords are stored in:

```text
keywords.txt
```

Add one keyword per line. Blank lines are ignored.

Example:

```text
nybyggeri
familiehus
typehus
parcelhus
renovering
totalentreprise
```

## Running In VS Code

Open the `codes/temp` folder in VS Code.

Then open:

```text
website_keyword_scan.py
```

Press **Run Python File**.

No command-line arguments are needed if the script is in the same folder as:

```text
companies_with_public_financials.csv
keywords.txt
```

## Useful Options

Preview without using the network:

```bash
python3 website_keyword_scan.py --dry-run
```

Test only the first 10 companies:

```bash
python3 website_keyword_scan.py --limit 10
```

Save companies even when no keyword is found:

```bash
python3 website_keyword_scan.py --include-non-hits
```

Scan more pages per website:

```bash
python3 website_keyword_scan.py --max-pages 25
```

Disable guessed domains and only use websites found through CVR data:

```bash
python3 website_keyword_scan.py --no-domain-guess
```

## Limitations

Website discovery is best-effort. Some companies do not publish a website in CVR data, and guessed domains can miss the correct site.

The script scans normal HTML pages. It does not scan PDFs, images, JavaScript-only content, or social media pages.

Keyword matches are literal text matches. They indicate that the website mentions the terms, not that the company definitely specializes in that service.

Some websites may block automated requests or disallow crawling through `robots.txt`. The script respects that.

When that happens, the website is still included in the output CSV with `status = blocked_by_robots`.


# CVR Financial Enrichment Script

## Script

`enrich_cvr_financials.py`

## What It Does

This script enriches a list of Danish construction companies with financial data from public annual reports.

It reads company names from an `.xlsx` or `.csv` file, resolves each company to a CVR number when needed, finds the latest public annual-report XML/XBRL file, and exports a CSV containing:

- company name
- resolved CVR number
- resolved legal company name
- industry code and description
- latest annual-report period
- `nettoomsætning`, when available
- `bruttofortjeneste` / `bruttoresultat`, when available
- average employees from the annual report, when available
- status fields explaining what was found or missing

## Why It Exists

The original company list only contained names and/or CVR numbers. Basic CVR company data does not reliably include revenue.

Revenue-like figures are usually found in annual reports instead. Danish annual reports are published publicly through Virk, often as XBRL/XML files. This script uses those public report files to extract financial facts automatically where possible.

It extracts both `nettoomsætning` and `bruttofortjeneste` because many smaller Danish companies do not publish exact net revenue. In those cases, gross profit is often the best available public size indicator.

## Data Sources

The script uses:

- `apicvr.dk` to resolve company names to CVR numbers when the input file does not already contain CVR numbers.
- `distribution.virk.dk/offentliggoerelser/_search` to find public annual-report publications.
- `regnskaber.virk.dk` to download the annual-report XML/XBRL files.

No Proff.dk scraping is used.

## How To Run

From the workspace folder:

```bash
python3 enrich_cvr_financials.py /path/to/names.xlsx
```

If the spreadsheet has a CVR column:

```bash
python3 enrich_cvr_financials.py /path/to/names.xlsx --cvr-column CVR
```

The default name column is `Navn`. If your file uses another column name:

```bash
python3 enrich_cvr_financials.py /path/to/names.xlsx --name-column "Company Name"
```

## Output

The default output file is:

```text
companies_with_public_financials.csv
```

Important output columns:

- `net_revenue_dkk`
- `net_revenue_tag`
- `gross_profit_dkk`
- `gross_profit_tag`
- `average_employees_xbrl`
- `financial_status`
- `annual_report_xml_url`

`financial_status` explains whether net revenue, gross profit, both, or neither were found.

## Notes And Limitations

Not every company has a public annual report with usable XML/XBRL data.

Some companies publish `nettoomsætning`; many smaller companies only publish `bruttofortjeneste` or `bruttoresultat`.

Name-to-CVR matching is best-effort. If the input contains CVR numbers, use `--cvr-column` for more reliable results.

The script saves progress after every company and uses a local cache so it can be resumed without repeating every network request.
