# aud-metaphor-thesis-scripts
Python scripts used in my PhD thesis (UCL IOE), for corpus processing and analysis

Thesis Scripts: Corpus-Assisted Metaphor Analysis (AUD, UK)

Python scripts used in the PhD thesis:

A Corpus-Assisted Analysis of Metaphor Use in Discourse Surrounding Alcohol Use Disorder in the UK 

README

This repository contains scripts supporting (i) corpus data acquisition and restructuring, (ii) USAS tag mapping and tag-set construction, (iii) defensible sampling of types for manual coding workflows, and (iv) quantitative comparison of metaphor vehicle-group frequencies across corpora (e.g., SSC vs LEC), including significance and effect size estimation and plotting. 

What this repo includes:

The core scripts are numbered to reflect a typical workflow order (data → structuring → annotation support → sampling/estimation → inferential comparison). Some scripts are templates and require you to replace placeholders (e.g., PATH/TO/...) and set file paths inside the script before running.

Repository structure 
scripts/
  01_scrapereddit.py
  02_organiseredditxml.py
  03_examplespider.py
  04_mappingusastagsandmanuallabels.py
  05_creatingpotentialtagsets.py
  06_takeasampleoftypes.py
  07_frequenciesandestimates.py
  08_calculatingsignificanceandeffect.py
requirements.txt

Scripts
01 — scrapereddit.py

Communicates with Reddit’s API (via PRAW) for a specified subreddit and listing type (e.g., top, new) and returns a Pandas DataFrame containing basic post metadata (title, URL, score, comment count). By default, the DataFrame is created in memory; export can be added with to_csv() after creation. 

Note: praw is required for this script but is not listed in the provided requirements.txt. If you intend to run this script, install praw separately.

02 — organiseredditxml.py

Restructures a flat XML export of Reddit data into a hierarchical XML format grouped by subreddit. Reads <row> elements, extracts fields (e.g., Subreddit, PostID, score, title, body), and writes a new XML where each subreddit is a node containing its posts. Input/output paths must be set inside the script. 

03 — examplespider.py

A Scrapy CrawlSpider template to crawl a target website and export page-level text content to CSV. Follows allowed links and extracts page URL, HTML <title>, and concatenated paragraph text (<p>). Placeholder values (name, domain, start URL, site name) must be replaced before use. 

04 — mappingusastagsandmanuallabels.py

Maps Wmatrix USAS semantic tags onto a manually labelled list of items/types. Reads (1) a manual list (item/type + manual label + vehicle grouping), and (2) a USAS-tagged sample sheet where each instance may have multiple tags across multiple columns. Aggregates unique tags per item and exports an Excel file with Tag1…TagN columns. 

05 — creatingpotentialtagsets.py

Counts USAS semantic tags for a single vehicle grouping in an Excel sheet and outputs a results sheet containing: each tag, frequency within that VG, associated strings/items where the tag appeared, and tag definitions (looked up from a SEMTAGS sheet). 

06 — takeasampleoftypes.py

Creates a new worksheet containing a 20% sample of rows for each “type” (defined by values in Column C), retaining all rows for low-frequency types. Sampling rule: keep all rows if a type occurs <20 times; otherwise sample 20% (at least 1 row). Random sampling is re-drawn each run unless you add a fixed seed. 

07 — frequenciesandestimates.py

Estimates total metaphorical instances in the full corpus from an Excel-based coding workflow combining: (i) a 20% sampled sheet for high-frequency types, (ii) a low-frequency sheet retaining all LF types, and (iii) an optional “extra” sheet with additional coded rows. Processes relevant workbooks/sheets in a directory, outputs per-sheet/per-workbook summaries, and writes consolidated CSV results. 

08 — calculatingsignificanceandeffect.py

Compares metaphor vehicle-group frequencies across two corpora (e.g., SSC vs LEC). Computes: log-likelihood (G²) with chi-square p-values (df=1), log ratio (log2 effect size) using relative frequencies per 1,000 words. Outputs a CSV plus two figures: (i) diverging bar chart of log ratios with significance markers; (ii) bubble plot of effect size vs statistical significance (bubble size scaled by log-likelihood). Figures are saved as high-resolution PNG. 


Requirements

Workflow assumes:

Python 3.9+

pip (and ideally a virtual environment)

ability to read/write local .xlsx files

updating placeholder paths (PATH/TO/...) to your own file locations 

Non-standard library dependencies (as provided): 

pandas==2.2.3

scrapy==2.11.2

requests==2.31.0

openpyxl==3.1.5

tabulate==0.9.0

numpy==1.24.0

scipy==1.14.1

matplotlib==3.7.5

Install (recommended: virtual environment)
python3 -m venv .venv
source .venv/bin/activate    # macOS/Linux
# .venv\Scripts\activate     # Windows PowerShell
pip install -r requirements.txt


If using the Reddit script:

pip install praw

Usage notes

Several scripts require manual configuration of file paths and/or worksheet names inside the script before running. 

Many scripts assume that inputs were produced by earlier pipeline steps (e.g., scraped data; Excel workbooks containing sampled sheets; tag-mapping outputs). 

For reproducibility in sampling (06_takeasampleoftypes.py), set a fixed random seed if you need exact re-runs. 

Data availability

This repository provides code only. Research data (e.g., corpora, scraped content exports, coded spreadsheets) are not included here and may be restricted by ethics, platform terms, or licensing. Scripts are designed to run on local copies of data files. (Where appropriate, you can add a small synthetic/example file to illustrate format without sharing restricted data.)
