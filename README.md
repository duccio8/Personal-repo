[README.md](https://github.com/user-attachments/files/28285098/README.md)
# Cockpit Statistics

Python script that pulls visit data from the Matomo Analytics API
(`Live.getLastVisitsDetails`) for the SSM Cockpit site and produces an
Excel workbook with three sheets:

1. **Unique Users** — all unique user IDs since go-live, with the total
   count shown prominently at the top of the sheet.
2. **Active Users by Area** — top users for the reporting window, split
   by `dimension1` (business area), with a bar chart.
3. **Top Tools** — most visited tools during the reporting window, with
   a bar chart.

## Project layout

```
.
├── Cockpit statistics.py   # main script
├── requirements.txt        # Python dependencies
├── .env.example            # template for the secret token (copy to .env)
├── .gitignore              # excludes .env and the generated xlsx
└── README.md
```

## Requirements

* Python 3.10+
* Network access to `ecbdata.matomo.cloud` (typically the project VM /
  corporate VPN).

Install dependencies:

```bash
pip install -r requirements.txt
```

## Configuration

The Matomo auth token is **not** stored in the code. It is loaded from
a local `.env` file via `python-dotenv`.

1. Copy the template:

   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and set your real token:

   ```env
   MATOMO_TOKEN=your_real_matomo_token
   ```

3. `.env` is git-ignored — it will never be committed.

Optional settings (constants at the top of `Cockpit statistics.py`):

| Constant             | Meaning                                     |
| -------------------- | ------------------------------------------- |
| `MATOMO_URL`         | Matomo base URL                             |
| `SITE_ID`            | Matomo site id                              |
| `GO_LIVE_DATE`       | Anchor date for Sheet 1 (unique users)      |
| `REPORT_START`       | Start of the reporting window (Sheet 2 & 3) |
| `REPORT_END`         | End of the reporting window (Sheet 2 & 3)   |
| `TOP_USERS_PER_AREA` | Top-N users shown per business area         |
| `TOP_TOOLS`          | Top-N tools in Sheet 3                      |
| `OUTPUT_FILE`        | Name of the generated Excel file            |

## Run

```bash
python "Cockpit statistics.py"
```

The script will:

1. Load the token from `.env`.
2. Page through the Matomo API week by week from `GO_LIVE_DATE` to
   `REPORT_END`.
3. Compute the three datasets.
4. Write the report to `cockpit_statistics.xlsx` in the current
   directory.

## How the data is derived

* **User identifier** — `userId` if present, otherwise the anonymous
  `visitorId` (one record per visit).
* **Business area (Sheet 2)** — visit-level field `dimension1` (custom
  dimension configured in Matomo).
* **Tool (Sheet 3)** — first path segment after `/tools/` in the action
  URL (e.g. `https://ssmcockpit.escb.eu/tools/athena/...` → `athena`).
  Only entries with `actionDetails[].type == "action"` are counted, so
  content impressions / interactions are excluded.

## Security notes

* The token is loaded from environment / `.env` — never from source.
* `.env` and the generated `cockpit_statistics.xlsx` are in `.gitignore`.
* If a real token has ever been committed in the past, revoke it in
  Matomo (Personal → Security → Auth tokens) and generate a new one.
