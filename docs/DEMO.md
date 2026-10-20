# Five-minute reviewer demo

This walkthrough uses only the synthetic images in [demo/](demo/).
It needs no customer records, credentials, assistant account, or Office documents.

## 1. Start the app

Follow the [local build instructions](../README.md#run-locally), then choose
**Workspace** from the dashboard.

## 2. Compare two documents

Add or drag these images into the canvas:

- [Synthetic review summary](demo/synthetic-review-summary.png)
- [Synthetic date checklist](demo/synthetic-date-checklist.png)

Move and resize them side by side. Try changing each document's zoom without
changing its frame size. Hide a document, then restore it from the file drawer.

**What to look for:** independently managed document state rather than a single
static PDF view.

## 3. Check the date and history tools

Open **Work History** in the workspace (or **Dates & Work History** on the
dashboard). Enter two synthetic employment periods:

| Description | Start | End |
| --- | --- | --- |
| Demo period A | 01/01/2023 | 01/01/2024 |
| Demo period B | 07/01/2023 | 01/01/2024 |

The combined history should count the overlap once: **12 whole months**, not 18.
The per-row month values can still total 18; the overall figure uses merged
intervals. Leave real names, SSNs, phone numbers, and identifiers blank.

Try the date calculator and the past/future quick-reference toggle.

## 4. Save and resume

Save a workspace to a local writable folder. Close and reopen the application,
then load that workspace. Confirm the embedded documents, frame positions, zoom,
and history return.

**What to look for:** a saved review session, not just a list of paths to documents.
Do not commit the resulting `.lqw` file.

## 5. Inspect the criteria and optional tools

Open **Review Criteria** from the dashboard, then **Edit Minimums...**. Review the
starter reference and empty organization/industry templates.

Word and Excel drawers require their desktop applications. Browser and assistant
features need WebView2 Runtime; assistant sign-in and data sharing are optional
and governed by the external provider. The core demo does not need these features.

## Suggested interview talking points

- Why a persistent canvas suits comparative review better than a linear viewer.
- How overlapping time intervals differ from a sum of individual employment rows.
- Why embedding bytes improves portability while increasing file size and
  privacy obligations.
- What native Office integration buys, and the lifetime/compatibility costs.
- Why structured assistant output still requires source verification.
- Where to separate code-behind into testable services as the application grows.

See [the architecture walkthrough](ARCHITECTURE.md) for the corresponding source.
