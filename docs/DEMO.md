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
- How the Copilot round trip connects attachments, application context, browser
  automation, structured output, and shared workspace state.
- Why structured assistant output still requires source verification.
- Where to separate code-behind into testable services as the application grows.

See [the architecture walkthrough](ARCHITECTURE.md) for the corresponding source.

## Optional Copilot round trip

This extension needs WebView2 Runtime, internet access, and whatever sign-in or
upload capabilities the public Copilot website currently requires. It is not part
of the offline five-minute demo. Use **only** the supplied synthetic documents;
leave all personal identifiers blank and review the provider's data-sharing terms.

1. Open **Assistant** in the workspace and complete sign-in if the website requires
   it. This is an embedded website, not the installed Windows Copilot app.
2. Use a single synthetic demo image to keep the evidence set unambiguous.
   **Gather All Dates** attempts to attach a stored file and submit a prompt
   assembled from workspace context. Confirm the correct attachment is visible;
   if automation cannot attach it, use the clipboard/manual upload route.
3. Inspect the submitted context and the returned `[APP_UPDATE]` block. The demo
   image contains illustrative employment periods, not real qualification records;
   missing qualification dates should remain blank.
4. Observe whether supported results populate the workspace. Audit/date extraction
   can auto-apply detected updates; **Apply Suggested Dates** also supports reading
   structured output from the conversation or clipboard.
5. Compare every populated value with the synthetic source. An empty update is
   valid when no supported facts are found. Do not treat a generated finding as
   a verified compliance decision.

**What to explain in an interview:** the integration handles the app-to-browser-to-app
workflow, while Copilot supplies the external response. Website changes, account
capabilities, attachment failures, and uncertain extraction remain explicit
limitations. No live Copilot exchange is claimed by the repository screenshots.
