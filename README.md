# Call_Log_6

Google Apps Script for the 6Roster workbook. Teacher Tools prepares PowerSchool handoffs, while the PowerSchool Helper extension completes navigation and fills supported fields.

## PowerSchool dialog workflow

Every PowerSchool action uses the same explicit workflow:

1. Choose an action from **Teacher Tools**.
2. Apps Script validates and encodes the handoff for `https://californiak12.powerschool.com`.
3. A small dialog appears. No encoded handoff marker is shown in a toast or Automation Log.
4. Click **Open PowerSchool** once. The link opens one new tab with `noopener` and `noreferrer`.
5. The extension reads the URL hash, removes it from the visible address, and completes the Demographics, SCC, or ECC preparation.
6. Review the result and click **Submit** manually. Neither Apps Script nor the extension submits a PowerSchool log.

Refreshing or opening the Google Sheet does not launch PowerSchool. Closing or canceling the dialog opens no tab.

The primary Teacher Tools actions appear in this order:

1. **Open Demographics**
2. **Save SCC Call Entry**
3. **Reset SCC Call Entry**
4. **Log Selected SCC/Attempt in PowerSchool**

Maintenance and ECC actions follow in separate menu sections.

## Demographics

On **Call Entry**, **Teacher Tools > Open Demographics** uses the student currently selected in the form, regardless of which Call Entry cell is active. On another tab, select exactly one cell containing a valid Student Number before choosing the action. Multi-cell selections, blank cells, and invalid Student Numbers do not create a PowerSchool dialog.

The former `onSelectionChange`/B15 trigger is not used.

## Student Connection Calls

Saving and PowerSchool logging are separate actions:

- **Save SCC Call Entry** writes the current call to the SCC roster. Successful calls use **SCC Notes**; unsuccessful calls use the next available **Attempt 1–5** column.
- **Log Selected SCC/Attempt in PowerSchool** prepares a handoff without saving or changing the SCC roster.
- From **Call Entry**, the logging action uses the current form. For an unsuccessful call, it reuses the number of an identical saved attempt or selects the first open Attempt column without writing to it.
- From the **SCC** tab, select one populated **SCC Notes** or **Attempt 1–5** cell. SCC Notes uses the **SCC Success** settings. Attempt columns use **SCC Attempt** settings and the matching Attempt tag.

The compatibility function `saveSccAndOpenPowerSchool` also logs only; it does not save.

## ECC

On the ECC tab, select one student row and choose **Log Current ECC Row & Open PowerSchool**. The script archives the dated note in **Old ECC Dates and Notes**, preserves the recent note fields, and displays one PowerSchool dialog.

- **ECC Date** and at least one note field are required.
- **ECC Type** may be Conversation or Attempt when the column exists; otherwise it defaults to Conversation.
- An Attempt requires an available Attempt 1 or Attempt 2 cell.
- Repeating an identical action does not duplicate the archived history entry, but it does create a fresh handoff with a unique request ID.

## PowerSchool Settings table

On **Instructions and Settings**, the PowerSchool contact log settings table uses these four workflow rows:

| Workflow | Used for |
|---|---|
| SCC Success | Successful Student Connection Call |
| SCC Attempt | Unsuccessful call and matching Attempt tag |
| ECC Conversation | ECC conversation |
| ECC Attempt | ECC contact attempt |

Enter the exact PowerSchool option values in **Log Type value** and **Subtype value**. The three date-field columns identify **Date & Time**, **Incident Date**, and **Action Date**. The Attempt tag map connects Attempt numbers to their visible PowerSchool tag labels. **Additional dropdowns (JSON)** may contain safe name/value pairs such as `[{"name":"result","value":"no_answer"}]`.

Each handoff reads the current settings. SCC dates come from the call note; ECC uses the row's ECC Date. The extension does not fill the time portion or Action Taken End Date.

## Automated tests

From the repository root, run:

```bash
node --test tests/*.test.cjs
```

The tests cover Demographics student selection, invalid and multi-cell selections, single-dialog creation, fixed-host/type URL generation, request IDs, absence of marker toasts, SCC save/log separation, SCC attempt numbering, date and tag settings, ECC archiving, and counselor synchronization.

These Node tests do not prove live Google Sheets behavior, popup-blocker behavior, authenticated PowerSchool end-to-end behavior, or compatibility with the current production PowerSchool DOM. They are simulated unit/integration tests, not end-to-end browser tests.

## Manual smoke test

- [ ] Refreshing the Google Sheet opens zero PowerSchool tabs.
- [ ] **Open Demographics** displays one dialog.
- [ ] Clicking **Open PowerSchool** opens exactly one tab.
- [ ] Call Entry uses the student selected in the form.
- [ ] One selected Student Number cell on another tab works.
- [ ] SCC Success fills Date & Time, Incident Date, and Action Date and uses the success settings.
- [ ] SCC Attempt selects the correct Attempt tag.
- [ ] A signed-out PowerSchool handoff resumes after authentication in the same tab.
- [ ] No workflow clicks **Submit**.

## Updating

After changing Apps Script, refresh the Google Sheet. After changing the extension, reload the unpacked extension from `chrome://extensions` or `edge://extensions` and refresh any open PowerSchool tab.
