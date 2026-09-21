# Call_Log_6

Apps Script for the roster workbook. The `final-6roster-architecture` branch contains the proposed 6Roster tools.

## Student Connection Calls

On **Call Entry**, saving and PowerSchool logging are separate actions:

- Use **Teacher Tools > Save SCC Call Entry** to save the call to the SCC tab. Successful calls save to SCC Notes; unsuccessful calls save to the next available Attempt column.
- From **Call Entry**, use **Teacher Tools > Log Selected SCC/Attempt in PowerSchool** to prepare a PowerSchool log from the current form without changing the SCC tab.
- From the **SCC** tab, select one populated **SCC Notes** or **Attempt 1–5** cell and use the same menu action. An SCC Notes selection uses the **SCC Success** PowerSchool settings; an Attempt 1–5 selection uses the **SCC Attempt** settings and the text from that selected cell. The older function name `saveSccAndOpenPowerSchool` remains as a compatibility alias, but it also logs only.

The call is usually with a parent at the start of the semester. The PowerSchool helper never clicks Submit: review the log and submit it yourself.

After selecting a student on **Call Entry**, select cell **B15**, labeled **Demographics Correct?**, to open that student's PowerSchool Demographics screen. This is a Google Sheets selection action, so choose another cell before selecting B15 again if you want to reopen it. Refresh the Sheet after updating Apps Script and reload the unpacked extension after updating its files.

**Teacher Tools > Refresh Student Data** also refreshes counselor information. Counselor names come from the source workbook's **Upstream** `COUNSELOR` field and are saved to **SCC column AA**. Counselor emails are reused from the existing **SCC Counselor/Counselor Email** directory; the script leaves an unknown email blank rather than guessing it. **Call Entry C40** displays the selected student's counselor and email automatically and is preserved when the Call Entry form is reset.

## PowerSchool selections in the Settings tab

On **Instructions and Settings**, use the **PowerSchool contact log settings** table in A34:G38. The four rows are **SCC Success**, **SCC Attempt**, **ECC Conversation**, and **ECC Attempt**. Enter the exact PowerSchool option values in **Log Type value** (B) and **Subtype value** (D); C and E are human-readable labels. An optional **Additional dropdowns (JSON)** cell (F) accepts an array such as `[{"name":"result","value":"no_answer"}]`. Leave B and D both blank to select Type/Subtype manually on that PowerSchool log. The known ECC values are prefilled; SCC values need to be captured from your school's form.

Each handoff reads the current Settings cells and sends the matching row's values to the extension. Editing these cells changes the next handoff without reinstalling the extension. The temporary extension capture tool can copy Type/Subtype values for pasting into columns B:E. Do not put student numbers or notes in this settings table. The ECC date controls will be handled in extension code once their actual PowerSchool controls are captured.

## ECC notes in the final roster

Use the ECC tab directly. The new note columns are **Overall (E), Classes (F), Grades (G), and Socially (H)**. The script finds them by normalized header text, including the line breaks currently in the headers. **ECC Date (D)** and at least one note field are required to log an entry. The prior history in **Old ECC Dates and Notes (I)** stays visible while you write.

Select one student row on the ECC tab and use **Teacher Tools > Log Current ECC in PowerSchool**. The script handles only that selected row, archives its dated note in column I, and opens one PowerSchool handoff for review. It never submits the PowerSchool log.

- **ECC Type** may be set to Conversation or Attempt when that column is present; otherwise the current ECC defaults to Conversation. An Attempt needs a free Attempt 1/2 cell.
- The former batch actions are no longer shown in Teacher Tools. Existing **Ready to Log** or **ECC Type** columns are left in place so no roster data is deleted or shifted.
- Recent notes in E:H and the date in D remain visible. An identical dated entry is not appended to history twice, so the same current row can be retried without duplicating its archived note.
- If the action fails, open the **Automation Log** tab to review the recorded error. Do not share a handoff marker outside your authorized student-data systems.

Run the ECC checks from the repository root with `node --test tests/ecc.test.cjs`.
