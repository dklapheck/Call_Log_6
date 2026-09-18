# Call_Log_6

Apps Script for the roster workbook. The `final-6roster-architecture` branch contains the proposed 6Roster tools.

## Student Connection Calls

On **Call Entry**, use **Teacher Tools > Save SCC Call Entry** to save only to the SCC tab. Use **Save Student Connection Call & Open PowerSchool** to save the same note to SCC and prepare a PowerSchool log for the selected student. The call is usually with a parent at the start of the semester. Successful calls save to SCC Notes; unsuccessful calls save to the next available Attempt column. The PowerSchool helper never clicks Submit: review the log and submit it yourself. Refresh the Sheet after updating Apps Script and the extension.

## PowerSchool selections in the Settings tab

On **Instructions and Settings**, use the **PowerSchool contact log settings** table in A34:G38. The four rows are **SCC Success**, **SCC Attempt**, **ECC Conversation**, and **ECC Attempt**. Enter the exact PowerSchool option values in **Log Type value** (B) and **Subtype value** (D); C and E are human-readable labels. An optional **Additional dropdowns (JSON)** cell (F) accepts an array such as `[{"name":"result","value":"no_answer"}]`. Leave B and D both blank to select Type/Subtype manually on that PowerSchool log. The known ECC values are prefilled; SCC values need to be captured from your school's form.

Each handoff reads the current Settings cells and sends the matching row's values to the extension. Editing these cells changes the next handoff without reinstalling the extension. The temporary extension capture tool can copy Type/Subtype values for pasting into columns B:E. Do not put student numbers or notes in this settings table. The ECC date controls will be handled in extension code once their actual PowerSchool controls are captured.

## ECC notes in the final roster

Use the ECC tab directly. The new note columns are **Overall (E), Classes (F), Grades (G), and Socially (H)**. The script finds them by normalized header text, including the line breaks currently in the headers. **ECC Date (D)** and at least one note field are required to log an entry. The prior history in **Old ECC Dates and Notes (I)** stays visible while you write.

Run **Teacher Tools > Set Up / Repair ECC Batch Columns** once. It adds only **ECC Type** and **Ready to Log** after the existing populated columns; it does not move E:I, Attempt 1/2, or other roster fields. ECC Type defaults to Conversation, or you can choose Attempt. An Attempt needs a free Attempt 1/2 cell.

- To archive notes for several students, check **Ready to Log** on each row and use **Log Ready ECC Rows**. This archives dated notes in I. It does not open or submit PowerSchool logs for those students.
- To archive one row and open its PowerSchool ECC log, select a cell on that ECC row and use **Log Current ECC Row & Open PowerSchool**. Review the prepared log and submit it yourself.
- Recent notes in E:H and the date in D remain visible after either action. Identical dated entries are not appended to history twice, and you can retry a PowerSchool handoff for an already archived entry. Ready checkboxes are unchecked only when their row was handled successfully.
- If the action fails, open the **Automation Log** tab to review the recorded error. Do not share a handoff marker outside your authorized student-data systems.

Run the ECC checks from the repository root with `node --test tests/ecc.test.cjs`.
