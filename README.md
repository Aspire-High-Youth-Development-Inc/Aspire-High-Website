# Aspire High Youth Development Inc. — website

The public site for Aspire High Youth Development Inc., in a single
self-contained file. No build step, no dependencies, no server.

**Live site:** https://aspire-high-youth-development-inc.github.io/Aspire-High-Website/#home

---

## Uploading from scratch

The files in this folder go at the **top level** of the repository — not inside
a folder. If `index.html` ends up in a subfolder the site address changes and
the link above stops working.

### If the page still looks old

GitHub Pages caches hard, and so does the browser. Reload with
**Ctrl + Shift + R** (**⌘ ⇧ R** on a Mac), or open the link in a private
window to confirm the new version is live.

## Switching on the student application (do this first)

The application — English and Spanish — is built into `index.html`; there
are no separate pages. Both languages send every submission to one Google Sheet.
`#apply` links straight to the English form and `#aplicar` to the Spanish one.
Until the steps below are done the forms load and check answers, but tell the
family that online submission is not open yet and give the phone and email
instead — nothing a parent types is silently lost.

1. **Make the sheet.** In the organization's Google Drive: **New → File
   upload** → `Aspire-High-Applications-2026-27.xlsx`. Right-click it →
   **Open with → Google Sheets** → **File → Save as Google Sheets**. Leave its
   sharing on **Restricted** and add only the staff who need it — it holds
   children's personal details.
2. **Add the script.** In that Google Sheet: **Extensions → Apps Script**.
   Delete what is there, paste in all of `Code.gs`, click **Save**. Under
   **Project Settings** set the time zone to **New York**.
3. **Publish it.** **Deploy → New deployment →** type **Web app**.
   *Execute as:* **Me**. *Who has access:* **Anyone**. Click **Deploy** and
   approve the permissions — Google warns that it has not verified the app;
   it is your own script, so choose **Advanced → Go to … (unsafe)**. Copy the
   **Web app URL**; it ends in `/exec`.
   "Anyone" lets people *send* an application. The script never sends any
   data back, and the sheet stays private.
4. **Connect the site.** In `index.html`, find `applicationEndpoint: ""`
   near the top of the script and paste the URL between the quotes (on GitHub:
   open the file → pencil icon → Ctrl+F `applicationEndpoint` → **Commit
   changes**). That one line switches on both languages.
5. **Test it.** Submit a test application, check the row appears in the
   sheet, then delete that row.
6. **Show it to staff.** Put the sheet's link in `applicationSheet` near the
   top of the script in `index.html` (or send it to whoever maintains the
   site). The staff Applications section then shows the live table and a
   **Download for Excel** button.

Optional: set `NOTIFY_EMAIL` near the top of `Code.gs` to be emailed when an
application arrives. The email says only that one arrived — no student details
travel by email. After any change to `Code.gs`: **Deploy → Manage deployments →
edit → Version: New version**.

**How answers are stored.** Both pages send the same field names and the same
English answer values, so a Spanish form's "Almuerzo gratuito" arrives as
"Free lunch" and every answer lands in the same column. The **Form language**
column records which page the family used. ZIP codes and phone numbers are
kept as text, so New Jersey's leading zero is never lost. Anything a visitor
types that starts with `=`, `+`, `-` or `@` is stored as plain text, never as
a formula.

**The workbook's other tabs.** **Summary** counts applications, grades,
first-generation students, free or reduced lunch, English learners,
race/ethnicity, tutoring and parent-session interest — the figures grant
reports ask for. It also counts **Photo/video NOT authorised**: filter
Responses on that column before any photo or video is taken or posted.
**Field guide** explains every column.

---

## The flyer's QR code

The flyer on the site now points at `https://pateldixit28603.github.io/aspire-high/#apply`,
which opens the site on the application with the English form showing.
Upload these files **before** printing it — until then the code leads to a
"page not found".

The original code on the flyer is a **Flowcode**, which can be pointed
somewhere new from the Flowcode dashboard without reprinting anything. If
flyers with that code are already out, change its destination to the address
above so the printed copies lead to the new application too.

---

## What is in this build

This is the **public** build. It comes from the master file through
`build-public.py` (in the server bundle), which strips out the values that
point at student data — GitHub Pages serves static files and has no server to
hide them:

- Drive folder ids for the workshop and response folders
- The attendance spreadsheet link
- The application responses sheet
- The HR Center address and its forms library

The staff sections still render as a demonstration, but hold no real links.

### The sign-in here is a demonstration

Sign in on this build and the staff sections appear, but that check runs **in
the browser**. Anyone who opens developer tools can get past it. That is fine
for a public showcase, and **not** fine for student records, timesheets or the
time clock. For those, run the Flask build (`aspire-high-server`), where staff
content is removed on the server before the page is sent and the Airtable key
never reaches the browser. That build needs a host that runs Python, such as
PythonAnywhere or Render.

---

## What the site contains

**Public**

- Home, About, Programs, Get involved
- Program calendar — the full 2026–27 Strategic Calendar: After School
  workshops, Parent Engagement, Life Skills, key dates and closures
- Program details, workshop topics and dates, key dates with legend, and goals
- After School workshop forms — Workshop 1 and Workshop 2, listed live from
  Drive, so a form staff add appears without anyone touching the site
- HR — joining, onboarding, training, reference requests, contact, and the
  organization's trailer video
- Training, Documents, FAQ, Support us, Contact
- English and Spanish throughout, switched from the header

**Staff — real version is the Flask build**

- Clock in / clock out, with CSV export that opens in Excel
- Task tracker
- HR Center forms library and the HR site itself
- Applications, staff files, directory, attendance, GPA
- Looker Studio dashboards, Power BI, Airtable tables
- Funder summary, data entry, and the Manage page

## Editing content

Sign in as staff and use **Manage**: programs, documents, dashboards, calendar
events, training videos, the After School form links, contact details and the
FAQ can all be changed there without touching code.

Manage changes are stored in that person's own browser. To change what every
visitor sees, edit the `CONTENT` object near the top of the script in
`index.html` and upload the file again.

---

## Still to do outside the site

These have to be done in Google or Airtable — the website cannot do them:

0. **The student application** — see *Switching on the student application*
   above. Until it is done, families cannot apply online.
1. **Looker Studio sharing.** All five reports currently redirect to a Google
   sign-in, so the dashboards show a login box instead of a report. In each
   report: **Share → Anyone with the link → Viewer**, then **Share → Embed
   report → Enable embedding**. Both are needed; one alone will not work.
2. **Airtable tables.** Create `Time clock` with fields `name`, `email`,
   `date`, `clock_in`, `clock_out`, `hours`, and `Tasks` with `title`,
   `owner`, `due`, `status` — names matching exactly. Until then the clock and
   the tracker still work in the browser and export to CSV.
3. **Drive ownership.** The workshop forms are owned by someone who has left
   the organization. Transfer ownership to current staff, or access to the
   forms and their responses is at risk if that account is closed.
4. **Grade-level form links.** The grade cards describe each year but have no
   form of their own yet. Add each grade's Google Form link under
   **Manage → After School sign-up by grade**.
5. **Dashboard names.** Three reports are labelled `Looker report 3/4/5`.
   Rename them under **Manage → Looker Studio dashboards**.

## Two date clashes in the Strategic Calendar

Both are in the calendar exactly as supplied, each carrying a warning note:

- Life skills **Time Management**, 30 December — falls inside Winter Break
- Life skills **Career Readiness**, 31 March — falls inside Spring Break

Decide whether to move them or run them through the break.

## Credits

Built by the Data Analytics and IT teams at Aspire High Youth Development Inc.
