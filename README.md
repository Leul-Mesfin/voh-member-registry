# VOICE OF HERs — Member Registry

Google Apps Script automation that turns a Google Form submission into a member
record: it issues a member ID, generates a formatted Word profile document,
stores it in Drive, updates the responses spreadsheet, and notifies the HR team.

Built for the internal records of VOICE OF HERs.

---

## What it does

When a member submits the registration form, the script:

1. Issues a member ID based on the person's team
2. Builds a profile document laid out to the organization's template
3. Saves it as a Word `.docx` in a shared Drive folder
4. Writes the ID, photo preview, and document links into the responses sheet
5. Emails the HR team with the document attached

Repeat submissions are recognised by phone number: the member keeps their
original ID and the earlier document is replaced rather than duplicated.

---

## Member ID scheme

```
VOH + two letters of the team + team code + sequence number
```

Board members carry a `-B` suffix. Board membership is recorded as a flag
rather than a team, since board members also serve on a regular team.

| Team | Letters | Code | First ID |
|---|---|---|---|
| Legal | le | 100 | VOHle101 |
| Secretary | se | 200 | VOHse201 |
| HR | hr | 300 | VOHhr301 |
| IT | it | 400 | VOHit401 |
| Partnership | pa | 500 | VOHpa501 |
| Sponsorship | sp | 600 | VOHsp601 |
| Communication | co | 700 | VOHco701 |
| Accounting | ac | 800 | VOHac801 |
| Event and Training | ev | 900 | VOHev901 |
| Logistics | lo | 1000 | VOHlo1001 |
| Graphics | gr | 2000 | VOHgr2001 |
| Project Management | pr | 3000 | VOHpr3001 |
| Social Media | so | 4000 | VOHso4001 |

Counters are stored in Script Properties, not in the spreadsheet, so deleting a
row never recycles a number.

---

## Files

| File | Purpose |
|---|---|
| `Forms.gs` | Builds and maintains the registration form |
| `Profile.gs` | Generates documents, issues IDs, updates the sheet, emails HR |

Both files live in a single Apps Script project and share one global scope.

---

## Form structure

The form branches on two yes/no questions rather than a single status question,
so that a member who both works and studies answers all four fields without the
form duplicating any page:

```
Page 1  personal details, team, board flag, joined date
        "Are you currently working?"
          Yes -> Work Information -> continues to Education
          No  -> Education

Education
        "Are you currently studying?"
          Yes -> Study Information -> continues to Additional Information
          No  -> Additional Information

Additional Information
        about, emergency contact, documents
```

No `setGoToPage` calls are used. That method controls the page *before* the page
break it is set on, which is a common source of off-by-one branching bugs.

---

## Setup

1. Create a new Apps Script project at [script.google.com](https://script.google.com)
2. Add two script files named `Forms` and `Profile`, and paste in the
   corresponding source
3. Upload the organization logo to Drive and put its file ID in
   `VOH_LOGO_FILE_ID` in `Forms.gs` and `VOH_FALLBACK_LOGO_ID` in `Profile.gs`
4. Set `VOH_HR_EMAIL` to the address that should receive submissions
5. Run `createAndBrandForm` and authorize when prompted
6. Add the two file upload questions by hand on the last page — Apps Script
   cannot create them on personal Google accounts:
   - `Passport Size Photo`
   - `National Id / Kebele Id`
7. Run `vohLockDownDuplicates` to limit the form to one response per account

The form, spreadsheet and folder IDs are saved to Script Properties, so nothing
needs to be pasted between files.

---

## Function reference

### Setup and maintenance

| Function | Purpose |
|---|---|
| `createAndBrandForm` | Builds the form, sheet, folder, and trigger. Run once |
| `showMyLinks` | Prints the form, sheet, and folder links |
| `vohAddDepartmentToForm` | Adds a new team to the live form's dropdown |
| `vohAddAgeQuestion` | Adds the Age question to an existing form |
| `vohLockDownDuplicates` | Disables response editing, limits to one per account |
| `vohStartFresh` | Disconnects the current form and resets ID numbering |
| `vohResetIdNumbers` | Resets team counters only |
| `vohClearMemberIndex` | Wipes the phone-to-ID index |

### Diagnostics

| Function | Purpose |
|---|---|
| `vohListFormQuestions` | Prints every page and question title as saved |
| `vohDebugLastResponse` | Prints every answer of the most recent submission |
| `vohCheckUploads` | Shows which upload matched which slot, and whether it can embed |
| `vohCheckLogo` | Confirms the logo file is readable |
| `vohShowMemberIndex` | Lists the phone-to-ID index |
| `vohTestProfileDoc` | Builds a sample document without a form response |

### Corrections

| Function | Purpose |
|---|---|
| `vohRebuildLastResponse` | Regenerates the last submission's document, same ID |
| `vohReassignMember` | Reissues a member's ID after they change team |
| `vohSetStartingNumbers` | Continues numbering from IDs issued by hand |

---

## Known constraints

**Google account required.** File upload questions force respondents to sign in.
Members without a Google account cannot submit; HR collects their photo and ID
separately.

**File upload questions cannot be scripted** on personal Google accounts. They
are added through the form editor.

**Image formats.** Google Docs cannot embed HEIC files from iPhones, or PDFs.
The script inserts a clickable link instead of failing silently.

**Email quota.** Personal Google accounts allow roughly 100 script emails per
day. Documents are still created if the notification fails.

**Upload matching.** The photo and ID are matched by keyword in the question
title. The ID is checked first, on `kebele`, `national`, `identity`, so a title
containing both "passport" and "national" still resolves correctly.

---

## Privacy

The form collects passport photos and National ID scans. Access to the form, the
responses spreadsheet, the profiles folder, and the Forms uploads folder should
be limited to the HR team, matching the confidentiality notice shown to members.

Before publishing this repository, replace real email addresses and file IDs
with placeholders, or keep the repository private.
