# Outlook Email Templater

A Windows (WPF, .NET 8) app for organizing "file → person" email sends through classic Outlook.

## What it does

You define any number of **groups**. Each group has:
- Recipients (name + email)
- Attachments (drag-and-drop or browse)
- Subject + body with `{{token}}` variables
- Optional group-local variables
- A send mode (use global default, open draft in Outlook, or send automatically)

Hit **Send via Outlook** and Cavt Email will either pop the draft open in classic Outlook for review, or send it automatically (with a confirmation prompt).

## Requirements

- Windows
- .NET 8 SDK to build
- **Classic Outlook (desktop)** installed for sending. "New Outlook" / web Outlook is not COM-automatable.

## Build & run

```powershell
dotnet build
dotnet run --project CavtEmail
```

## Publish (for distribution to coworkers)

Multi-file self-contained Windows x64 build. **Not** single-file on purpose — single-file
bundles extract to a temp folder on first launch on every new machine, which causes a
noticeable delay. A plain folder launches instantly.

```powershell
.\scripts\publish.ps1
```

Output:
- Folder: `publish\CavtEmail-win-x64\` — ship the whole folder; `CavtEmail.exe` is the entry point.
- Zip:    `publish\CavtEmail-win-x64.zip` — same content, for handing out.

Pass `-NoZip` to skip the zip step.

## Variable tokens

Put these anywhere in the **Subject** or **Body**:

| Token | Value |
| --- | --- |
| `{{recipients}}` | comma-separated names (falls back to email) |
| `{{recipientEmails}}` | comma-separated email addresses |
| `{{recipientCount}}` | number of recipients |
| `{{files}}` | newline list of attached file names |
| `{{fileList}}` | bulleted list (` - name`) |
| `{{fileCount}}` | number of attached files |
| `{{sender}}` | global sender name |
| `{{date}}`, `{{time}}` | today's date / current time |
| `{{group}}` | the group's name |

## Quick actions in the editor

- Click a **token chip** on the right panel to insert it at the body or subject cursor (toggle at top of panel).
- Each recipient / attachment row has **Subj** and **Body** buttons that insert the person's name or the filename into that field.
- **Drag files** (or a folder) onto the Attachments panel to attach them.

## Persistence

Your groups, recipients, attachments, and variables are saved as JSON. The default location is:

```
%APPDATA%\CavtEmail\config.json
```

Use **Save As…** to keep named configurations (e.g., one per workflow), and **Open…** to load one.

## Sending

All recipients in a group go into the single email's `To:` field. Send modes:

- **OpenDraft** — builds the mail in Outlook and opens it for review. Default and safest.
- **SendAutomatically** — sends without opening the window (asks for confirmation first).
- **UseGlobal** (per-group) — defer to the global default shown in the toolbar.

## Notes / Security

- Attachments are added by absolute path; if a file has been moved/deleted at send time you'll be warned before continuing.
- Auto-send mode requires classic Outlook. Depending on your Outlook / antivirus configuration, programmatic `Send()` may trigger a security prompt.
- The app never contacts any server — everything happens locally via Outlook COM.
