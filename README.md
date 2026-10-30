# EA_outlook_addin

Outlook add-in "Naar Notion". In a mail (read or compose), click **Naar Notion** in the ribbon, pick a lane (Vandaag, Volgende dag, Deze week, Backlog) and optionally type a note. The add-in:

- creates the master categories once if missing: `Notion: Sprint 1 day` (orange), `Notion: Sprint next day` (yellow), `Notion: Sprint 1 week` (green), `Notion: Backlog` (gray);
- puts that category on the mail (only one Notion lane at a time);
- stores the note as custom property `evaNote` (plus `evaLane`), best effort;
- "Ongedaan maken" removes the category and the note.

Eva later finds mails with a `Notion: *` category via Graph and creates the Notion cards.

Static HTML/JS/CSS plus Office.js. No secrets, no backend.

## Publish (one time)

1. On github.com create an **empty public** repo `EA_outlook_addin` under `leon-hartkoorn` (no README, no license).
2. From this folder:
   ```
   git remote add origin https://github.com/leon-hartkoorn/EA_outlook_addin.git
   git branch -M main
   git push -u origin main
   ```
3. Repo Settings > Pages > Source "Deploy from a branch", branch `main`, folder `/ (root)`, Save. After a minute, check https://leon-hartkoorn.github.io/EA_outlook_addin/taskpane.html loads.

## Install in Outlook

1. Open https://aka.ms/olksideload (Outlook on the web, "Add-ins for Outlook" dialog) > My add-ins.
2. Under Custom add-ins, remove the old test add-in.
3. "Add a custom add-in" > "Add from file" > choose `manifest.xml` from this folder.
4. Open a mail: the **Naar Notion** button sits in the ribbon (in the new Outlook / web it may be under the "Apps" button or the `...` menu; pin it via Settings > Mail > Customize actions).

New Outlook for Windows and classic Outlook pick the add-in up automatically for the same account (classic may need a restart).

## For Eva: reading the note via Graph

Office.js custom properties live in one extended property holding JSON:

```
GET /me/messages?$filter=categories/any(c:c eq 'Notion: Sprint 1 day')
    &$expand=singleValueExtendedProperties($filter=id eq 'String {00020329-0002-0000-C000-000000000046} Name cecp-2c2658f7-4faa-450c-9387-4f9ccf079dbb')
```

The value is a JSON string like `{"evaNote":"...","evaLane":"Notion: Sprint 1 day"}`.

## Updating

Change files, commit, push: Pages redeploys. Only re-add the manifest when `manifest.xml` changes (bump `<Version>`).
