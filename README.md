# EA_outlook_addin

Outlook add-in "Eva: Naar Notion". In a mail, click **Naar Notion** in the ribbon (pinnable; the pane follows the selected mail). The pane shows, top to bottom:

1. **Eva triage card** (read mode only): Eva's summary of this mail, a context line, one recommended action (primary button, with reason) and alternatives as chips. Without a triage it shows "Nog geen triage van Eva voor deze mail".
2. **Lane chips** Sprint 1D, Next day, Sprint 1W, Backlog, On hold: puts category `Notion: Sprint 1 day` / `Notion: Sprint next day` / `Notion: Sprint 1 week` / `Notion: Backlog` / `Notion: On hold/waiting for` on the mail (one lane at a time; master categories created once if missing).
3. **Note for Eva**, stored as custom property `evaNote` (plus `evaLane`), autosaved.

Layout is compact and works at 250-300px pane width.

## Eva triage: data contract

All values live in the add-in's custom properties, which Outlook stores as ONE Graph extended property holding a JSON object of string values:

```
String {00020329-0002-0000-C000-000000000046} Name cecp-2c2658f7-4faa-450c-9387-4f9ccf079dbb
```

- `evaTriage` (written by Eva via `04_Assistant/tools/mail_triage.py write`), a JSON string:
  ```
  {"v":1,"at":"2026-10-30T13:05:00+00:00","summary":"...","context":"OoakT, Notion: ..., WO 28/10",
   "recommend":{"type":"archive","folderPath":"Portfolio Companies/OoaKTravel_2023","folderLabel":"OoakT","label":"Naar map OoakT","reason":"..."},
   "alternatives":[{"type":"notion","lane":"Sprint 1 week","cardTitle":"..."},{"type":"fyi"},{"type":"draft","gist":"..."}]}
  ```
  Action types: `archive` {folderPath, folderLabel}, `notion` {lane: key/label/Notion status, cardTitle}, `draft` {gist}, `fyi`, `none`. `label` is optional (a Dutch default is used).
- `evaAction` (written by the pane on click), a JSON string `{type, params, label, at}`, together with category `Eva: actie` (purple, Preset8). Eva's poller executes it (`mail_triage.py pending` / `move` / `clear`) and removes both.

Click behaviour:
- `archive`: first tries a direct EWS move (FindFolder by path under msgfolderroot, UpdateItem IsRead, MoveItem). Works in classic Outlook and Outlook on the web while EWS is available. Shows "Gearchiveerd in ..." with undo (moves back to Postvak IN, restores unread). If EWS is unavailable or fails (new Outlook, EWS retired), it records `evaAction` and shows "Eva verplaatst hem zo".
- `fyi`: EWS mark read (undo restores), else recorded.
- `draft`: always recorded; Eva writes the draft with `tools/reply_draft.py`.
- `notion`: same as clicking that lane, note prefilled with `cardTitle`.
- `none`: nothing is stored.
- Recorded actions can be undone in the pane (removes `evaAction` and `Eva: actie`).

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

The value is a JSON string like `{"evaNote":"...","evaLane":"Notion: Sprint 1 day","evaTriage":"{...}"}`.

## Updating

Change files, commit, push: Pages redeploys. Only re-add the manifest when `manifest.xml` changes (bump `<Version>`).
