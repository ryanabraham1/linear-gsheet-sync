# Machining Tracker → Linear setup

This script is preconfigured for:

- Google Sheets tab: **Machining Tracker**
- Linear team: **WarriorBorgs 3256** (`WB`)
- Linear project: **Dumper**
- Linear label: **Fabrication** (the workspace has no separate label named `Fab`)
- Linear milestone: **Dumper Fabbed**

It preserves the existing tracker in columns A–N. Setup adds sync metadata in columns O–T, then hides the internal **Linear ID** and **_Sync Hash** columns.

## Script versions

- `Code.local.gs` is the local source of truth. It contains the testing API key and is ignored by Git.
- `Code.gs` is safe to publish. It contains no key and asks the installer for one during setup.

After changing `Code.local.gs`, run `npm run sync-public`. This regenerates `Code.gs` with the key removed. Run `npm run check` before committing; it fails if the two versions differ anywhere except the key.

## Install

1. Open the [Hardware Resources Google Sheet](https://docs.google.com/spreadsheets/d/1WlHAodErTZT3_4QGyWwXvK8GIcZfhQbtOCWkGZ7PArk/edit?gid=0#gid=0).
2. Choose **Extensions → Apps Script**.
3. Delete the sample function and paste all of `Code.local.gs` for this private test deployment. Anyone installing the public GitHub version should paste `Code.gs` and enter their Linear key when prompted.
4. Click **Save**. Do not use **Deploy → New deployment**; this is a Sheet-bound script.
5. Return to the Sheet and reload it.
6. Choose **Linear Sync → Install / reconnect**.
7. Approve the Google authorization prompts. If Google shows an unverified-app warning for your own script, confirm that this is your script, then choose **Advanced → Go to project → Allow**.

That one command:

- verifies the Linear team, project, and label;
- creates the tracking columns;
- installs an immediate edit trigger;
- installs a five-minute reconciliation trigger; and
- uploads all existing rows that have a **Part #_Name**.

## Field mapping

| Machining Tracker | Linear |
|---|---|
| Part #_Name | Issue title |
| Status: Not Started | Todo |
| Status: In Progress | In Progress |
| Status: Finished | Done |
| Priority: #1–#4 | Urgent–Low |
| Subsystem, quantities, stock, dimensions, length, tapped, machine, drawing, notes | Markdown issue description |
| — | Project: Dumper |
| — | Label: Fabrication |
| — | Milestone: Dumper Fabbed |

Every Linear description includes a link back to its source row.

Every Linear issue title is prefixed with **`Fab: `**. For example, `0100_Drivebase Studio_Bellypan` becomes `Fab: 0100_Drivebase Studio_Bellypan`.

## Deletions and restores

- Clearing a row's **Part #_Name** archives its linked Linear issue and leaves the hidden Linear ID in place.
- Restoring **Part #_Name** unarchives and updates that same issue instead of creating a duplicate.
- Physically deleting a spreadsheet row is detected through an installable change trigger. The issue ID registry is stored in the Sheet's document properties so the removed issue can still be identified and archived.
- The script uses Linear's trash/archive operation rather than permanent deletion, so an accidental Sheet deletion remains recoverable.

## Confirm it worked

1. The populated rows should receive a **Linear Key**, **Linear URL**, and **Last Synced** value.
2. Open one Linear URL and verify its project is **Dumper**, its label is **Fabrication**, and its milestone is **Dumper Fabbed**.
3. Change a tracker field. The existing Linear issue should update within a few seconds instead of creating another issue.
4. If a row reports an error, fix the value and choose **Linear Sync → Sync selected rows now**.

## Notes

- Rows without **Part #_Name** are ignored, including section or spacer rows.
- Clearing **Part #_Name** or physically deleting a tracked row archives its Linear issue.
- Updates made directly in Linear do not flow back into Sheets.
- The automation runs as the Google account that installs the triggers.
- To disable it, choose **Linear Sync → Remove automatic triggers**.
