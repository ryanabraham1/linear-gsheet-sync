# Machining Tracker → Linear setup

This script is preconfigured for:

- Google Sheets tab: **Machining Tracker**
- Linear team: **WarriorBorgs 3256** (`WB`)
- Linear projects selected by **Bot**: **Aimbot Changes**, **Dumper**, or **EveryBot**
- Linear label: **Fabrication** (the workspace has no separate label named `Fab`)
- Linear fabrication milestone matching the selected bot project

It preserves the existing tracker in columns A–O. Setup adds sync metadata in columns P–U, then hides the internal **Linear ID** and **_Sync Hash** columns.

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

- verifies the Linear team, all three bot projects and fabrication milestones, and the label;
- creates the tracking columns;
- installs an immediate edit trigger;
- installs a five-minute reconciliation trigger; and
- uploads all existing rows that have a **Part #_Name**.

## Field mapping

| Machining Tracker | Linear |
|---|---|
| Part #_Name | Issue title |
| Bot: Aimbot | Project: Aimbot Changes; milestone: Fully Fabbed |
| Bot: Dumper | Project: Dumper; milestone: Dumper Fabbed |
| Bot: EveryBot | Project: EveryBot; milestone: Fully Fabbed |
| Bot: blank | No project or project milestone |
| Status: Not Started | Todo |
| Status: In Progress | In Progress |
| Status: Finished | Done |
| Priority: #1–#4 | Urgent–Low |
| Subsystem, quantities, stock, dimensions, length, tapped, machine, drawing, notes | Markdown issue description |
| — | Label: Fabrication |

Every Linear description includes a link back to its source row.

Every Linear issue title is prefixed with **`Fab: `**. For example, `0100_Drivebase Studio_Bellypan` becomes `Fab: 0100_Drivebase Studio_Bellypan`.

Changing **Bot** moves the existing issue to the matching project and fabrication milestone. Clearing **Bot** removes both the project and project milestone without deleting or archiving the issue.

## Dynamic label groups

The script creates and maintains two team label groups in Linear:

- **Subsystem** — the selected Subsystem becomes a child label. Numeric ordering prefixes are removed, so `01. Drivebase` becomes the `Drivebase` label.
- **Fab Machine** — the selected Machine becomes a child label, such as `Bridgeport Mill` or `Lathe`.

When either selection changes, the old child label from that group is removed and the new child label is added. The permanent **Fabrication** label and unrelated labels added manually in Linear are preserved. Subsystem and Machine also remain visible in the issue description.

## Deletions and restores

- Clearing a row's **Part #_Name** archives its linked Linear issue and leaves the hidden Linear ID in place.
- Restoring **Part #_Name** unarchives and updates that same issue instead of creating a duplicate.
- Physically deleting a spreadsheet row is detected through an installable change trigger. The issue ID registry is stored in the Sheet's document properties so the removed issue can still be identified and archived.
- The script uses Linear's trash/archive operation rather than permanent deletion, so an accidental Sheet deletion remains recoverable.

## Confirm it worked

1. The populated rows should receive a **Linear Key**, **Linear URL**, and **Last Synced** value.
2. Open one Linear URL and verify its project and fabrication milestone match **Bot**, or that both are blank when **Bot** is blank. Its label should remain **Fabrication**.
3. Change a tracker field. The existing Linear issue should update within a few seconds instead of creating another issue.
4. If a row reports an error, fix the value and choose **Linear Sync → Sync selected rows now**.

## Notes

- Rows without **Part #_Name** are ignored, including section or spacer rows.
- Clearing **Part #_Name** or physically deleting a tracked row archives its Linear issue.
- Updates made directly in Linear do not flow back into Sheets.
- The automation runs as the Google account that installs the triggers.
- To disable it, choose **Linear Sync → Remove automatic triggers**.
