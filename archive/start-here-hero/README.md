# Archive: "Start Here Hero"

A saved stopping point, per request (2026-07-14). This is the workbook design
**before** the MapPages restructure that moved the job info onto the MapPages
sheet.

- **Git tag:** `start-here-hero` -> commit `74b1d43`
  (`git checkout start-here-hero` to get the full source tree of this version).
- **These two `.xlsm` files** are the exact built workbooks from that commit,
  extracted with `git show` (byte-for-byte: 403145 / 404672 bytes). Both verified
  to open in Excel.

## What defines this version

- **Start Here is the hub.** On the Site Inspector product, the job info
  (WO / DI / Disaster / Applicant / Output Folder) lives on **Start Here**, and
  every workflow launches from there.
- Three sheets only: **Start Here / Sites / Sources**. MapPages is *disposable* -
  it's created by "Prepare Map Pages" and deleted/recreated on each run.
- Map-page stamps are baked at page-creation time (no "Update Stamps" concept).
- RoadReviewer's export dropdown already includes the map-page workflow items.

## To restore this version later

```powershell
# inspect / branch from it
git checkout start-here-hero            # detached HEAD at 74b1d43
git switch -c back-to-hero              # or start a branch here

# OR just drop these built files back at the repo root
Copy-Item "archive\start-here-hero\RoadReviewer.xlsm" .
Copy-Item "archive\start-here-hero\Site Inspector Review Tool.xlsm" .
```
