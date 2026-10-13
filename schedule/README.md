# Neumann Residence – Construction Schedule

1567 Steveles Crescent, Mississauga. Status date: Oct 13, 2026.

| File | Purpose |
|---|---|
| `Neumann_Residence_Schedule.xml` | Microsoft Project schedule (MSPDI XML). 175 lines, 15 phases, 226 logic links, baselined. |
| `Neumann_Residence_Schedule_Preview.pdf` | Gantt preview, 11x17, for reading without MS Project. |
| `build_schedule.py` | Source of the schedule: tasks, durations, logic, notes. Edit here and rebuild. |
| `render_gantt.py` | Regenerates the PDF preview from the XML. |
| `schedule_dates.txt` | Computed start and finish for every line. |

## Converting to .mpp

1. In MS Project, go to **File → Open** and select `Neumann_Residence_Schedule.xml`. Choose "Open as new project" if asked.
2. Go to **File → Save As** and save as **Project (*.mpp)**.

The calendar ("Chatsworth Standard": Mon–Fri, Ontario stat holidays, Dec 24–Jan 1 shutdown), the logic links, deadlines, the baseline, task notes and the **Responsible** field (Text1) all carry over.

## Key assumptions

- Site possession is Oct 15, 2026. This still needs to be confirmed: Stephen's Oct 8 email mentions Nov 15.
- The building permit application goes in Oct 29. City review is assumed at 25 working days, which puts issuance on Dec 3. This is the main driver of the schedule; once the City confirms its review time, change the `permit_rev` duration.
- Window, door and skylight lead time is 10 weeks from PO.
- Cabinetry design approval starts Jan 4 and is followed by 15 days of shop drawings and 45 days of fabrication. This chain is on the critical path.
- Rough-ins start in the existing house in December, after the interior structural alterations. The rear addition gets its own rough-in once it is watertight.
- Exterior cladding work is held to April 5 or later.

## Results against the Oct 8 targets

| Target | Scheduled | Status |
|---|---|---|
| Interior demo complete: mid-Nov | Nov 19, 2026 | On target |
| Watertight: mid-Jan | Mar 8, 2027 | **About 7 weeks late.** Foundations can't start until the permit is issued (Dec 3). The existing house is weathertight by Feb 4 once its windows are in. |
| Rough-in inspections: early Feb | Jan 25, 2027 (existing house) | On target |
| Insulation & drywall: April | Insulation from Jan 26 (existing house); drywall Mar 31 – Apr 27, 2027 | On target |
| Occupancy: Aug 30, 2027 | Jul 29, 2027 | About 4 weeks of float |
| Client move-in: Sep 1, 2027 | Aug 3, 2027 | On target |

Budget Rev 4 adds the loft / 3rd floor back into scope and is expected to use some of that float.
