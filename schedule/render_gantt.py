"""Render a PDF Gantt preview from the MS Project XML schedule.

Usage: python3 render_gantt.py <schedule.xml> <output.pdf>
"""
import sys
from datetime import date, datetime, timedelta

import jpype
import jpype.imports
import mpxj  # noqa: F401

jpype.startJVM()

import matplotlib  # noqa: E402
matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from org.mpxj.reader import UniversalProjectReader  # noqa: E402

ROWS_PER_PAGE = 48
STATUS_DATE = date(2026, 10, 13)
INK = "#2b2b2b"
TASK = "#5b7fa3"
CRIT = "#b4553f"
DONE = "#8c9a7a"
PHASE = "#3d3d3d"
GRID = "#d9d9d9"


def py(dt):
    return None if dt is None else datetime(dt.getYear(), dt.getMonthValue(), dt.getDayOfMonth(),
                                            dt.getHour(), dt.getMinute())


def load(path):
    project = UniversalProjectReader().read(path)
    rows = []
    for t in project.getTasks():
        if t.getID() is None or t.getID() == 0:
            continue
        rows.append({
            "name": str(t.getName()),
            "summary": bool(t.getSummary()),
            "milestone": bool(t.getMilestone()),
            "critical": bool(t.getCritical()) if t.getCritical() is not None else False,
            "done": t.getPercentageComplete() is not None and float(str(t.getPercentageComplete())) >= 100,
            "start": py(t.getStart()),
            "finish": py(t.getFinish()),
            "deadline": py(t.getDeadline()),
            "resp": str(t.getText(1)) if t.getText(1) else "",
        })
    return project, rows


def draw_page(pdf, rows, x0, x1, page, pages, title):
    fig = plt.figure(figsize=(17, 11))
    ax = fig.add_axes([0.30, 0.06, 0.68, 0.86])
    n = len(rows)
    ax.set_ylim(ROWS_PER_PAGE - 0.5, -0.5)
    ax.set_xlim(x0, x1)
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))
    ax.xaxis.tick_top()
    ax.tick_params(axis="x", labelsize=8, colors=INK, length=0)
    ax.set_yticks([])
    for s in ax.spines.values():
        s.set_visible(False)
    ax.grid(axis="x", color=GRID, linewidth=0.6)
    ax.axvline(datetime(STATUS_DATE.year, STATUS_DATE.month, STATUS_DATE.day), color=CRIT, lw=0.8, ls="--")

    for i, r in enumerate(rows):
        y = i
        label_x = -0.44
        if r["summary"]:
            ax.axhspan(y - 0.5, y + 0.5, color="#f1f1f1", zorder=0)
            fig.text(0.01, 0, "")
            ax.text(label_x, y, r["name"], transform=ax.get_yaxis_transform(), ha="left", va="center",
                    fontsize=8, fontweight="bold", color=INK)
            ax.barh(y, r["finish"] - r["start"], left=r["start"], height=0.28, color=PHASE)
            continue
        ax.text(label_x + 0.01, y, r["name"][:70], transform=ax.get_yaxis_transform(), ha="left", va="center",
                fontsize=6.8, color=INK)
        ax.text(-0.005, y, r["resp"][:24], transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=6, color="#6b6b6b")
        colour = DONE if r["done"] else (CRIT if r["critical"] else TASK)
        if r["milestone"]:
            ax.plot(r["start"], y, marker="D", ms=6, color=INK if not r["critical"] else CRIT)
            ax.text(r["start"] + timedelta(days=4), y, r["start"].strftime("%b %d"), va="center", fontsize=6.5, color=INK)
        else:
            width = max(r["finish"] - r["start"], timedelta(hours=12))
            ax.barh(y, width, left=r["start"], height=0.55, color=colour)
        if r["deadline"]:
            ax.plot(r["deadline"], y - 0.42, marker="v", ms=5, color=CRIT)

    fig.text(0.01, 0.975, title, fontsize=12, fontweight="bold", color=INK, va="top")
    fig.text(0.01, 0.955, "Chatsworth Fine Homes  |  Status date Oct 13, 2026  |  Page %d of %d" % (page, pages),
             fontsize=8, color="#6b6b6b", va="top")
    fig.text(0.98, 0.015, "Blue: task   Red: critical path   Green: complete   Diamond: milestone   "
             "Red triangle: target date   Dashed line: status date", fontsize=7, color="#6b6b6b", ha="right")
    pdf.savefig(fig)
    plt.close(fig)


def main(src, out):
    project, rows = load(src)
    x0 = datetime(2026, 9, 25)
    x1 = datetime(2027, 9, 15)
    title = "Neumann Residence  -  1567 Steveles Crescent, Mississauga  -  Construction Schedule"
    pages = (len(rows) + ROWS_PER_PAGE - 1) // ROWS_PER_PAGE
    with PdfPages(out) as pdf:
        for p in range(pages):
            draw_page(pdf, rows[p * ROWS_PER_PAGE:(p + 1) * ROWS_PER_PAGE], x0, x1, p + 1, pages, title)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
