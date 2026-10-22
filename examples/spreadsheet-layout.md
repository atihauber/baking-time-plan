# Spreadsheet layout example

Create a sheet named `Kenyér` with the following headers in row 1:

| Fázis | Késleltetés (min) | Időpont megadása / felülírás | Számolt időpont |
| --- | ---: | --- | --- |
| Autolízis | 0 | | |
| Dagasztás | 60 | 15:15 | |
| Dagasztás sóval | 30 | | |
| Laminálás | 45 | | |
| Hajtogatás 1. | 45 | | |

The time `15:15` is an example override. With the delays above, the calculated times from that row onward would be 15:15, 15:45, 16:30, and 17:15.

Create a second sheet named `Summary`. The script populates its four output columns automatically.

Important: keep phase rows in the intended processing order. The delay on a row is added to the previous calculated phase time. Rows before the first override remain without calculated times.
