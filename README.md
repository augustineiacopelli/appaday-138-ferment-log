# AppADay 138: Ferment Log

Ferment Log is a small brewery tracker for homebrewers that keeps every record in a Google Sheet you own. It follows a beer from the ingredients on the shelf, through planning and brew day, gravity readings in the fermenter, time in the tank, and into kegs, bottles, or cans until the last one is gone. The web app is a single HTML file; a short Apps Script bound to your sheet acts as its backend.

Part of [AppADay](https://augustineiacopelli.github.io/appaday/), one complete web app shipped every day.

## Setup

1. Create a blank Google Sheet.
2. Open Extensions, then Apps Script.
3. Replace the contents of `Code.gs` with the `Code.gs` from this repo and save.
4. Choose `setup` in the function menu, click Run, and authorize the script when asked. It creates the seven tabs with headers and generates an access token.
5. Open the execution log and copy the line that begins with `TOKEN:`.
6. Click Deploy, then New deployment. Choose Web app, set Execute as to Me and Who has access to Anyone, then click Deploy.
7. Copy the Web App URL that ends in `/exec`.
8. Open the app, tap the gear, paste the URL and token, tap Test connection, then Save.

If you set up an earlier version, paste the new `Code.gs` and run `setup` again. It adds the Par and Status columns, moves any Conditioning batches to Crashing, and marks existing tanks clean without touching your other data.

When you change `Code.gs` later, publish it through Deploy, then Manage deployments, edit the existing deployment, and choose New version. The URL stays the same, so the app keeps working without new settings.

The token and URL are stored only in your browser. Anyone who has both can read and change the sheet, so keep them private.

## Features

Inventory groups grain, hops, yeast, and adjuncts, flags anything at or below its reorder point, and adjusts stock with plus and minus buttons. Each item can carry a par level, and low items appear in a shopping list showing how much to order to get back to par. Batches are planned with ingredients pulled straight from inventory, and each card shows a seven stage timeline from Planned through Brewed, Fermenting, Crashing, Carbonating, Packaged, and Kicked with the date each stage was reached. Crashing always comes before Carbonating, and only a carbonating batch can be packaged. The Ferment tab logs gravity, temperature, and notes, then charts gravity and temperature on two axes against the target final gravity. Tanks are drawn as conical fermenters that show their state at a glance: fermenting, crashing, or carbonating when full, with the liquid level drawn to the approximate volume, or dirty, clean, or sanitized when empty. A tank must be marked sanitized before a batch can be assigned, and releasing or packaging out of a tank marks it dirty. Packaged records kegs, bottles, and cans from common presets, reports yield and loss, and counts down the units on hand until the batch is marked Kicked automatically.

The app caches the last synced state, so it opens instantly and still shows your data if the network drops.

## Formulas

ABV is (OG minus SG) times 131.25, rounded to one decimal. Apparent attenuation is (OG minus SG) divided by (OG minus 1), times 100, rounded to a whole number. OG is the earliest reading by date, falling back to the batch target, and SG is the latest reading. Days in tank is the whole number of days since the assignment date. Packaged gallons are the sum of unit volume times unit count, yield is packaged gallons divided by batch volume, and loss is batch volume minus packaged gallons.

## Sheet tabs

Inventory, Batches, BatchIngredients, Readings, StageLog, Tanks, and Packages. Dates are stored as YYYY-MM-DD text and IDs use a short prefix such as `INV-3f9a1c2e`.

## Stack

HTML, CSS, and vanilla JavaScript in one file, Chart.js 4.4.1 for the fermentation chart, and Google Apps Script with Google Sheets for storage.
