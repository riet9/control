# Screen Time Tracker

Windows desktop tracker for active computer time, focused on study discipline and daily limits.

The app tracks only active time, pauses after idle time, supports exact browser-domain classification, runs in tray, and helps enforce limits for:

- total computer time
- study
- browser fun / manga / manhwa
- social media

Default limits:

- total: `3h`
- study target: `2h` to `2.5h`
- browser fun: `30m`
- socials: `18m`

## Highlights

- Active-time tracking instead of simple PC uptime
- Idle detection
- Rules by process, window title, URL, or domain
- Custom categories mapped to built-in parent categories
- Review screen for uncategorized activity
- Health screen with rule suggestions
- Weekly review and calendar heatmap
- Focus mode and hard block mode
- Tray controls and quick glance mini window
- Per-day storage with summary cache for better performance
- CSV / JSON export

## Tech Stack

- PowerShell 5.1
- WinForms
- Windows only

## Project Structure

- `ScreenTimeTracker.ps1` - main application
- `start-tracker.vbs` - silent launcher without console window
- `start-tracker.bat` - convenience launcher
- `settings.json` - default settings
- `rules.json` - built-in classification rules
- `browser-extension/` - unpacked Chromium extension for exact site tracking
- `Image/` - application icon assets
- `data/` - runtime data, logs, caches, exports, backups

## Quick Start

1. Double-click `start-tracker.vbs`
2. Open the tray icon if the window starts minimized
3. Adjust limits in `Settings`
4. Refine classification through `Rules`, `Categories`, `Review`, and `Health`

Manual run from terminal:

```powershell
powershell -ExecutionPolicy Bypass -File .\ScreenTimeTracker.ps1
```

Run self-test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\ScreenTimeTracker.ps1 -SelfTest
```

## Browser Extension

To classify browser activity by exact domain instead of only the tab title:

1. Open `edge://extensions` or `chrome://extensions`
2. Turn on `Developer mode`
3. Click `Load unpacked`
4. Select the `browser-extension` folder

The extension sends active-tab data to `127.0.0.1:38945`.

## Main Screens

- `Today` - top activities, apps, and categories for the current day
- `History` - recent daily totals
- `Insights` - compact weekly overview
- `Week review` - slipped days, top distractions, and heatmap
- `Goals` - progress toward daily limits and study target
- `Analytics` - charts, streaks, and exports
- `Timeline` - hourly activity and session view
- `Review` - uncategorized items to fix
- `Health` - classification coverage and suggested rules
- `Rules` - loaded rule set

## Customization

Built-in parent categories stay fixed:

- `study`
- `browser_fun`
- `socials`
- `other`

You can add your own categories under those parents from the app. This keeps limits and dashboards stable while allowing more detailed classification.

You can edit:

- `Settings`
- `Categories`
- `Rules`
- `Classify current`

You can also edit `settings.json` and `rules.json` manually if you want.

## Data and Privacy

All data is stored locally.

Runtime files are written under `data/`, including:

- per-day usage files
- browser bridge cache
- summary cache
- exports
- backups
- startup logs

These runtime files are ignored by git and should normally not be committed.

## Performance Notes

The app already includes:

- per-day storage instead of one huge JSON file
- summary cache for history and analytics
- lighter background refresh behavior
- tray-first workflow

## Repository Notes

- This repository is intended for Windows users
- No external package installation is required
- Runtime data is generated automatically on first run
- License: MIT

## Roadmap

See [ROADMAP.md](ROADMAP.md).
