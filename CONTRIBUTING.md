# Contributing

Thanks for helping improve Screen Time Tracker.

## Before Opening a PR

1. Run the self-test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\ScreenTimeTracker.ps1 -SelfTest
```

2. If you changed the launcher, rebuild it locally:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-launcher.ps1
```

3. Avoid committing runtime files from `data/`.

## Project Notes

- Main app: `ScreenTimeTracker.ps1`
- Launcher source: `launcher/ScreenTimeTrackerLauncher.cs`
- Browser extension: `browser-extension/`
- Runtime data is local-only and git-ignored

## Good Issues to Help With

- classification accuracy
- performance and startup time
- analytics and reporting
- focus mode and blocking behavior
- UI polish and reliability
