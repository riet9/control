# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-04-01

### Added

- Active screen-time tracking based on the current foreground window
- Idle detection so inactive PC time is not counted
- Daily limits for total time, study, browser fun, and social media
- Rules-based classification by process, window title, URL, and domain
- Browser extension bridge for exact site tracking in Chromium browsers
- Tray workflow with quick actions and quick glance popup
- Review screen for uncategorized activity
- Classification health screen with rule suggestions
- Custom categories mapped to fixed parent buckets
- Weekly review screen with slipped days, top distractions, and coach note
- Calendar heatmap for recent day quality
- Focus mode and hard block mode for distracting windows
- Analytics, history, timeline, sessions, goals, and exports
- Per-day storage and summary cache for better performance

### Changed

- Storage was optimized away from one giant JSON file into day-based files
- Heavy UI tabs now refresh less often and more selectively
- Relaunch now reopens the running tray instance instead of spawning duplicates

### Fixed

- Startup and runtime stability issues around idle detection, session parsing, and insights
- Safer property access under `Set-StrictMode`
- Storage path bug that could create malformed `[string]YYYY-MM-DD.json` artifacts

