# Roadmap

## 1. Stabilization

- Fix and clean up malformed storage artifacts like `[string]YYYY-MM-DD.json`
- Add a dedicated runtime error log with function and line context
- Add a small diagnostics screen for startup, browser bridge, and storage health
- Add migration checks for future storage schema updates

## 2. Classification Quality

- Suggest category together with rule target and pattern
- Add rule confidence levels
- Add bulk review actions for uncategorized items
- Improve browser classification for mixed domains like YouTube
- Show which rules are stale or never used

## 3. Anti-Distraction

- Stronger anti-reopen blocking after hard limit
- Optional blocklist by process and domain
- Timed unlock tokens instead of simple snooze
- Study-first mode: unlock fun only after study minimum is reached

## 4. Analytics

- Export a readable weekly review report
- Add monthly review
- Add session-quality metrics like average study block length
- Add reopen-count analytics for distracting apps
- Add "most expensive hour" and "most common slip hour"

## 5. UX

- Add a dedicated onboarding flow for first launch
- Add a lightweight "publishable" dashboard theme for screenshots and sharing
- Make tray commands configurable
- Add a compact first-run settings preset for students

## 6. Architecture

- Move active-window polling and storage helpers into a small C# helper
- Separate UI code from storage / tracking logic
- Add a small test suite for classification and storage edge cases
- Consider a future migration to a compiled desktop app if startup speed becomes a top priority

## 7. GitHub / Open Source

- Add screenshots or short demo GIFs
- Decide on a license
- Add a changelog
- Add issue templates for bug reports and feature requests
- Add a release checklist for shipping builds
