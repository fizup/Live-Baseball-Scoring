# Live Game Tracker

A live baseball/softball scorekeeping tool built entirely in Excel VBA. Track pitches, plays, baserunners, substitutions, and pitcher stats in real time, then export a fully formatted, human-readable play-by-play log to a worksheet.

## Features

- **Live pitch tracking** — balls, strikes, fouls, automatic walk/strikeout handling
- **Dynamic play resolution** — every play (hits, outs, advances) is driven by an editable glossary worksheet, not hardcoded logic
- **Baserunner resolution** — a dedicated form lets you resolve what every runner does on a play, including multi-segment resolution for plays involving a fielding error (e.g. "advanced to third, then scored on error")
- **Error tracking** — one misplay is counted once, regardless of how many runners it affects; hits and errors are tallied per half-inning
- **Full undo** — every action (pitch, play, substitution) can be undone via a memento-based undo stack, including lineup changes
- **Substitutions & lineup changes** — a flexible staging editor lets you assign any player to any batting spot / defensive position in one atomic operation, correctly phrasing pinch hitters, pinch runners, and DP/Flex-style swaps in the log
- **Pitcher stats** — live pitch count and innings pitched per pitcher
- **Pause & resume** — hide the tracker mid-game and resume exactly where you left off
- **Styled export** — the exported log alternates row colors by team, bolds scoring plays, italicizes substitutions, and inserts a Runs/Hits/Errors/LOB summary after every half-inning

## Getting Started

1. Open the workbook in Excel (macros must be enabled).
2. Set up two sheets, `LineUpAway` and `LineUpHome`, each listing your roster starting at row 3 (Spot, Position, Jersey, Name, IsPitcher).
3. Fill in `Baseball_Glossary` with the plays your league uses (see below).
4. Run `modMain.StartLiveGameTracker` to begin a game.
5. Use `modMain.ResumeGame` to reopen a paused session.

## Project Structure

### Classes
| File | Purpose |
|---|---|
| `clsBaseballGame.cls` | Core game engine — score, outs, inning/half, base state, batting order, memento capture/restore |
| `clsPlayer.cls` | A single roster player (spot, position, jersey, name, pitcher flag, pitch count, outs pitched) |
| `clsEvent.cls` | One row of the glossary — a definable play type (code, text, section, target base, hit/error flags) |
| `clsRunnerOutcome.cls` | One resolved baserunner decision from `frmResolveRunners` |
| `clsPlayByPlayEvent.cls` | A single logged play-by-play row, including score/base/count context at the time |
| `clsGameHistoryLogger.cls` | Stores all logged events and exports them to a styled worksheet |
| `clsGameMemento.cls` | A full snapshot of engine state, used for undo |
| `clsLineupMemento.cls` | A deep-copied snapshot of both rosters, used for undo across substitutions |
| `clsLineupAssignment.cls` | One staged Spot/Position change from the substitution form |

### Forms
| File | Purpose |
|---|---|
| `frmRecordPlay.frm` | Main live-input window — pitch buttons, play selection, undo, pause/export |
| `frmResolveRunners.frm` | Multi-segment baserunner resolution, including error segments |
| `frmSubstitution.frm` | Stages and applies any number of lineup/position changes in one batch |

### Modules
| File | Purpose |
|---|---|
| `modMain.bas` | Entry points (`StartLiveGameTracker`, `ResumeGame`) and shared constants |
| `modUtils.bas` | Roster loading (`GetRoster`) and glossary loading (`GetEvents`) from worksheets |

## The Glossary

`Baseball_Glossary` is the single source of truth for what plays exist. Each row defines:

- **Code** — short identifier (e.g. `1B`, `K`, `SB`)
- **Play Text** — human-readable phrase used in the log
- **Name** — unique key used to look up the event
- **Section** — `GetOnBase`, `GetOut`, or `AdvanceBase`
- **BallIsHit** — whether the pitch sequence gets an `X` appended
- **TargetBase** — where the batter ends up (for `GetOnBase` events)
- **IsHit** — whether it counts as a hit in the summary line

Adding a new play type is usually just adding a new row — no code changes required.

## Architecture Notes

- **Memento pattern for undo** — every mutating action pushes a full state snapshot before executing. Lineup changes additionally attach a deep-copied roster snapshot, since substitutions aren't covered by the base engine snapshot alone.
- **`Spot` vs. `Position`** — a player's batting order position (`Spot`) and defensive position (`Position`) are independent fields. This is what makes DP/Flex substitutions (common in softball) possible: a player can keep her batting spot while moving to the bench defensively, or vice versa.
- **`CurrentBatter`** resolves by searching for the player whose `Spot` matches the current index — not by positional index into the lineup collection — since substitutions change `Spot` values without reordering the underlying collection.
- **`CurrentPitcher`** resolves by scanning for `IsPitcher = True`, independent of `Spot`, so it stays correct through any substitution automatically.

## Known Limitations

- Rule 5.08(b) (a run doesn't count if it scores on the same play as a force-out that ends the inning) is not automatically enforced — the scorer is trusted to enter the correct outcome, same as with a paper scorebook.
- Fielder/position-specific error attribution (e.g. "E6") is not tracked individually; only a per-play error count is recorded.
- The tool assumes a 9-player active lineup at any time; there is no built-in validation against duplicate batting spots or positions when staging lineup changes.

## Requirements

- Excel with VBA macros enabled (developed and tested primarily on Windows Excel; some `Collection`/`New` object patterns were adjusted for compatibility with older Excel versions)
