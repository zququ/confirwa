# confirwa

Lightweight Windows PowerShell watcher that shows one borderless GIF card per active Codex agent.

## Preview

Agent directory names are mosaiced in this screenshot:

![confirwa redacted preview](images/confirwa-redacted.png)

## What It Does
- Creates one floating card per detected Codex agent process.
- Reads local Codex session logs to infer per-agent state.
- Shows different GIFs for each state.
- Keeps cards small, top-most, borderless, and auto-laid out in fixed slots.

## Interaction
- Left click a card: switch to mapped terminal tab (`Ctrl+Alt+<slot>`).
- Left drag a card: reorder slot mapping only (cards stay in fixed row layout).
- Right drag a card: move the whole confirwa strip as a group.
- Slot order and strip offset are persisted in `~/.cache`.

## States And GIF Files
Put GIF files in `images/`:
- `1giphy.gif` -> `working`
- `2giphy.gif` -> `approval`
- `3giphy.gif` -> `reconnecting`
- `4giphy.gif` -> `idle`
- `5giphy.gif` -> `silent`

## State Detection
- `approval`: prefers real approval signals from Codex events/function calls (`sandbox_permissions=require_escalated`) and approval prompts like `Yes, proceed (y)`.
- `reconnecting`: uses session events plus Codex DB log fallback (`state_5.sqlite`) for lines like `stream disconnected`, `retrying turn`, `reconnecting`.
- `working/silent/idle`: inferred from active turns and recent event timestamps.
- Priority is designed so `approval` is not easily overridden by `silent`.

## Agent Mapping
- Active cards are mapped from live Codex processes and live thread/session mapping in `~/.codex/state_5.sqlite`.
- Card count uses the max of:
  - mapped active rollout threads
  - live `codex` process count
  - fresh session count
- This lets new Codex instances appear immediately, even before full session mapping is available.

## Requirements
- Windows PowerShell 5.1+ or PowerShell 7+
- Codex CLI session logs available under `~/.codex/sessions`
- `sqlite3` available in `PATH` (for active-thread/reconnect fallback)

## Run
From this repo directory:

```powershell
pwsh -File .\confirwa-watcher.ps1
```

Or with custom image directory:

```powershell
pwsh -File .\confirwa-watcher.ps1 -ImageDir "D:\path\to\your\gifs"
```

Optional tuning:

```powershell
pwsh -File .\confirwa-watcher.ps1 -PollMs 400 -IdleSeconds 1.5
```

## Stop
Stop the PowerShell process running `confirwa-watcher.ps1`.

## Privacy
- No network upload logic.
- Reads only local process info and local Codex session files.
- Reads local Codex state DB (`~/.codex/state_5.sqlite`) for live mapping/status fallback.
- Does not include project/source content in this repository.

## License
MIT
