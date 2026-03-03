# confirwa

Lightweight Windows PowerShell watcher that shows one borderless GIF card per active Codex agent.

## Preview

Agent directory names are mosaiced in this screenshot:

![confirwa redacted preview](images/confirwa-redacted.png)

## What It Does
- Creates one floating card per detected Codex agent process.
- Reads local Codex session logs to infer per-agent state.
- Shows different GIFs for each state.
- Keeps cards small, top-most, borderless, and auto-laid out near the bottom-right corner.

## States And GIF Files
Put GIF files in `images/`:
- `1giphy.gif` -> `working`
- `2giphy.gif` -> `approval`
- `3giphy.gif` -> `reconnecting`
- `4giphy.gif` -> `idle`
- `5giphy.gif` -> `silent`

## Requirements
- Windows PowerShell 5.1+ or PowerShell 7+
- Codex CLI session logs available under `~/.codex/sessions`

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
- Does not include project/source content in this repository.

## License
MIT
