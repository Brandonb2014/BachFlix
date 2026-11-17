# QuadPlex Media Server Watchdog: Setup Instructions

**For:** the Claude session running on the QuadPlex server.
**Goal:** keep Plex, Emby and Jellyfin running. A scheduled task runs a small batch script every 5 minutes. The script starts any server that has stopped.

Jellyfin keeps shutting down on QuadPlex, and Brandon has to remote in to start it again. This setup fixes that.

## Background

- `MediaServerWatchdog.bat` is in the same folder as this file. It was written and tested on the main PC (NODNARB).
- This folder reaches QuadPlex through a **read-only sync** from NODNARB. Changes made here never sync back and may be overwritten. So **copy the script to `C:\Scripts\` and edit only that copy.**
- How the script works: for each server, if a Windows service with the configured name exists and isn't running, it runs `net start`. If there's no such service, it checks for the program's process and launches the program if it's missing. It writes to `MediaServerWatchdog.log`, next to the script, only when it starts something or hits a problem.

## Steps

### 1. Find out how each server is installed

Run this in an admin prompt:

```
sc query type= service state= all | findstr /i "plex emby jellyfin"
```

For each server, note whether it's a **service** (and its exact `SERVICE_NAME`) or a **normal program**. If it's a program, find its exe path and process name:

```
tasklist | findstr /i "plex emby jellyfin"
```

Common locations (check them, don't assume):

| Server   | Usual service name | Usual program path |
|----------|--------------------|--------------------|
| Plex     | *(usually not a service)* | `C:\Program Files\Plex\Plex Media Server\Plex Media Server.exe` |
| Emby     | `EmbyServer` | `%AppData%\Emby-Server\system\EmbyServer.exe` |
| Jellyfin | `JellyfinServer` | `C:\Program Files\Jellyfin\Server\jellyfin.exe` |

### 2. Copy the script and set it up

1. Create `C:\Scripts\` and copy `MediaServerWatchdog.bat` into it.
2. In `C:\Scripts\MediaServerWatchdog.bat`, edit the three `call :Check` lines in the settings block so they match what step 1 found. The format is:
   ```
   call :Check "Name" "ServiceName" "ProcessName.exe" "Full\Path\To\Program.exe"
   ```
   - If a server runs as a program, not a service, leave its service name as a name that doesn't exist (or `""`). The script will then fall back to the process check.
   - Check the process name. It must exactly match the name `tasklist` shows.

### 3. Test it once

While all three servers are running, run `C:\Scripts\MediaServerWatchdog.bat` from an admin prompt. **The script should print nothing.** If it prints a "not found" line, fix the settings and run it again.

### 4. Create the scheduled task

Choose based on step 1:

- **All three are services.** Run the task as SYSTEM:
  ```
  schtasks /Create /TN "QuadPlex Media Watchdog" /TR "C:\Scripts\MediaServerWatchdog.bat" /SC MINUTE /MO 5 /RU SYSTEM /RL HIGHEST /F
  ```
- **Any server is a desktop program (Plex usually is).** Run the task as Brandon's Windows account. If it ran as SYSTEM, Plex would start under the SYSTEM profile with a different, empty library. `/RP *` asks for the password, so **Brandon must type it. Never enter it for them.**
  ```
  schtasks /Create /TN "QuadPlex Media Watchdog" /TR "C:\Scripts\MediaServerWatchdog.bat" /SC MINUTE /MO 5 /RU "<Brandon's account>" /RP * /RL HIGHEST /F
  ```

Then check the task: `schtasks /Query /TN "QuadPlex Media Watchdog" /V /FO LIST`

### 5. Optional (recommended): Windows auto-restart for Jellyfin

If Jellyfin is a service, Windows can restart it within a minute of a crash, and the watchdog becomes the backup for clean stops. **Ask Brandon first**, since this changes the service's settings:

```
sc failure JellyfinServer reset= 86400 actions= restart/60000/restart/60000/restart/60000
```

Use the real service name from step 1.

### 6. Check that it works

**Ask Brandon first.** This briefly takes Jellyfin down for anyone watching.

1. Stop Jellyfin (`net stop JellyfinServer`, or close the program).
2. Run the task now: `schtasks /Run /TN "QuadPlex Media Watchdog"`
3. Check that Jellyfin is back and that `C:\Scripts\MediaServerWatchdog.log` has a "started" line.

### 7. Report back to Brandon

Tell Brandon:
- how each server is installed (service name or program path)
- which task account option you used
- whether step 5 was done
- the result of step 6

### 8. Extra: why Jellyfin keeps stopping

Look at the latest files in `C:\ProgramData\Jellyfin\Server\log\` and at Windows Event Viewer (Windows Logs > Application and System, around the times it stopped). Give Brandon a short summary of the likely cause. **Don't change Jellyfin's settings without asking.**

## If anything in the script needs to change

Don't edit the copy in the synced folder. Edit `C:\Scripts\MediaServerWatchdog.bat`. If the change should go back into the repo, write a short handoff note for NODNARB (what changed and why) in a handoff folder on the Desktop, since the sync is one-way.
