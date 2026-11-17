# BachFlix Requests: Seerr setup

Replaces the "Movie/TV Show Requests" Google Form/Sheet with a request site that family and
friends sign into with their Plex account. They can see every request's status themselves.

Seerr is the successor to Overseerr. The Overseerr and Jellyseerr projects merged into Seerr in
February 2026, and the old Overseerr images are no longer maintained.

Everything below runs on the **Windows Plex server (192.168.0.5)**.

---

## 1. Install Docker Desktop (one time)

1. In an admin PowerShell, run `wsl --install`, then reboot. Seerr's docs say WSL2 is
   required on Windows to avoid database corruption.
2. Install Docker Desktop with `winget install Docker.DockerDesktop`, or download it from docker.com.
   Leave **"Use WSL 2 based engine"** checked.
3. In Docker Desktop, go to Settings > General and turn on **"Start Docker Desktop when you sign in"**.

> Docker Desktop only starts after a Windows user signs in. If the server reboots for an update
> and sits at the login screen, the site stays down until someone signs in. Turn on automatic
> sign-in on the server (`netplwiz`), or plan to sign in after reboots.

## 2. Start Seerr

Copy this `seerr` folder to the server, for example `C:\Seerr`, and run:

```powershell
cd C:\Seerr
docker volume create seerr-data
docker compose up -d
```

Open http://192.168.0.5:5055. To update later, run `docker compose pull seerr; docker compose up -d`.

`docker-compose.yml` sets `TZ=America/Denver`. Change it if you're in a different time zone.
The config is stored in the named volume `seerr-data`, not in a `C:\` folder. That's deliberate:
a Windows folder mount corrupts Seerr's database.

## 3. First-run wizard

1. **Sign in with Plex** using your own Plex account. That account becomes the Seerr admin.
2. **Plex server:** choose your server. If it isn't listed, enter it by hand: host
   `host.docker.internal`, port `32400`.
3. **Libraries:** tick your Movies and TV libraries, then run **Start Scan**. This is how Seerr
   knows what's already available and flips requests to "Available" automatically.
4. **Radarr/Sonarr:** skip this. You get titles by recording and FileBot, so you'll approve
   requests yourself and Seerr marks them Available when they show up in Plex.

## 4. Users and permissions

- **Users > Import Plex Users** brings in everyone you've shared your Plex server with.
- Under **Settings > Users**, set the default permissions for new users to **Request** (movies and
  TV). Leave **Auto-Approve** off so new requests land in your queue as **Pending**.
- **People without a Plex account:** use Users > Create Local User and give them an email and
  password. They sign in with those instead of Plex.
- Turn on **Settings > Users > Enable New Plex Sign-In** if you want anyone you've shared Plex
  with to be able to sign in without being imported first.

## 5. Notifications (so people hear about changes without checking)

**Settings > Notifications > Email** is the simplest, because Plex users already have an email
address on file. For Gmail:
- SMTP host `smtp.gmail.com`, port `587`, encryption STARTTLS
- Username: your Gmail address. Password: a Gmail **app password** (Google Account > Security >
  App passwords). Don't use your normal Gmail password.

Requesters get an email when a request is approved, declined, or becomes available. You get
one for every new request. Discord, Telegram and web push are also available on that page.

## 6. Outside access with a Cloudflare Tunnel

You need a domain on Cloudflare (moving a domain's DNS to Cloudflare is free).

1. In the Cloudflare dashboard, go to **Zero Trust > Networks > Tunnels > Create a tunnel**
   (type Cloudflared). Name it something like `bachflix`.
2. Pick Windows. Cloudflare shows an install command like
   `cloudflared.exe service install <token>`. Run it in an admin PowerShell on the Plex server.
   It installs as a Windows service, so it starts on boot without anyone signing in.
3. Add a **Public Hostname**: `requests.<yourdomain>` with service type **HTTP** and URL
   `localhost:5055`.
4. In Seerr, go to **Settings > General**:
   - **Application URL:** `https://requests.<yourdomain>`. This keeps links in emails correct.
   - Turn on **Enable Proxy Support**.
   - Turn on **Enable CSRF Protection** once the HTTPS URL works.

No router ports need to be opened.

## 7. Move the open requests from the Sheet

Do this after users are imported (step 4), so each request can be filed under the right person.

1. In the Google Sheet, go to File > Download > **Comma-separated values (.csv)** on the
   *Form Responses 1* tab and save it as `requests.csv` in this folder.
2. Copy the API key from Seerr > Settings > General, then:

```powershell
$env:SEERR_API_KEY = '<paste key>'
$env:SEERR_URL = 'http://192.168.0.5:5055'
.\Import-SheetRequests.ps1 -Prepare -SheetCsv .\requests.csv
```

This skips rows marked **Done** or **Not Available** and searches Seerr for each open row.
It writes `import-review.csv` with its best match and alternatives, plus `seerr-users.csv`
listing the user IDs.

3. Open `import-review.csv` and check every row. The sheet titles are free text, so expect a few
   wrong or missing matches.
   - If a match is wrong, take the right one from **Alternatives**, or look the title up on
     themoviedb.org and copy the number from its URL. Set **MediaType** (`movie` or `tv`) and **TmdbId**.
   - For TV, set **Seasons** to `all` or a list like `3` or `1,2`.
   - Fill in **SeerrUserId** from `seerr-users.csv` wherever it's blank. Mel's "for Gabe" and
     "for Shea" requests default to Mel.
   - Set **Include** to `Y` to import a row or `N` to skip it.
4. Do a dry run, then the real import:

```powershell
.\Import-SheetRequests.ps1 -Import          # shows what it would create
.\Import-SheetRequests.ps1 -Import -Apply  # creates the requests
```

Because they're created with the admin key, imported requests come in as **Approved**, which
means "on my list". The script checks the first one to make sure Seerr filed it under the right
person rather than under you. If Seerr didn't, the script stops so you don't end up with 29
requests in your own name.

## 8. Retire the form

- Edit the Google Form's description to say "We've moved! Request at https://requests.<yourdomain>",
  then turn off **Accepting responses**.
- Send everyone the link, and tell them to click **Sign in with Plex**.

## How your old statuses map to Seerr

| Sheet status | Seerr |
|---|---|
| (blank) | **Pending**: you haven't looked at it yet |
| Grabbing / Recording from ... / Preparing to Add | **Approved**: you're working on it |
| Done | **Available**: set automatically when it appears in Plex |
| Not Available | **Declined** |

Seerr's statuses are fixed, so there's no field like the sheet's Notes column for "recording
from Apple TV+, waiting on 1080p". For that kind of detail, message the person directly.
