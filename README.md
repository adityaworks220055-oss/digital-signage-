# Digital Signage Application

Flask-SocketIO digital signage application for managing screens, images, image
groups, fallback content, and task-completion notifications.

## Requirements

- Windows with Python 3.10 or newer
- Python packages:
  - `flask`
  - `flask-socketio`
  - `pandas`
  - `openpyxl`
  - `werkzeug`
  - `waitress` (production mode)

Install the packages with:

```powershell
python -m pip install -r requirements.txt
```

## Start the server

Open PowerShell in the project folder:

```powershell
cd "C:\Users\Adityavikram.Bhatt\OneDrive - Motherson Group\Desktop\prog"
python sign.txt.py
```

### Production mode on Windows

For a production-style HTTP server, install the dependencies from
`requirements.txt` and start with:

```powershell
$env:PRODUCTION = "true"
python sign.txt.py
```

Production mode uses Waitress with the Socket.IO WSGI application wrapper, so
normal Flask routes and Socket.IO long-polling events continue to work. The
default worker count is eight; adjust it with:

```powershell
$env:WAITRESS_THREADS = "16"
```

Waitress is now the default when the script is started directly, so an
accidental launch does not expose Flask's Werkzeug development server. To
explicitly run the development server for local debugging:

```powershell
$env:DEVELOPMENT = "true"
python sign.txt.py
```

For public or security-sensitive deployments, put the service behind HTTPS and
an authenticated reverse proxy. Waitress does not provide TLS termination.

Waitress does not provide WebSocket support. The dashboard and screen clients
therefore explicitly use Socket.IO long-polling with WebSocket upgrades
disabled. The production entry point serves Flask-SocketIO's own WSGI
middleware rather than a raw `python-socketio` wrapper, preserving Flask's
request context for Socket.IO event handlers and avoiding WSGI
socket-environment errors.

After restarting Waitress, browsers that were open during the restart can
briefly send their old Socket.IO session ID. The clients are configured to
create a fresh polling connection and reconnect automatically. If an old
browser tab continues showing connection errors, refresh that tab once.
Nginx polling proxying has buffering and caching disabled for the same reason.

### Nginx reverse proxy for the intranet

The project includes [`nginx\nginx.conf`](nginx/nginx.conf), which exposes the
application on port 80 and proxies requests to Waitress on `127.0.0.1:3000`.
It includes the required `/socket.io/` upgrade headers and supports the
application's large multi-image uploads.

1. Download the Windows Nginx package from [nginx.org](https://nginx.org/en/download.html)
   and extract it, for example to `C:\nginx`.
2. Copy `nginx\nginx.conf` from this project to `C:\nginx\conf\nginx.conf`.
3. Start the application locally behind Nginx:

   ```powershell
   $env:PRODUCTION = "true"
   $env:HOST = "127.0.0.1"
   $env:PORT = "3000"
   python sign.txt.py
   ```

4. In an elevated PowerShell, validate and start Nginx:

   ```powershell
   cd C:\nginx
   .\nginx.exe -t
   .\nginx.exe
   ```

5. Allow the Nginx intranet port through Windows Firewall:

   ```powershell
   New-NetFirewallRule `
     -DisplayName "Digital Signage Nginx HTTP 80" `
     -Direction Inbound `
     -Protocol TCP `
     -LocalPort 80 `
     -Action Allow `
     -Profile Domain,Private
   ```

   Port 3000 should remain restricted to localhost when Nginx is used.

Open `http://<host-lan-ip>/health` from another LAN device. To stop or reload
Nginx:

```powershell
.\nginx.exe -s quit
.\nginx.exe -s reload
```

Nginx is not installed automatically by the Python dependency file because it
is a separate Windows service. If the firewall command is denied, the machine
is likely governed by corporate policy and the exception must be approved by
IT.

By default the application listens on all network interfaces at port `3000`.
Open the local dashboard at:

```text
http://localhost:3000
```

The default administrator username is `admin`. Set a secure password before
first use:

```powershell
$env:ADMIN_PASSWORD = "replace-with-a-strong-password"
python sign.txt.py
```

For production-like use, also set a stable secret key:

```powershell
$env:SECRET_KEY = "replace-with-a-long-random-secret"
```

`ADMIN_PASSWORD` is used when the administrator record is created. Changing
the environment variable later does not automatically change an existing
password stored in Excel.

## LAN and other-device access

The default bind address is `0.0.0.0:3000`, so another device on the same
network can connect using the host computer's IPv4 address:

```text
http://192.168.1.25:3000
```

Find the host IPv4 address with:

```powershell
ipconfig
```

If another device cannot connect:

1. Confirm both devices are on the same LAN or Wi-Fi.
2. Allow Python or inbound TCP port `3000` through Windows Firewall. Run
   PowerShell **as Administrator** and use:

   ```powershell
   New-NetFirewallRule -DisplayName "Digital Signage TCP 3000" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3000 -Profile Domain,Private
   ```

3. Confirm the server is running and listening on the expected port.
4. Use the host's LAN IPv4 address, not `localhost`.

Router port forwarding is not required for devices on the same LAN. On a
corporate or guest Wi-Fi network, client isolation or VLAN separation may
prevent devices from reaching one another even when the server is configured
correctly.

Do not expose the development server directly to the public internet. Use a
VPN or a secured reverse proxy for remote access.

## Data and storage

The application stores persistent data in the project folder:

| Location | Purpose |
| --- | --- |
| `database.xlsx` | Excel-backed metadata and bookkeeping |
| `uploads\` | Uploaded image files |
| `templates\` | HTML pages |
| `sign.txt.py` | Flask application |

Excel worksheets:

- `admins`: usernames, password hashes, and roles
- `images`: image IDs, physical filenames, display names, groups, upload times
- `screens`: screen IDs, assigned image IDs, server groups, playback mode, and last heartbeat
- `audit_logs`: historical actions and bookkeeping
- `notifications`: task-completion notifications and dismissal timestamps

Image bytes are stored in `uploads\`; Excel stores their metadata and
references. Keep `database.xlsx` and `uploads\` together when backing up or
migrating the application.

The application contains legacy migration support for older workbook and
upload locations. Do not delete legacy files until migration has been checked.

## Dashboard workflows

### Screens

The dashboard shows:

- Total screens
- Online screens
- Current image or fallback state
- Whether the current content is a single image or image group
- A link to manage content for each screen
- Server deletion and screen search

Screen content assignment is handled on the screen-specific **Manage content**
page. Select exactly one mode:

- **Single image**: assigns one image. Selecting an empty fallback option clears
  the assignment.
- **Image group**: assigns the first image in the group and enables grouped
  playback on the screen.

The dashboard refreshes screen status data every five seconds without
reloading the page.

### Screen online expiry

A screen is shown as **Online** when its most recent heartbeat is less than
45 seconds old. After 45 seconds without a heartbeat it is shown as
**Offline**. The screen record is not deleted and becomes online again when
heartbeats resume.

Heartbeat timestamps are persisted to Excel at most once every 30 seconds per
screen. Runtime online state is held in memory.

### Image library

The Image library supports:

- Searching images by display name, filename, or group
- Selecting several images
- Grouping selected images under one group name
- Mass deletion of selected images
- Renaming an image display name
- Deleting individual images

When images are deleted, screens using those images are cleared to fallback.
If deletion removes the last image in a group, screens assigned to that group
also return to fallback.

### Uploads

The Add images area contains a drag-and-drop upload box.

- One selected file is treated as a single-image upload by default.
- Multiple files are treated as a group upload by default.
- The user can explicitly choose **Treat as single images** for multiple files;
  those files are stored separately without a group.
- Maximum upload batch: 100 images.
- Maximum size per image: 5 MB.
- Supported extensions: PNG, JPG, JPEG, GIF, and WEBP.

Grouped uploads require a group name. A group may contain no more than 100
images. Upload validation is performed both in the browser and on the server.

### Image groups

The Image groups page supports:

- Searching and paging through groups
- Viewing all images in a selected group
- Deleting images from a group
- Renaming a group
- Deleting a group

Deleting a group keeps its image files and library records, but clears their
group names. Screens using images from the deleted group return to fallback.

### Paired deployment groups

The **Deployment groups** page supports one-to-one distribution when a server
group and an image group contain the same number of items:

1. Select screens and save them under a server group name.
2. Create an image group in the Image library.
3. Choose both groups on the Deployment groups page.
4. Review the pairing preview and deploy.

Screens are ordered by screen number and images retain their image-library
order. For example, screen 1 receives image 1, screen 2 receives image 2, and
so on. The deployment is stored in the `screens` worksheet and connected
screens receive a live update. Paired screens display only their assigned image
and do not enter the normal image-group playlist. A later individual image or
image-group assignment clears the paired mode.

The Deployment groups page also provides a server search box while creating a
group. Deleting a server group only removes the group label from its screens;
it does not delete screens, images, or current content assignments.

Each server group has an ordered queue of up to 20 image groups, stored in the
`queues` worksheet. Only image groups whose image count matches the number of
servers in the group can be queued. On the Deployment groups page, each server
group card shows:

- **Now playing**: the image currently assigned to each screen.
- **Queue**: numbered items with buttons to move an item to the top, up, down or
  to the bottom, and to remove it. **Clear queue** empties the whole queue.
- **Add to queue** / **Play now**: pick an image group and either append it to
  the queue or play it immediately without changing the queue.
- **Play next**: plays queue item 1 and removes it from the queue. The same
  action is available from the grouped server controls on the dashboard.

Renaming an image group updates the queue entries that use it. Deleting an image
group or a server group removes its queue entries. Values in the old single-slot
`queued_group_name` column are moved into the queue automatically at startup.

### Fallbacks

The Fallback screens page assigns a fallback image to selected screens.
Fallback assignment is persisted in Excel and connected screens receive live
Socket.IO updates.

### Notifications

When a screen emits **Mark as done**:

1. A historical `mark_done` action is written to `audit_logs`.
2. A user-facing notification is written to `notifications`.
3. Connected dashboards receive a Socket.IO toast.

Dashboard toasts appear in the bottom-right, remain for 10 seconds, include an
X button, and display at most two at a time. Older visible toasts are removed
when newer ones arrive.

The Notifications page supports searching, paging five notifications per page,
and dismissing notifications. Dismissal sets `dismissed_at`; it does not delete
the Excel row.

## Screen client behavior

Screen pages connect through Socket.IO and support:

- Registration and heartbeat updates
- Live image updates
- Grouped Previous and Next playback
- Mark as done on the final group image

The screen page URL is:

```text
http://<server-ip>:3000/screen/<screen-id>
```

## Useful configuration

The server address and port can be changed through environment variables:

```powershell
$env:HOST = "0.0.0.0"
$env:PORT = "3000"
python sign.txt.py
```

Important application limits are defined near the top of `sign.txt.py`:

- `MAX_UPLOAD_IMAGES = 100`
- `MAX_GROUP_IMAGES = 100`
- `MAX_IMAGE_BYTES = 5 * 1024 * 1024`

## Backup and recovery

Stop the application before making a backup. Copy both:

```text
database.xlsx
uploads\
```

Restoring only the workbook will leave image records pointing to missing
files. Restoring only `uploads\` will leave files without their metadata.

Avoid keeping `database.xlsx` open in Excel while the application is writing.
Excel file locking can prevent atomic workbook replacement.

## Troubleshooting

### The application will not start

- Confirm Python is installed: `python --version`
- Install the required packages.
- Run `python -m py_compile sign.txt.py` to check syntax.
- Check that port `3000` is not already in use.

### Upload fails

- Confirm the extension is supported.
- Confirm each image is 5 MB or smaller.
- Confirm the batch contains no more than 100 files.
- For grouped uploads, provide a group name.
- Confirm `uploads\` is writable.

### A screen is offline

- Confirm the screen device can reach `http://<server-ip>:3000`.
- Confirm the screen is using the correct screen ID.
- Check that the host firewall permits port `3000`.
- Remember that the offline threshold is 45 seconds without a heartbeat.

### Excel writes fail

Close `database.xlsx` in Excel and retry. The application writes the workbook
through a temporary file and replaces the active workbook atomically.

## Validation commands

Compile the application:

```powershell
python -m py_compile sign.txt.py
```

The project currently uses focused Flask smoke tests rather than a committed
test suite. Tests should use temporary workbook state and temporary upload
folders so the real database and image library are not modified.
