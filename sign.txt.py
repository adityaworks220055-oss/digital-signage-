import os
import time
import uuid
import secrets
import threading
from datetime import datetime, timedelta
from pathlib import Path
import shutil
import pandas as pd
from flask import (
    Flask,
    request,
    redirect,
    url_for,
    session,
    jsonify,
    render_template,
    send_from_directory,
    flash,
)
from flask_socketio import SocketIO, emit, join_room
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

BASE = Path(__file__).resolve().parent

LEGACY_DATA = BASE / "signage_data.xlsx"
LEGACY_LOCAL_DATA = Path(
    os.getenv(
        "LOCALAPPDATA",
        Path.home() / "AppData" / "Local",
    )
) / "DigitalSignage" / "signage_data.xlsx"

APP_DATA = Path(
    os.getenv(
        "LOCALAPPDATA",
        Path.home() / "AppData" / "Local",
    )
) / "DigitalSignage"
LEGACY_DATABASE = APP_DATA / "database.xlsx"
LEGACY_LOCAL_UPLOADS = APP_DATA / "uploads"

DATA = BASE / "database.xlsx"
UPLOADS = BASE / "uploads"
LEGACY_UPLOADS = BASE / "static" / "uploads"
DATA.parent.mkdir(parents=True, exist_ok=True)
UPLOADS.mkdir(parents=True, exist_ok=True)

def migrate_legacy_storage():
    if not DATA.exists():
        for legacy_data in (
            LEGACY_DATABASE,
            LEGACY_LOCAL_DATA,
            LEGACY_DATA,
        ):
            if legacy_data.exists():
                shutil.copy2(legacy_data, DATA)
                break

    for legacy_uploads in (LEGACY_LOCAL_UPLOADS, LEGACY_UPLOADS):
        if legacy_uploads.is_dir():
            for legacy_file in legacy_uploads.iterdir():
                persistent_file = UPLOADS / legacy_file.name
                if legacy_file.is_file() and not persistent_file.exists():
                    shutil.copy2(legacy_file, persistent_file)

ALLOWED = {"png", "jpg", "jpeg", "gif", "webp"}

LOCK = threading.RLock()

SHEETS = {
    "admins": ["username", "password_hash", "role"],
    "images": ["id", "filename", "display_name", "group_name", "uploaded_at"],
    "screens": [
        "screen_id",
        "image_id",
        "group_name",
        "playback_mode",
        "playback_group_name",
        "queued_group_name",
        "last_seen",
    ],
    "audit_logs": ["at", "username", "action", "details", "dismissed_at"],
    "notifications": ["at", "username", "action", "details", "dismissed_at"],
    "queues": ["id", "server_group", "image_group", "queued_at", "queued_by"],
}
MAX_QUEUE_ITEMS = 20
MAX_GROUP_IMAGES = 100
MAX_UPLOAD_IMAGES = 100
MAX_IMAGE_BYTES = 5 * 1024 * 1024

ONLINE = {}
LAST_HEARTBEAT_WRITE = {}
HEARTBEAT_WRITE_INTERVAL = 30
app = Flask(
    __name__,
    template_folder=str(BASE / "templates"),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_IMAGES * MAX_IMAGE_BYTES + (2 * 1920 * 1080)

app.secret_key = os.getenv(
    "SECRET_KEY",
    secrets.token_hex(32),
)

app.permanent_session_lifetime = timedelta(hours=8)

socketio = SocketIO(
    app,
    async_mode="threading",
    cors_allowed_origins=[],
    logger=False,
    engineio_logger=False,
)


def init_book():
    with LOCK:
        migrate_legacy_storage()

        if not DATA.exists():
            with pd.ExcelWriter(
                DATA,
                engine="openpyxl",
            ) as writer:

                for sheet, cols in SHEETS.items():
                    pd.DataFrame(columns=cols).to_excel(
                        writer,
                        sheet_name=sheet,
                        index=False,
                    )

        admins = read("admins")

        if admins.empty:
            admins.loc[len(admins)] = [
                "admin",
                generate_password_hash(
                    os.getenv(
                        "ADMIN_PASSWORD",
                        "change-me",
                    )
                ),
                "admin",
            ]

            write("admins", admins)

        write("images", read("images"))
        migrate_legacy_queue()


def migrate_legacy_queue():
    """Move single-slot queued_group_name values into the ordered queues sheet."""
    screens = read("screens")
    pending = screens.queued_group_name.astype(str).str.strip() != ""
    if not pending.any():
        return

    queues = read("queues")
    seen = set()
    for _, screen in screens.loc[pending].iterrows():
        server_group = str(screen["group_name"]).strip()
        image_group = str(screen["queued_group_name"]).strip()
        key = (server_group.casefold(), image_group.casefold())
        if not server_group or key in seen:
            continue
        seen.add(key)
        queues.loc[len(queues)] = [
            uuid.uuid4().hex,
            server_group,
            image_group,
            datetime.now().isoformat(timespec="seconds"),
            "migration",
        ]

    screens.loc[pending, "queued_group_name"] = ""
    write("queues", queues)
    write("screens", screens)


def read(sheet):

    with LOCK:

        try:
            # Read every cell as text so names like "1" are not turned into
            # 1 or 1.0 depending on whether the column also has blank cells.
            frame = pd.read_excel(
                DATA,
                sheet_name=sheet,
                dtype=str,
            ).fillna("")
            for column in SHEETS[sheet]:
                if column not in frame.columns:
                    frame[column] = ""
            if sheet == "images":
                missing_names = frame["display_name"].astype(str).str.strip() == ""
                frame.loc[missing_names, "display_name"] = frame.loc[missing_names, "filename"]
            return frame
        except Exception:
            return pd.DataFrame(
                columns=SHEETS[sheet]
            )


def write(sheet, frame):

    with LOCK:

        frames = {
            name: read(name)
            for name in SHEETS
        }

        frames[sheet] = frame[
            SHEETS[sheet]
        ]

        tmp = DATA.with_name(
            f".{DATA.stem}.{uuid.uuid4().hex}.tmp.xlsx"
        )

        with pd.ExcelWriter(
            tmp,
            engine="openpyxl",
        ) as writer:

            for name, df in frames.items():
                df.to_excel(
                    writer,
                    sheet_name=name,
                    index=False,
                )

        try:
            for attempt in range(5):
                try:
                    os.replace(tmp, DATA)
                    break
                except PermissionError:
                    if attempt == 4:
                        raise RuntimeError(
                            f"Cannot update {DATA.name}; close it in Excel and retry."
                        )
                    time.sleep(1)
        finally:
            if tmp.exists():
                tmp.unlink()


def audit(action, details=""):

    logs = read("audit_logs")

    logs.loc[len(logs)] = [
        datetime.now().isoformat(
            timespec="seconds"
        ),
        session.get("user", "system"),
        action,
        details,
        "",
    ]

    write("audit_logs", logs)


def create_notification(action, details=""):
    notifications = read("notifications")
    notifications.loc[len(notifications)] = [
        datetime.now().isoformat(timespec="seconds"),
        session.get("user", "system"),
        action,
        details,
        "",
    ]
    write("notifications", notifications)


def logged_in():
    return bool(
        session.get("user")
    )


def login_required(fn):
    def wrapped(*args, **kwargs):

        if not logged_in():
            return redirect(
                url_for(
                    "login",
                    next=request.path,
                )
            )

        return fn(*args, **kwargs)

    wrapped.__name__ = fn.__name__
    return wrapped


TEMPLATE_LOGIN = "login.html"

TEMPLATE_DASH = "dashboard.html"

TEMPLATE_GROUPS = "groups.html"

TEMPLATE_DEPLOYMENT_GROUPS = "deployment_groups.html"

TEMPLATE_FALLBACKS = "fallbacks.html"

TEMPLATE_SCREEN = "screen.html"

TEMPLATE_NOTIFICATIONS = "notifications.html"

def image_url(image_id):
    if not image_id:
        return "/static/fallback.svg"
    images = read("images")
    hit = images[images.id.astype(str) == str(image_id)]
    if hit.empty:
        return "/static/fallback.svg"
    filename = str(hit.iloc[0].filename)
    if not (UPLOADS / filename).is_file():
        return "/static/fallback.svg"
    return "/uploads/" + filename


@app.get("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(UPLOADS, filename)


@app.errorhandler(413)
def file_too_large(e):
    return (
        f"Upload too large. Each image can be at most {MAX_IMAGE_BYTES // (1024 * 1024)} MB, "
        f"with up to {MAX_UPLOAD_IMAGES} images per upload.",
        413,
    )


@app.route(
    "/login",
    methods=["GET", "POST"],
)
def login():

    if request.method == "POST":

        users = read("admins")

        hit = users[
            users.username.astype(str)
            ==
            request.form.get(
                "username",
                "",
            )
        ]

        if (
            not hit.empty
            and
            check_password_hash(
                hit.iloc[0].password_hash,
                request.form.get(
                    "password",
                    "",
                )
            )
        ):
            session["user"] = (
                request.form["username"]
            )
            session.permanent = True

            audit("login")

            return redirect(
                request.args.get(
                    "next",
                    "/",
                )
            )

        audit(
            "failed_login",
            request.form.get("username", ""),
        )
        flash(
            "Invalid credentials"
        )

    return render_template(
        TEMPLATE_LOGIN
    )


@app.get("/logout")
def logout():
    if logged_in():
        audit("logout")
    session.clear()
    return redirect(
        url_for("login")
    )


@app.get("/")
@login_required
def dashboard():

    all_images = read("images").to_dict("records")
    screen_q = request.args.get("screen_q", "").strip()
    image_q = request.args.get("image_q", "").strip()
    image_query = image_q.casefold()
    imgs = [
        image for image in all_images
        if not image_query
        or image_query in " ".join(
            str(image.get(field, "")) for field in
            ("display_name", "filename", "group_name")
        ).casefold()
    ]
    image_groups = sorted({
        str(image.get("group_name", "")).strip()
        for image in all_images
        if str(image.get("group_name", "")).strip()
    }, key=str.casefold)
    page_size = 3
    page = max(1, request.args.get("page", 1, type=int))
    image_page_size = 10
    image_page = max(1, request.args.get("image_page", 1, type=int))

    img_map = {
        str(i["id"]): i
        for i in all_images
    }

    queued_by_group = {}
    for _, item in read("queues").iterrows():
        queued_by_group.setdefault(
            str(item["server_group"]).strip().casefold(),
            [],
        ).append(str(item["image_group"]))

    result = []

    for _, s in read("screens").iterrows():

        image = img_map.get(
            str(s.image_id),
            {},
        )
        queue = queued_by_group.get(str(s.group_name).strip().casefold(), [])

        screen = {
            **s.to_dict(),
            "queue_next": queue[0] if queue else "",
            "queue_length": len(queue),
            "filename": image.get("display_name", ""),
            "selection_type": (
                "Image group" if str(image.get("group_name", "")).strip()
                else "Single image" if image
                else "Fallback"
            ),
            "online": time.time() - ONLINE.get(str(s.screen_id), 0) < 45,
        }
        if not screen_q or screen_q.casefold() in " ".join(
            str(screen.get(field, "")) for field in
            ("screen_id", "group_name", "filename", "image_id")
        ).casefold():
            result.append(screen)

    page_count = max(1, (len(result) + page_size - 1) // page_size)
    page = min(page, page_count)
    page_start = (page - 1) * page_size
    image_page_count = max(1, (len(imgs) + image_page_size - 1) // image_page_size)
    image_page = min(image_page, image_page_count)
    image_start = (image_page - 1) * image_page_size
    grouped_screens = [
        screen for screen in result
        if str(screen.get("group_name", "")).strip()
    ]
    grouped_page = max(1, request.args.get("group_page", 1, type=int))
    grouped_page_count = max(
        1,
        (len(grouped_screens) + page_size - 1) // page_size,
    )
    grouped_page = min(grouped_page, grouped_page_count)
    grouped_start = (grouped_page - 1) * page_size

    return render_template(
        TEMPLATE_DASH,
        images=all_images,
        image_groups=image_groups,
        library_images=imgs[image_start:image_start + image_page_size],
        screens=result[page_start:page_start + page_size],
        page=page,
        page_count=page_count,
        screen_total=len(result),
        online_total=sum(1 for screen in result if screen["online"]),
        screen_q=screen_q,
        image_q=image_q,
        image_page=image_page,
        image_page_count=image_page_count,
        grouped_screens=grouped_screens[grouped_start:grouped_start + page_size],
        grouped_page=grouped_page,
        grouped_page_count=grouped_page_count,
    )


@app.get("/api/screen-status")
@login_required
def screen_status():
    screens = []
    for _, screen in read("screens").iterrows():
        screen_id = str(screen["screen_id"])
        screens.append(
            {
                "screen_id": screen_id,
                "online": time.time() - ONLINE.get(screen_id, 0) < 45,
            }
        )
    return jsonify(
        {
            "screens": screens,
            "online_total": sum(1 for screen in screens if screen["online"]),
            "screen_total": len(screens),
        }
    )


@app.get("/screen/<int:screen_id>/content")
@login_required
def screen_content(screen_id):
    screens = read("screens")
    hit = screens[screens.screen_id.astype(str) == str(screen_id)]
    if hit.empty:
        return "Screen not found", 404
    current = hit.iloc[0].to_dict()
    images = read("images").to_dict("records")
    image_groups = sorted(
        {
            str(image.get("group_name", "")).strip()
            for image in images
            if str(image.get("group_name", "")).strip()
        },
        key=str.casefold,
    )
    selected_image = next(
        (image for image in images if str(image["id"]) == str(current["image_id"])),
        None,
    )
    return render_template(
        "screen_content.html",
        screen_id=screen_id,
        current=current,
        current_image=selected_image,
        images=images,
        image_groups=image_groups,
    )


@app.get("/deployment-groups")
@login_required
def deployment_groups():
    screens = read("screens").to_dict("records")
    images = read("images").to_dict("records")
    image_by_id = {str(image.get("id")): image for image in images}
    for screen in screens:
        screen["current_image"] = image_by_id.get(
            str(screen.get("image_id", "")),
            {},
        )
    server_groups = {}
    image_groups = {}

    for screen in screens:
        name = str(screen.get("group_name", "")).strip()
        if name:
            server_groups.setdefault(name, []).append(screen)

    for image in images:
        name = str(image.get("group_name", "")).strip()
        if name:
            image_groups.setdefault(name, []).append(image)

    def screen_sort_key(screen):
        value = str(screen.get("screen_id", ""))
        return (0, int(value)) if value.isdigit() else (1, value.casefold())

    for group in server_groups.values():
        group.sort(key=screen_sort_key)

    images_by_group = {
        name.casefold(): group_images
        for name, group_images in image_groups.items()
    }
    queues = read("queues")
    server_group_cards = []
    for name in sorted(server_groups, key=str.casefold):
        group_screens = server_groups[name]
        queue = []
        for _, item in queue_rows(queues, name).iterrows():
            queued_images = images_by_group.get(
                str(item["image_group"]).strip().casefold(),
                [],
            )
            queue.append(
                {
                    **item.to_dict(),
                    "image_count": len(queued_images),
                    "ready": len(queued_images) == len(group_screens),
                    "thumbnails": [
                        image["filename"] for image in queued_images[:4]
                    ],
                }
            )
        server_group_cards.append(
            {
                "name": name,
                "screens": group_screens,
                "active_group": next(
                    (
                        str(screen.get("playback_group_name", "")).strip()
                        for screen in group_screens
                        if str(screen.get("playback_group_name", "")).strip()
                    ),
                    "",
                ),
                "queue": queue,
            }
        )

    grouped_data = []
    selected_server_group = request.args.get("server_group", "").strip()
    selected_image_group = request.args.get("image_group", "").strip()
    selected_servers = server_groups.get(selected_server_group, [])
    selected_images = image_groups.get(selected_image_group, [])

    if selected_server_group and selected_image_group:
        grouped_data = list(zip(selected_servers, selected_images))

    return render_template(
        TEMPLATE_DEPLOYMENT_GROUPS,
        screens=screens,
        server_groups=server_group_cards,
        max_queue_items=MAX_QUEUE_ITEMS,
        image_groups=[
            {"name": name, "images": image_groups[name]}
            for name in sorted(image_groups, key=str.casefold)
        ],
        selected_server_group=selected_server_group,
        selected_image_group=selected_image_group,
        selected_servers=selected_servers,
        selected_images=selected_images,
        grouped_data=grouped_data,
        image_by_id=image_by_id,
    )


def name_matches(values, name):
    return values.astype(str).str.strip().str.casefold() == str(name).strip().casefold()


def queue_rows(queues, server_group):
    return queues.loc[name_matches(queues.server_group, server_group)]


def screen_order(values):
    return values.map(
        lambda value: (0, int(str(value)))
        if str(value).isdigit()
        else (1, str(value).casefold())
    )


def deploy_pairs(server_group, image_group):
    """Pair a server group with an image group by position and push the result.

    Returns (assignments, error). Nothing is written when an error is returned.
    """
    screens = read("screens")
    images = read("images")
    selected_screens = screens.loc[name_matches(screens.group_name, server_group)]
    selected_images = images.loc[name_matches(images.group_name, image_group)]

    if selected_screens.empty:
        return None, f"Server group {server_group} was not found."
    if selected_images.empty:
        return None, f"Image group {image_group} was not found."
    if len(selected_screens) != len(selected_images):
        return None, (
            f"{server_group} has {len(selected_screens)} servers but "
            f"{image_group} has {len(selected_images)} images. "
            "The counts must match."
        )

    selected_screens = selected_screens.sort_values(by="screen_id", key=screen_order)
    image_group = str(selected_images.iloc[0].group_name).strip()
    pairs = []
    for (screen_index, screen), (_, image) in zip(
        selected_screens.iterrows(),
        selected_images.iterrows(),
    ):
        screens.loc[screen_index, "image_id"] = str(image["id"])
        screens.loc[screen_index, "playback_mode"] = "paired"
        screens.loc[screen_index, "playback_group_name"] = image_group
        pairs.append((str(screen["screen_id"]), str(image["id"])))

    write("screens", screens)
    for screen_id, image_id in pairs:
        socketio.emit(
            "image_update",
            {"url": image_url(image_id), "index": 0},
            room=f"screen-{screen_id}",
        )
    return [f"{screen_id}:{image_id}" for screen_id, image_id in pairs], None


def back_to_deployments(server_group="", image_group="", anchor=""):
    target = url_for(
        "deployment_groups",
        **({"server_group": server_group} if server_group else {}),
        **({"image_group": image_group} if image_group else {}),
    )
    return redirect(target + (f"#{anchor}" if anchor else ""))


def group_anchor(server_group):
    return f"group-{server_group}" if server_group else ""


@app.post("/deployment-groups/servers")
@login_required
def group_servers():
    group_name = request.form.get("group_name", "").strip()
    selected_ids = {
        str(screen_id).strip()
        for screen_id in request.form.getlist("screen_ids")
        if str(screen_id).strip()
    }
    if not group_name or len(group_name) > 10 or not group_name.isalnum():
        flash("Server group names must be 1-10 letters or numbers.", "error")
        return back_to_deployments()
    if not selected_ids:
        flash("Select at least one server.", "error")
        return back_to_deployments()

    screens = read("screens")
    matches = screens.screen_id.astype(str).isin(selected_ids)
    if matches.sum() != len(selected_ids):
        flash("One or more selected servers were not found.", "error")
        return back_to_deployments()

    screens.loc[matches, "group_name"] = group_name
    write("screens", screens)
    audit("group_servers", f"{len(selected_ids)} servers to group {group_name}")
    flash(f"Saved {len(selected_ids)} server(s) to {group_name}.", "success")
    return back_to_deployments(server_group=group_name, anchor=group_anchor(group_name))


@app.post("/deployment-groups/delete-server-group")
@login_required
def delete_server_group():
    group_name = request.form.get("group_name", "").strip()
    screens = read("screens")
    matches = name_matches(screens.group_name, group_name)
    if not group_name or not matches.any():
        flash("Server group not found.", "error")
        return back_to_deployments()

    screens.loc[matches, "group_name"] = ""
    write("screens", screens)
    queues = read("queues")
    write("queues", queues.loc[~name_matches(queues.server_group, group_name)])
    audit("delete_server_group", group_name)
    flash(f"Deleted server group {group_name} and its queue.", "success")
    return back_to_deployments()


@app.post("/deployment-groups/deploy")
@login_required
def deploy_image_group():
    server_group = request.form.get("server_group", "").strip()
    image_group = request.form.get("image_group", "").strip()
    if not server_group or not image_group:
        flash("Select both a server group and an image group.", "error")
        return back_to_deployments(server_group, image_group)

    assignments, error = deploy_pairs(server_group, image_group)
    if error:
        flash(error, "error")
        return back_to_deployments(server_group, image_group)

    audit(
        "deploy_image_group",
        f"{server_group} <- {image_group}: {', '.join(assignments)}",
    )
    flash(f"{image_group} is now playing on {server_group}.", "success")
    return back_to_deployments(server_group, image_group, group_anchor(server_group))


@app.post("/deployment-groups/queue")
@login_required
def queue_image_group():
    server_group = request.form.get("server_group", "").strip()
    image_group = request.form.get("image_group", "").strip()
    if not server_group or not image_group:
        flash("Select both a server group and an image group to queue.", "error")
        return back_to_deployments(server_group)

    screens = read("screens")
    images = read("images")
    server_count = int(name_matches(screens.group_name, server_group).sum())
    image_count = int(name_matches(images.group_name, image_group).sum())
    queues = read("queues")
    queue_length = len(queue_rows(queues, server_group))
    error = (
        f"Server group {server_group} was not found." if not server_count
        else f"Image group {image_group} was not found." if not image_count
        else (
            f"{server_group} has {server_count} servers but {image_group} has "
            f"{image_count} images. The counts must match."
        ) if server_count != image_count
        else f"The queue for {server_group} is full ({MAX_QUEUE_ITEMS} items)."
        if queue_length >= MAX_QUEUE_ITEMS
        else ""
    )
    if error:
        flash(error, "error")
        return back_to_deployments(server_group, anchor=group_anchor(server_group))

    queues.loc[len(queues)] = [
        uuid.uuid4().hex,
        server_group,
        image_group,
        datetime.now().isoformat(timespec="seconds"),
        session.get("user", "system"),
    ]
    write("queues", queues.reset_index(drop=True))
    audit("queue_image_group", f"{server_group} <- {image_group}")
    flash(
        f"Added {image_group} to the {server_group} queue (position {queue_length + 1}).",
        "success",
    )
    return back_to_deployments(server_group, anchor=group_anchor(server_group))


@app.post("/deployment-groups/queue/move")
@login_required
def move_queue_item():
    item_id = request.form.get("item_id", "").strip()
    direction = request.form.get("direction", "")
    queues = read("queues")
    hit = queues.loc[queues.id.astype(str) == item_id]
    if hit.empty or direction not in {"up", "down", "top", "bottom"}:
        flash("That queue item no longer exists.", "error")
        return back_to_deployments()

    server_group = str(hit.iloc[0].server_group).strip()
    slots = list(queue_rows(queues, server_group).index)
    order = slots.copy()
    position = order.index(hit.index[0])
    moved = order.pop(position)
    new_position = {
        "up": max(0, position - 1),
        "down": min(len(order), position + 1),
        "top": 0,
        "bottom": len(order),
    }[direction]
    order.insert(new_position, moved)

    if order != slots:
        queues.loc[slots] = queues.loc[order].to_numpy()
        write("queues", queues)
        audit(
            "move_queue_item",
            f"{server_group}: {hit.iloc[0].image_group} {position + 1}->{new_position + 1}",
        )
    return back_to_deployments(server_group, anchor=group_anchor(server_group))


@app.post("/deployment-groups/queue/delete")
@login_required
def delete_queue_item():
    item_id = request.form.get("item_id", "").strip()
    queues = read("queues")
    match = queues.id.astype(str) == item_id
    if not match.any():
        flash("That queue item no longer exists.", "error")
        return back_to_deployments()

    item = queues.loc[match].iloc[0]
    server_group = str(item.server_group).strip()
    write("queues", queues.loc[~match])
    audit("delete_queue_item", f"{server_group}: {item.image_group}")
    flash(f"Removed {item.image_group} from the {server_group} queue.", "success")
    return back_to_deployments(server_group, anchor=group_anchor(server_group))


@app.post("/deployment-groups/queue/clear")
@login_required
def clear_queue():
    server_group = request.form.get("server_group", "").strip()
    queues = read("queues")
    match = name_matches(queues.server_group, server_group)
    if server_group and match.any():
        write("queues", queues.loc[~match])
        audit("clear_queue", f"{server_group}: {int(match.sum())} items")
        flash(f"Cleared the {server_group} queue.", "success")
    return back_to_deployments(server_group, anchor=group_anchor(server_group))


@app.post("/deployment-groups/deploy-next")
@login_required
def deploy_next_image_group():
    server_group = request.form.get("server_group", "").strip()
    return_to = (
        "dashboard" if request.form.get("return_to") == "dashboard"
        else "deployment_groups"
    )

    def finish():
        if return_to == "dashboard":
            return redirect(url_for("dashboard"))
        return back_to_deployments(server_group, anchor=group_anchor(server_group))

    queues = read("queues")
    pending = queue_rows(queues, server_group)
    if not server_group or pending.empty:
        flash(f"There is nothing queued for {server_group or 'this group'}.", "error")
        return finish()

    item = pending.iloc[0]
    image_group = str(item.image_group).strip()
    assignments, error = deploy_pairs(server_group, image_group)
    if error:
        flash(f"Could not deploy the next queued group: {error}", "error")
        return finish()

    write("queues", queues.loc[queues.id.astype(str) != str(item.id)])
    audit("deploy_next_image_group", f"{server_group} <- {image_group}: {', '.join(assignments)}")
    flash(f"{image_group} is now playing on {server_group}.", "success")
    return finish()


@app.get("/groups")
@login_required
def groups():
    images = read("images").to_dict("records")
    grouped = {}

    for image in images:
        name = str(image.get("group_name", "")).strip() or "Ungrouped"
        grouped.setdefault(name, []).append(image)

    groups_data = [
        {"name": name, "images": grouped[name]}
        for name in sorted(grouped, key=str.casefold)
    ]
    group_q = request.args.get("group_q", "").strip()
    if group_q:
        group_data_query = group_q.casefold()
        groups_data = [
            group for group in groups_data
            if group_data_query in group["name"].casefold()
        ]
    group_page_size = 10
    group_page = max(1, request.args.get("page", 1, type=int))
    group_page_count = max(
        1,
        (len(groups_data) + group_page_size - 1) // group_page_size,
    )
    group_page = min(group_page, group_page_count)
    group_start = (group_page - 1) * group_page_size
    selected = request.args.get("group", "").strip()
    selected_group = next(
        (
            group for group in
            [{"name": name, "images": grouped[name]} for name in grouped]
            if group["name"] == selected
        ),
        None,
    )

    return render_template(
        TEMPLATE_GROUPS,
        groups=groups_data[group_start:group_start + group_page_size],
        selected=selected,
        selected_group=selected_group,
        group_q=group_q,
        page=group_page,
        page_count=group_page_count,
    )


@app.get("/notifications")
@login_required
def notifications():
    query = request.args.get("q", "").strip()
    query_value = query.casefold()
    records = read("notifications").reset_index().rename(
        columns={"index": "row_index"}
    ).to_dict("records")
    records = [
        record for record in records
        if record.get("action") == "mark_done"
        and not str(record.get("dismissed_at", "")).strip()
        and (
            not query_value
            or query_value in " ".join(
                str(record.get(field, ""))
                for field in ("at", "username", "details")
            ).casefold()
        )
    ]
    page_size = 5
    page = max(1, request.args.get("page", 1, type=int))
    page_count = max(1, (len(records) + page_size - 1) // page_size)
    page = min(page, page_count)
    start = (page - 1) * page_size

    return render_template(
        TEMPLATE_NOTIFICATIONS,
        notifications=records[start:start + page_size],
        query=query,
        page=page,
        page_count=page_count,
    )


@app.post("/notifications/<int:notification_index>/delete")
@login_required
def delete_notification(notification_index):
    notifications = read("notifications")
    if notification_index < 0 or notification_index >= len(notifications):
        return "Notification not found", 404

    row = notifications.iloc[notification_index]
    if str(row.get("action", "")) != "mark_done":
        return "Notification not found", 404

    notifications.loc[notification_index, "dismissed_at"] = datetime.now().isoformat(
        timespec="seconds"
    )
    write("notifications", notifications)

    return redirect(url_for("notifications"))


@app.get("/fallbacks")
@login_required
def fallbacks():
    images = {
        str(image["id"]): image
        for image in read("images").to_dict("records")
    }
    screens = []

    for screen in read("screens").to_dict("records"):
        image = images.get(str(screen["image_id"]), {})
        screens.append(
            {
                **screen,
                "filename": image.get("display_name", ""),
            }
        )

    query = request.args.get("q", "").strip()
    if query:
        query_value = query.casefold()
        screens = [
            screen for screen in screens
            if query_value in " ".join(
                str(screen.get(field, ""))
                for field in ("screen_id", "filename")
            ).casefold()
        ]

    page_size = 5
    page = max(1, request.args.get("page", 1, type=int))
    page_count = max(1, (len(screens) + page_size - 1) // page_size)
    page = min(page, page_count)
    page_start = (page - 1) * page_size

    return render_template(
        TEMPLATE_FALLBACKS,
        screens=screens[page_start:page_start + page_size],
        query=query,
        page=page,
        page_count=page_count,
    )


@app.post("/fallbacks")
@login_required
def set_fallbacks():
    selected_ids = {
        str(screen_id)
        for screen_id in request.form.getlist("screen_ids")
    }
    screens = read("screens")
    match = screens.screen_id.astype(str).isin(selected_ids)

    if match.any():
        screens.loc[match, "image_id"] = ""
        write("screens", screens)
        for screen_id in selected_ids:
            socketio.emit(
                "image_update",
                {"url": image_url("")},
                room=f"screen-{screen_id}",
            )
        audit("set_fallbacks", ",".join(sorted(selected_ids)))

    return redirect(url_for("fallbacks"))


def store_image(file, group_name=""):
    ext = Path(
        secure_filename(
            file.filename
            if file else ""
        )
    ).suffix.lower().lstrip(".")

    if not file or ext not in ALLOWED or upload_size(file) > MAX_IMAGE_BYTES:
        return None

    filename = (
        f"{uuid.uuid4().hex}.{ext}"
    )

    image_path = UPLOADS / filename
    temporary_path = UPLOADS / f".{filename}.{uuid.uuid4().hex}.tmp"
    try:
        file.save(temporary_path)
        os.replace(temporary_path, image_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()

    return [
        uuid.uuid4().hex,
        filename,
        Path(file.filename).stem[:120],
        group_name,
        datetime.now().isoformat(),
    ]


def upload_size(file):
    if not file or not file.stream:
        return 0
    position = file.stream.tell()
    file.stream.seek(0, os.SEEK_END)
    size = file.stream.tell()
    file.stream.seek(position)
    return size


@app.post("/upload")
@login_required
def upload():

    file = request.files.get("image")
    record = store_image(file)
    if record is None:
        return "Invalid image", 400

    images = read("images")
    images.loc[len(images)] = record

    write(
        "images",
        images,
    )

    audit(
        "upload",
        record[1],
    )

    return redirect("/")


@app.post("/upload-multiple")
@login_required
def upload_multiple():

    files = request.files.getlist("images")
    add_to_group = request.form.get("add_to_group") == "yes"
    group_name = request.form.get("group_name", "").strip()

    if add_to_group and not group_name:
        return "A group name is required when grouping uploads", 400

    if not files:
        return "Choose at least one image", 400
    if len(files) > MAX_UPLOAD_IMAGES:
        return f"An upload can contain at most {MAX_UPLOAD_IMAGES} images", 400

    for file in files:
        ext = Path(secure_filename(file.filename if file else "")).suffix.lower().lstrip(".")
        if ext not in ALLOWED:
            return "All uploads must be valid image files", 400
        if upload_size(file) > MAX_IMAGE_BYTES:
            return "Each image must be 5 MB or smaller", 400

    images = read("images")
    if add_to_group:
        existing_count = (
            images["group_name"].astype(str).str.strip().str.casefold()
            == group_name.casefold()
        ).sum()
        if existing_count + len(files) > MAX_GROUP_IMAGES:
            return f"An image group can contain at most {MAX_GROUP_IMAGES} images", 400

    records = []
    for file in files:
        record = store_image(file, group_name if add_to_group else "")
        records.append(record)

    for record in records:
        images.loc[len(images)] = record
    write("images", images)
    audit(
        "upload_multiple",
        f"{len(records)} images" + (f" to group {group_name}" if add_to_group else ""),
    )

    return redirect(url_for("groups" if add_to_group else "dashboard"))


@app.post("/update-image/<image_id>")
@login_required
def update_image(image_id):

    images = read("images")
    match = images.id.astype(str) == str(image_id)

    if not match.any():
        return "Image not found", 404

    display_name = request.form.get("display_name", "").strip()
    if not display_name:
        return "Image name is required", 400

    images.loc[match, "display_name"] = display_name[:120]
    write("images", images)
    audit(
        "update_image",
        f"{image_id}:{display_name[:120]}",
    )

    return redirect("/")


@app.post("/group-images")
@login_required
def group_images():
    image_ids = {str(image_id).strip() for image_id in request.form.getlist("image_ids")}
    group_name = request.form.get("group_name", "").strip()[:80]
    if not image_ids:
        return "Select at least one image", 400
    if not group_name:
        return "A group name is required", 400

    images = read("images")
    selected = images.id.astype(str).isin(image_ids)
    if selected.sum() != len(image_ids):
        return "One or more selected images were not found", 404

    group_count = (
        images["group_name"].astype(str).str.strip().str.casefold()
        == group_name.casefold()
    ).sum()
    moving_into_group = (
        images.loc[selected, "group_name"].astype(str).str.strip().str.casefold()
        != group_name.casefold()
    ).sum()
    if group_count + moving_into_group > MAX_GROUP_IMAGES:
        return f"An image group can contain at most {MAX_GROUP_IMAGES} images", 400

    images.loc[selected, "group_name"] = group_name
    write("images", images)
    audit("group_images", f"{len(image_ids)} images to group {group_name}")
    return redirect(url_for("dashboard"))


@app.post("/rename-image-group")
@login_required
def rename_image_group():
    old_name = request.form.get("old_name", "").strip()
    new_name = request.form.get("new_name", "").strip()[:80]
    if not old_name or not new_name:
        return "Both current and new group names are required", 400

    images = read("images")
    old_match = (
        images["group_name"].astype(str).str.strip().str.casefold()
        == old_name.casefold()
    )
    if not old_match.any():
        return "Image group not found", 404

    new_match = (
        images["group_name"].astype(str).str.strip().str.casefold()
        == new_name.casefold()
    )
    if new_match.any() and not (old_name.casefold() == new_name.casefold()):
        return "An image group with that name already exists", 400
    images.loc[old_match, "group_name"] = new_name
    write("images", images)
    queues = read("queues")
    queued = name_matches(queues.image_group, old_name)
    if queued.any():
        queues.loc[queued, "image_group"] = new_name
        write("queues", queues)
    audit("rename_image_group", f"{old_name}:{new_name}")
    return redirect(url_for("groups", group=new_name))


@app.post("/delete-image-group")
@login_required
def delete_image_group():
    group_name = request.form.get("group_name", "").strip()
    if not group_name:
        return "Image group name is required", 400

    images = read("images")
    group_match = (
        images["group_name"].astype(str).str.strip().str.casefold()
        == group_name.casefold()
    )
    if not group_match.any():
        return "Image group not found", 404

    group_image_ids = set(images.loc[group_match, "id"].astype(str))
    images.loc[group_match, "group_name"] = ""
    write("images", images)
    queues = read("queues")
    queued = name_matches(queues.image_group, group_name)
    if queued.any():
        write("queues", queues.loc[~queued])

    screens = read("screens")
    affected = screens.image_id.astype(str).isin(group_image_ids)
    if affected.any():
        screens.loc[affected, "image_id"] = ""
        screens.loc[affected, "playback_mode"] = ""
        write("screens", screens)
        for screen_id in screens.loc[affected, "screen_id"]:
            socketio.emit(
                "image_update",
                {"url": image_url("")},
                room=f"screen-{screen_id}",
            )

    audit("delete_image_group", group_name)
    return redirect(url_for("groups"))


@app.post("/delete-screen/<int:screen_id>")
@login_required
def delete_screen(screen_id):

    screens = read("screens")
    match = screens.screen_id.astype(str) == str(screen_id)

    if not match.any():
        return "Screen not found", 404

    screens = screens.loc[~match].copy()
    write("screens", screens)
    ONLINE.pop(str(screen_id), None)
    LAST_HEARTBEAT_WRITE.pop(str(screen_id), None)
    audit("delete_screen", str(screen_id))

    return redirect(url_for("dashboard"))


@app.post("/update-screen/<int:screen_id>")
@login_required
def update_screen(screen_id):

    group_name = request.form.get("group_name", "").strip()
    if group_name and (
        len(group_name) > 10 or not group_name.isalnum()
    ):
        return "Server group must be 1-10 letters or numbers", 400

    screens = read("screens")
    match = screens.screen_id.astype(str) == str(screen_id)
    if not match.any():
        return "Screen not found", 404

    screens.loc[match, "group_name"] = group_name
    write("screens", screens)
    audit("update_screen_group", f"{screen_id}:{group_name}")

    return redirect(url_for("dashboard"))


@app.post("/assign-group/<int:screen_id>")
@login_required
def assign_group(screen_id):

    group_name = request.form.get("group_name", "").strip()
    images = read("images")
    group_images = images[
        images.group_name.astype(str).str.strip() == group_name
    ]
    if not group_name or group_images.empty:
        return "Image group not found", 404

    screens = read("screens")
    match = screens.screen_id.astype(str) == str(screen_id)
    if not match.any():
        return "Screen not found", 404

    first_image_id = str(group_images.iloc[0].id)
    screens.loc[match, "image_id"] = first_image_id
    screens.loc[match, "playback_mode"] = ""
    screens.loc[match, "playback_group_name"] = ""
    screens.loc[match, "queued_group_name"] = ""
    write("screens", screens)
    socketio.emit(
        "image_update",
        {"url": image_url(first_image_id), "index": 0},
        room=f"screen-{screen_id}",
    )
    audit("assign_group", f"{screen_id}:{group_name}")

    return redirect(url_for("dashboard"))


@app.post("/delete-image/<image_id>")
@login_required
def delete_image(image_id):

    images = read("images")

    hit = images[
        images.id.astype(str)
        ==
        str(image_id)
    ]

    if hit.empty:
        return "Image not found", 404

    filename = hit.iloc[0].filename

    images = images[
        images.id.astype(str)
        !=
        str(image_id)
    ]

    write("images", images)

    # Only screens showing the deleted image fall back. Their server group
    # (screens.group_name) is unrelated to image groups and is left alone.
    screens = read("screens")
    assigned = screens.image_id.astype(str) == str(image_id)
    if assigned.any():
        screens.loc[
            assigned,
            ["image_id", "playback_mode", "playback_group_name"],
        ] = ""
        write("screens", screens)
        for screen_id in screens.loc[assigned, "screen_id"]:
            socketio.emit(
                "image_update",
                {"url": image_url("")},
                room=f"screen-{screen_id}",
            )

    image_path = UPLOADS / filename
    if image_path.is_file():
        image_path.unlink()

    audit(
        "delete_image",
        filename,
    )

    return redirect("/")


@app.post("/delete-images")
@login_required
def delete_images():
    image_ids = {str(image_id).strip() for image_id in request.form.getlist("image_ids")}
    if not image_ids:
        return "Select at least one image", 400

    images = read("images")
    selected = images.id.astype(str).isin(image_ids)
    if not selected.any():
        return "Selected images were not found", 404

    selected_records = images.loc[selected].copy()
    images = images.loc[~selected].copy()
    write("images", images)

    screens = read("screens")
    affected = screens.image_id.astype(str).isin(image_ids)
    if affected.any():
        screens.loc[affected, "image_id"] = ""
        screens.loc[affected, "playback_mode"] = ""
        screens.loc[affected, "playback_group_name"] = ""
        write("screens", screens)
        for screen_id in screens.loc[affected, "screen_id"]:
            socketio.emit(
                "image_update",
                {"url": image_url("")},
                room=f"screen-{screen_id}",
            )

    for filename in selected_records["filename"]:
        image_path = UPLOADS / str(filename)
        if image_path.is_file():
            image_path.unlink()

    audit("delete_images", f"{len(selected_records)} images")
    return redirect(url_for("dashboard"))


@app.get("/screen/<int:screen_id>")
def screen(screen_id):
    screens = read("screens")
    hit = screens[screens.screen_id.astype(str) == str(screen_id)]
    current_image_id = str(hit.iloc[0].image_id) if not hit.empty else ""
    current_url = image_url(current_image_id)
    images = read("images")
    current = images[images.id.astype(str) == current_image_id]
    group_name = str(current.iloc[0].group_name).strip() if not current.empty else ""
    paired = (
        str(hit.iloc[0].get("playback_mode", "")).strip() == "paired"
        if not hit.empty
        else False
    )
    playlist_rows = (
        images[
            images.group_name.astype(str).str.strip() == group_name
        ]
        if group_name and not paired
        else images.iloc[0:0]
    )
    playlist = [
        {
            "id": str(row.id),
            "url": image_url(row.id),
            "name": str(row.display_name),
        }
        for _, row in playlist_rows.iterrows()
    ]
    current_index = next(
        (
            index for index, item in enumerate(playlist)
            if item["id"] == current_image_id
        ),
        0,
    )
    return render_template(
        TEMPLATE_SCREEN,
        sid=screen_id,
        image_url=current_url,
        playlist=playlist,
        current_index=current_index,
    )


@app.get("/api/screen/<int:screen_id>")
def screen_image(screen_id):
    screens = read("screens")
    hit = screens[screens.screen_id.astype(str) == str(screen_id)]
    image_id = hit.iloc[0].image_id if not hit.empty else ""
    return jsonify(url=image_url(image_id))


@app.post("/api/push")
@login_required
def push():

    sid = str(
        request.form.get(
            "screen_id"
        )
    )

    iid = str(
        request.form.get(
            "image_id",
            "",
        )
    )

    screens = read("screens")

    match = (
        screens.screen_id
        .astype(str)
        == sid
    )

    if match.any():
        screens.loc[
            match,
            "image_id"
        ] = iid
        screens.loc[match, "playback_mode"] = ""
        screens.loc[match, "playback_group_name"] = ""
        screens.loc[match, "queued_group_name"] = ""
    else:
        screens.loc[
            len(screens)
        ] = [
            sid,
            iid,
            "",
            "",
            "",
            "",
            datetime.now().isoformat(),
        ]

    write(
        "screens",
        screens,
    )

    url = image_url(iid)

    socketio.emit(
        "image_update",
        {"url": url},
        room=f"screen-{sid}",
    )

    audit(
        "assign",
        f"{sid}:{iid}",
    )

    return redirect("/")


@app.post("/api/push-all")
@login_required
def push_all():

    image_id = str(
        request.form.get(
            "image_id",
            "",
        )
    )

    if image_id and read("images")[lambda images: images.id.astype(str) == image_id].empty:
        return "Image not found", 404

    screens = read("screens")

    if not screens.empty:
        screens["image_id"] = image_id
        screens["playback_mode"] = ""
        screens["playback_group_name"] = ""
        screens["queued_group_name"] = ""
        write(
            "screens",
            screens,
        )

    url = image_url(image_id)

    socketio.emit(
        "image_update",
        {"url": url},
    )

    audit(
        "push_all",
        image_id,
    )

    return redirect("/")


@app.get("/api/screens")
@login_required
def api_screens():
    return jsonify(
        {
            k: time.time() - v < 45
            for k, v in ONLINE.items()
        }
    )


@app.get("/health")
def health():
    return jsonify(
        status="ok",
        time=datetime.now().isoformat(),
    )


@socketio.on("register")
def register(data):

    sid = str(
        data.get("screen_id")
    )

    ONLINE[sid] = time.time()

    join_room(
        f"screen-{sid}"
    )

    screens = read(
        "screens"
    )

    exists = (
        screens.screen_id
        .astype(str)
        == sid
    ).any()

    if not exists:

        screens.loc[
            len(screens)
        ] = [
            sid,
            "",
            "",
            "",
            "",
            "",
            datetime.now().isoformat(),
        ]

        write(
            "screens",
            screens
        )
        audit(
            "screen_registered",
            sid,
        )

    screens = read(
        "screens"
    )

    hit = screens[
        screens.screen_id
        .astype(str)
        ==
        sid
    ]

    image_id = str(hit.iloc[0].image_id) if not hit.empty else ""
    url = image_url(image_id)

    emit(
        "image_update",
        {"url": url},
    )


@socketio.on("join_dashboard")
def join_dashboard():
    join_room("dashboard")


@socketio.on("heartbeat")
def heartbeat(data):

    sid = str(
        data.get(
            "screen_id"
        )
    )
    now = time.time()
    ONLINE[sid] = now

    if now - LAST_HEARTBEAT_WRITE.get(sid, 0) < HEARTBEAT_WRITE_INTERVAL:
        return

    screens = read("screens")
    match = screens.screen_id.astype(str) == sid
    if match.any():
        screens.loc[match, "last_seen"] = datetime.now().isoformat(
            timespec="seconds"
        )
        write("screens", screens)
        LAST_HEARTBEAT_WRITE[sid] = now


@socketio.on("navigate_image")
def navigate_image(data):
    sid = str(data.get("screen_id"))
    direction = data.get("direction")
    screens = read("screens")
    screen_match = screens.screen_id.astype(str) == sid
    if not screen_match.any() or direction not in {"next", "previous"}:
        return
    # Paired screens show only their assigned image, never the group playlist.
    if str(screens.loc[screen_match, "playback_mode"].iloc[0]).strip() == "paired":
        return

    current_id = str(screens.loc[screen_match, "image_id"].iloc[0])
    images = read("images")
    current = images[images.id.astype(str) == current_id]
    if current.empty:
        return
    group_name = str(current.iloc[0].group_name).strip()
    if not group_name:
        return

    playlist = images[
        images.group_name.astype(str).str.strip() == group_name
    ].reset_index(drop=True)
    current_positions = playlist.index[
        playlist.id.astype(str) == current_id
    ].tolist()
    if not current_positions:
        return

    next_index = current_positions[0] + (1 if direction == "next" else -1)
    if next_index < 0 or next_index >= len(playlist):
        return

    next_id = str(playlist.iloc[next_index].id)
    screens.loc[screen_match, "image_id"] = next_id
    write("screens", screens)
    socketio.emit(
        "image_update",
        {
            "url": image_url(next_id),
            "index": next_index,
        },
        room=f"screen-{sid}",
    )
    audit("navigate_image", f"{sid}:{direction}:{next_id}")


@socketio.on("mark_done")
def mark_done(data):
    sid = str(data.get("screen_id"))
    details = f"Screen {sid} completed its image group"
    audit("mark_done", details)
    create_notification("mark_done", details)
    socketio.emit(
        "notification",
        {"message": f"Screen {sid} marked its image group as done"},
        room="dashboard",
    )


def cleanup_online():

    while True:

        now = time.time()

        stale = [
            sid
            for sid, ts
            in ONLINE.items()
            if now - ts > 120
        ]

        for sid in stale:
            ONLINE.pop(
                sid,
                None,
            )

        time.sleep(30)


if __name__ == "__main__":
    init_book()

    threading.Thread(
        target=cleanup_online,
        daemon=True,
    ).start()

    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 3000))

    development = os.getenv("DEVELOPMENT", "").casefold() in {
        "1",
        "true",
        "yes",
    }
    production = not development and os.getenv(
        "PRODUCTION",
        "true",
    ).casefold() in {
        "1",
        "true",
        "yes",
    }

    if production:
        from waitress import serve

        socketio.server.eio.allow_upgrades = False

        print(f"Starting Waitress on {host}:{port}")

        serve(
            app.wsgi_app,
            host=host,
            port=port,
            threads=int(os.getenv("WAITRESS_THREADS", "8")),
        )
    else:
        print(f"Starting Flask development server on {host}:{port}")

        socketio.run(
            app,
            host=host,
            port=port,
            allow_unsafe_werkzeug=True,
        )