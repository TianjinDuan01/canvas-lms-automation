"""
Canvas Discussion Grader
========================
Automatically checks each student's post and reply activity across
one or more Canvas discussion topics and exports a grading CSV.

Usage:
  python3 canvas_discussion_grader.py

You will be prompted for your API token, course ID, and discussion IDs.
"""

import requests
import csv
from datetime import datetime, timezone

# ── Pre-filled defaults (edit here to skip the prompts) ──────────────────────
CANVAS_BASE_URL = "https://jhu.instructure.com"
API_TOKEN       = ""  # Leave blank to be prompted
COURSE_ID       = "YOUR_COURSE_ID"
DISCUSSION_IDS  = ["ID1", "ID2", "ID3"]
# ─────────────────────────────────────────────────────────────────────────────

# Roles to exclude from grading (instructors, TAs, designers)
SKIP_ROLES = {"TeacherEnrollment", "TaEnrollment", "DesignerEnrollment"}

# Assignment deadline (Eastern Time = UTC-4 in summer)
# Please adjust as needed
DEADLINE = datetime(2026, 5, 31, 23, 59, 0, tzinfo=timezone.utc).replace(
    tzinfo=None
)  # Canvas timestamps are in UTC; ET (UTC-4) 11:59pm = UTC 03:59 June 1
DEADLINE_UTC = datetime(2026, 6, 1, 3, 59, 0, tzinfo=timezone.utc)


def prompt(message, default=""):
    value = input(message).strip()
    return value if value else default


def fetch_all(url, headers, params=None):
    """Fetch all pages from a paginated Canvas API endpoint."""
    results = []
    params = {**(params or {}), "per_page": 100}
    while url:
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        results.extend(resp.json())
        url = None
        for part in resp.headers.get("Link", "").split(","):
            if 'rel="next"' in part:
                url = part[part.index("<") + 1: part.index(">")]
                params = {}
                break
    return results


def get_enrollment_roles(base_url, headers, course_id):
    """Return a dict of {user_id: enrollment_type} for everyone in the course."""
    url = f"{base_url}/api/v1/courses/{course_id}/enrollments"
    enrollments = fetch_all(url, headers)
    return {str(e["user_id"]): e["type"] for e in enrollments}


def fetch_discussion_entries(base_url, headers, course_id, discussion_id):
    """
    Pull all top-level posts and their replies for a single discussion topic.

    Returns:
        top_authors  (set of user_id strings) — students who made an original post
        reply_authors (set of user_id strings) — students who replied to a peer
        participants  (dict of user_id -> display_name)
    """
    url = f"{base_url}/api/v1/courses/{course_id}/discussion_topics/{discussion_id}/entries"
    top_entries = fetch_all(url, headers)

    participants = {}
    top_authors = set()
    reply_authors = set()
    post_times = {}   # uid -> latest post timestamp
    reply_times = {}  # uid -> latest reply timestamp

    def parse_ts(ts_str):
        """Parse Canvas ISO timestamp to UTC datetime."""
        if not ts_str:
            return None
        try:
            return datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        except Exception:
            return None

    for entry in top_entries:
        uid = str(entry["user_id"])
        participants[uid] = entry.get("user_name", f"User {uid}")
        top_authors.add(uid)
        ts = parse_ts(entry.get("created_at"))
        if uid not in post_times or (ts and ts > post_times.get(uid, datetime.min.replace(tzinfo=timezone.utc))):
            post_times[uid] = ts

        replies_url = (
            f"{base_url}/api/v1/courses/{course_id}"
            f"/discussion_topics/{discussion_id}/entries/{entry['id']}/replies"
        )
        for reply in fetch_all(replies_url, headers):
            ruid = str(reply["user_id"])
            participants[ruid] = reply.get("user_name", f"User {ruid}")
            reply_authors.add(ruid)
            ts = parse_ts(reply.get("created_at"))
            if ruid not in reply_times or (ts and ts > reply_times.get(ruid, datetime.min.replace(tzinfo=timezone.utc))):
                reply_times[ruid] = ts

    return top_authors, reply_authors, participants, post_times, reply_times


def run(base_url, token, course_id, discussion_ids):
    headers = {"Authorization": f"Bearer {token}"}

    print("\nFetching course enrollment data...")
    role_map = get_enrollment_roles(base_url, headers, course_id)

    # student_data: {user_id: {"posts": [bool, ...], "replies": [bool, ...]}}
    student_data = {}
    all_participants = {}

    for idx, did in enumerate(discussion_ids, 1):
        print(f"Processing Discussion {idx} (ID: {did})...")
        try:
            top_authors, reply_authors, participants, post_times, reply_times = fetch_discussion_entries(
                base_url, headers, course_id, did
            )
            all_participants.update(participants)

            for uid in top_authors | reply_authors:
                if uid not in student_data:
                    student_data[uid] = {"posts": [], "replies": [], "post_late": [], "reply_late": []}
                has_post  = uid in top_authors
                has_reply = uid in reply_authors
                post_ts   = post_times.get(uid)
                reply_ts  = reply_times.get(uid)
                post_late  = has_post  and post_ts  is not None and post_ts  > DEADLINE_UTC
                reply_late = has_reply and reply_ts is not None and reply_ts > DEADLINE_UTC
                student_data[uid]["posts"].append(has_post)
                student_data[uid]["replies"].append(has_reply)
                student_data[uid]["post_late"].append(post_late)
                student_data[uid]["reply_late"].append(reply_late)

        except requests.HTTPError as e:
            print(f"  Warning: could not fetch Discussion {did} — {e}")

    # Build CSV
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"discussion_grades_{timestamp}.csv"

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        max_score = len(discussion_ids) * 2  # 3 rounds x 2 points = 6
        header = ["Student Name", "Canvas User ID", "Role"]
        for i in range(1, len(discussion_ids) + 1):
            header += [f"D{i} Post", f"D{i} Post Late", f"D{i} Reply", f"D{i} Reply Late"]
        header += ["Total Posts", "Total Replies", f"Score (out of {max_score})", "Any Late"]
        writer.writerow(header)

        for uid, data in sorted(student_data.items()):
            role = role_map.get(uid, "Unknown")
            if role in SKIP_ROLES:
                continue

            name = all_participants.get(uid, f"User {uid}")
            row = [name, uid, role]
            total_posts   = 0
            total_replies = 0
            any_late      = False

            for i in range(len(discussion_ids)):
                has_post   = data["posts"][i]      if i < len(data["posts"])      else False
                has_reply  = data["replies"][i]    if i < len(data["replies"])    else False
                post_late  = data["post_late"][i]  if i < len(data["post_late"])  else False
                reply_late = data["reply_late"][i] if i < len(data["reply_late"]) else False
                if has_post:
                    total_posts += 1
                if has_reply:
                    total_replies += 1
                if post_late or reply_late:
                    any_late = True
                row += [
                    "Yes" if has_post   else "No",
                    "LATE" if post_late  else "",
                    "Yes" if has_reply  else "No",
                    "LATE" if reply_late else "",
                ]

            score = total_posts + total_replies
            row += [total_posts, total_replies, score, "LATE" if any_late else ""]
            writer.writerow(row)

    print(f"\nDone! Results saved to: {output_file}")


def main():
    global API_TOKEN, DISCUSSION_IDS

    print("=" * 50)
    print("  Canvas Discussion Grader")
    print("=" * 50)
    print(f"  Canvas:    {CANVAS_BASE_URL}")
    print(f"  Course ID: {COURSE_ID}")

    if not API_TOKEN:
        API_TOKEN = prompt("\nEnter your Canvas API token: ")
    if not DISCUSSION_IDS:
        raw = prompt("Enter Discussion IDs separated by commas (e.g. 1266816,111,222): ")
        DISCUSSION_IDS = [d.strip() for d in raw.split(",") if d.strip()]

    run(CANVAS_BASE_URL, API_TOKEN, COURSE_ID, DISCUSSION_IDS)


if __name__ == "__main__":
    main()
