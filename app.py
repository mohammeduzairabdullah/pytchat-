import os
import re
from datetime import datetime
from urllib.parse import parse_qs, urlsplit

import pandas as pd
import pytchat
from flask import Flask, jsonify, render_template, request, send_file

app = Flask(__name__)


def extract_video_id(url: str) -> str:
    """Extract a clean 11-character YouTube video ID from common YouTube URL formats."""
    cleaned_url = (url or "").strip()
    if not cleaned_url:
        raise ValueError(f"Could not extract video ID from URL: {url}")

    if not re.match(r"^[a-zA-Z][a-zA-Z0-9+\-.]*://", cleaned_url):
        cleaned_url = f"https://{cleaned_url}"

    def valid_video_id(candidate: str) -> bool:
        return bool(re.fullmatch(r"[0-9A-Za-z_-]{11}", candidate or ""))

    parsed = urlsplit(cleaned_url)
    hostname = (parsed.netloc or "").lower()
    path = parsed.path or ""
    query = parse_qs(parsed.query)

    candidate = None

    if "youtube.com" in hostname:
        if path.startswith("/watch"):
            v_values = query.get("v", [])
            if v_values:
                candidate = v_values[0]

        path_parts = [part for part in path.split("/") if part]
        if candidate is None and len(path_parts) >= 2:
            if path_parts[0] in {"live", "embed", "shorts"}:
                candidate = path_parts[1]

    if candidate is None and "youtu.be" in hostname:
        path_parts = [part for part in path.split("/") if part]
        if path_parts:
            candidate = path_parts[0]

    if valid_video_id(candidate):
        print(f"Extracted video ID: {candidate}")
        return candidate

    # Regex fallbacks (required order and patterns)
    patterns = [
        r"(?:youtube\.com\/watch\?v=)([0-9A-Za-z_-]{11})",
        r"(?:youtube\.com\/live\/)([0-9A-Za-z_-]{11})",
        r"(?:youtu\.be\/)([0-9A-Za-z_-]{11})",
        r"(?:youtube\.com\/embed\/)([0-9A-Za-z_-]{11})",
        r"(?:m\.youtube\.com\/watch\?v=)([0-9A-Za-z_-]{11})",
        r"(?:youtube\.com\/shorts\/)([0-9A-Za-z_-]{11})",
    ]

    for pattern in patterns:
        match = re.search(pattern, cleaned_url)
        if match:
            candidate = match.group(1)
            if valid_video_id(candidate):
                print(f"Extracted video ID: {candidate}")
                return candidate

    raise ValueError(f"Could not extract video ID from URL: {url}")


def fetch_all_comments(video_id: str) -> tuple[list[dict] | None, str | None]:
    """Extract all available live chat replay comments from an ended YouTube live stream."""
    try:
        chat = pytchat.create(video_id=video_id)
        comments = []
        count = 0

        while chat.is_alive():
            for comment in chat.get().sync_items():
                comments.append(
                    {
                        "No": count + 1,
                        "Username": comment.author.name,
                        "Comment": comment.message,
                        "Timestamp": comment.datetime,
                        "Channel ID": comment.author.channelId,
                    }
                )
                count += 1
                if count % 500 == 0:
                    print(f"Extracted {count} comments so far...")

        try:
            chat.raise_for_status()
        except Exception as status_error:
            return None, str(status_error)

        if not comments:
            return None, "No comments found. Chat may be disabled or replay unavailable."

        return comments, None
    except Exception as extraction_error:
        return None, str(extraction_error)


def export_to_excel(
    comments: list[dict], video_id: str
) -> tuple[str | None, str | None, str | None]:
    """Export comments to an Excel file under downloads/."""
    try:
        os.makedirs("downloads", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"comments_{video_id}_{timestamp}.xlsx"
        filepath = os.path.join("downloads", filename)

        columns = ["No", "Username", "Comment", "Timestamp", "Channel ID"]
        df = pd.DataFrame(comments, columns=columns)
        df.to_excel(filepath, engine="openpyxl", index=False)

        return filepath, filename, None
    except Exception as export_error:
        return None, None, str(export_error)


def run_url_parsing_self_test() -> None:
    """Small URL parsing self-test for expected YouTube formats."""
    tests = {
        "https://www.youtube.com/live/4uE-OW3Jmh0?si=3g4bhI0b5b-LJWRn": "4uE-OW3Jmh0",
        "https://www.youtube.com/live/4uE-OW3Jmh0": "4uE-OW3Jmh0",
        "https://www.youtube.com/watch?v=4uE-OW3Jmh0": "4uE-OW3Jmh0",
        "https://www.youtube.com/watch?v=4uE-OW3Jmh0&feature=share": "4uE-OW3Jmh0",
        "https://youtu.be/4uE-OW3Jmh0": "4uE-OW3Jmh0",
        "https://youtu.be/4uE-OW3Jmh0?si=abc123": "4uE-OW3Jmh0",
        "https://www.youtube.com/embed/4uE-OW3Jmh0": "4uE-OW3Jmh0",
        "https://m.youtube.com/watch?v=4uE-OW3Jmh0": "4uE-OW3Jmh0",
        "https://www.youtube.com/shorts/4uE-OW3Jmh0": "4uE-OW3Jmh0",
    }

    for test_url, expected_id in tests.items():
        actual_id = extract_video_id(test_url)
        assert actual_id == expected_id, f"Expected {expected_id}, got {actual_id} for {test_url}"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/extract", methods=["POST"])
def extract_comments_route():
    data = request.get_json(silent=True)
    if data is None:
        return jsonify({"error": "No data received"}), 400

    url = (data.get("url") or "").strip()
    if not url:
        return jsonify({"error": "Please provide a YouTube URL"}), 400

    lowered_url = url.lower()
    if "youtube.com" not in lowered_url and "youtu.be" not in lowered_url:
        return jsonify({"error": "Please enter a valid YouTube URL"}), 400

    try:
        video_id = extract_video_id(url)
    except ValueError:
        return jsonify({"error": f"Could not extract video ID from URL: {url}"}), 400

    comments, extraction_error = fetch_all_comments(video_id)
    if extraction_error:
        if extraction_error == "No comments found. Chat may be disabled or replay unavailable.":
            return jsonify({"error": "No comments found. Chat may be disabled."}), 404
        return jsonify({"error": f"Extraction failed: {extraction_error}"}), 500

    filepath, filename, export_error = export_to_excel(comments, video_id)
    if export_error:
        return jsonify({"error": f"Excel export failed: {export_error}"}), 500

    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/debug", methods=["POST"])
def debug_route():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()

    try:
        video_id = extract_video_id(url)
        return jsonify({"url": url, "video_id": video_id, "valid": True})
    except Exception as error:
        return jsonify({"url": url, "video_id": None, "valid": False, "error": str(error)})


if __name__ == "__main__":
    run_url_parsing_self_test()
    app.run(debug=True, threaded=True, port=5000)
