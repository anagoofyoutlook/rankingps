import json
import csv
import os
import shutil
from datetime import datetime
import re
import zipfile
import random
from html import escape

# Define folder paths
input_folder = 'PS'
output_folder = 'docs'
html_subfolder = os.path.join(output_folder, 'HTML')
photos_folder = 'Photos'
docs_photos_folder = os.path.join(output_folder, 'Photos')
history_csv_file = os.path.join(output_folder, 'history.csv')

# Define CSV output path
csv_file = os.path.join(output_folder, 'output.csv')

# Ensure directories exist
for folder in [input_folder, output_folder, html_subfolder, photos_folder]:
    if not os.path.exists(folder):
        os.makedirs(folder)
        print(f"Created directory: {folder}")
    else:
        print(f"Directory already exists: {folder}")

# Copy Photos/ to docs/Photos/
if os.path.exists(photos_folder):
    if os.path.exists(docs_photos_folder):
        shutil.rmtree(docs_photos_folder)
    shutil.copytree(photos_folder, docs_photos_folder)
    print(f"Copied {photos_folder}/ to {docs_photos_folder}/")
else:
    os.makedirs(docs_photos_folder)
    print(f"Created empty {docs_photos_folder}/ (no photos found in {photos_folder}/)")

# Path to result.zip
zip_file = os.path.join(input_folder, 'result.zip')
temp_json_file = os.path.join(input_folder, 'result.json')

# Verify ZIP file existence and extract result.json
if not os.path.exists(zip_file):
    print(f"Error: 'result.zip' not found in '{input_folder}'. Exiting.")
    exit(1)

print(f"Extracting {zip_file}")
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        json_found = False
        for file_info in zip_ref.infolist():
            if file_info.filename.endswith('result.json'):
                zip_ref.extract(file_info, input_folder)
                extracted_path = os.path.join(input_folder, file_info.filename)
                if extracted_path != temp_json_file:
                    shutil.move(extracted_path, temp_json_file)
                json_found = True
                print(f"Extracted 'result.json' to {temp_json_file}")
                break
        if not json_found:
            print(f"Error: 'result.json' not found in '{zip_file}'. Exiting.")
            exit(1)
except zipfile.BadZipFile:
    print(f"Error: '{zip_file}' is not a valid ZIP file. Exiting.")
    exit(1)

# Verify extracted file existence
if not os.path.exists(temp_json_file):
    print(f"Error: Failed to extract 'result.json' from '{zip_file}'. Exiting.")
    exit(1)

# Load JSON data
print(f"Loading {temp_json_file}")
with open(temp_json_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

# Clean up the temporary JSON file
try:
    os.remove(temp_json_file)
    print(f"Cleaned up temporary file: {temp_json_file}")
except OSError as e:
    print(f"Warning: Could not remove {temp_json_file}: {e}")

# Access chats list
chats = data.get('chats', {}).get('list', [])
print(f"Found {len(chats)} chats in result.json")
if not chats:
    print("No chats found in 'result.json'. Please verify the file content.")
    exit(1)

# Define CSV columns
csv_columns = [
    'date', 'group name', 'total messages', 'Datedifference',
    'count of the hashtag "#FIVE"', 'count of the hashtag "#FOUR"',
    'count of the hashtag "#Three"', 'count of the hashtag "#SceneType"',
    'score', 'rank', 'total titles'
]

# Define history CSV columns
history_columns = ['date', 'group name', 'rank']

# Load existing history data
history_data = {}
if os.path.exists(history_csv_file):
    with open(history_csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            group = row.get('group name', 'Unknown')
            try:
                rank = int(float(row.get('rank', '0')))
                if group not in history_data:
                    history_data[group] = []
                history_data[group].append({'date': row.get('date', ''), 'rank': rank})
            except (ValueError, TypeError) as e:
                print(f"Skipping invalid rank for group '{group}': {row}. Error: {e}")
    print(f"Loaded {sum(len(v) for v in history_data.values())} history entries from {history_csv_file}")
else:
    print(f"No existing {history_csv_file} found")

# Initialize data storage
all_data = []
max_messages = 0
date_diffs = []
current_date = datetime.now().strftime('%Y-%m-%d')

# Function to sanitize filenames
def sanitize_filename(name):
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'\s+', '_', name)
    return name.lower()

# Function to find media file by serial number
def find_serial_match_media(serial_number, media_files):
    print(f"Searching for serial number '{serial_number}' in media files: {media_files}")
    for media in media_files:
        media_base = os.path.splitext(media)[0]
        if media_base == str(serial_number):
            print(f"Match found for serial number '{serial_number}': '{media}'")
            return media
    print(f"No match found for serial number '{serial_number}'")
    return None

# Process each chat
for chat in chats:
    if chat.get('type') == 'private_supergroup':
        group_name = chat.get('name', 'Unknown Group')
        group_id = str(chat['id'])
        telegram_group_id = group_id[4:] if group_id.startswith('-100') else group_id
        messages = chat.get('messages', [])
        print(f"Processing group: {group_name} (ID: {group_id})")

        total_messages = sum(1 for msg in messages if msg.get('type') == 'message')
        max_messages = max(max_messages, total_messages)

        # Hashtag counting
        hashtag_counts = {}
        for message in messages:
            if message.get('type') == 'message':
                text = message.get('text', '')
                if isinstance(text, list):
                    for entity in text:
                        if isinstance(entity, dict) and entity.get('type') == 'hashtag':
                            hashtag = entity.get('text')
                            if hashtag:
                                hashtag_upper = hashtag.upper()
                                special_ratings = ['#FIVE', '#FOUR', '#THREE']
                                special_scene_types = ['#FM', '#FF', '#FFM', '#FFFM', '#FFFFM', '#FMM', '#FMMM', '#FMMMM', '#FFMM', '#FFFMMM', '#ORGY']
                                if hashtag_upper in special_ratings + special_scene_types:
                                    hashtag = hashtag_upper
                                hashtag_counts[hashtag] = hashtag_counts.get(hashtag, 0) + 1

        # Calculate date_diff
        dates = []
        for message in messages:
            if message.get('type') == 'message':
                date_str = message.get('date')
                if date_str:
                    try:
                        date = datetime.fromisoformat(date_str)
                        dates.append(date)
                    except ValueError:
                        continue
        date_diff = None
        if dates:
            newest_date = max(dates)
            today = datetime.now()
            date_diff = (today - newest_date).days
            date_diffs.append(date_diff)
        print(f"Group {group_name}: Total messages = {total_messages}, Date diff = {date_diff}")

        # Hashtag lists
        special_ratings = ['#FIVE', '#FOUR', '#THREE']
        special_scene_types = ['#FM', '#FF', '#FFM', '#FFFM', '#FFFFM', '#FMM', '#FMMM', '#FMMMM', '#FFMM', '#FFFMMM', '#ORGY']
        ratings_hashtag_list = ''.join(f'<li class="hashtag-item">{h}: {hashtag_counts[h]}</li>\n' for h in sorted(hashtag_counts) if h in special_ratings) or '<li>No rating hashtags (#FIVE, #FOUR, #Three) found</li>'
        scene_types_hashtag_list = ''.join(f'<li class="hashtag-item">{h}: {hashtag_counts[h]}</li>\n' for h in sorted(hashtag_counts) if h in special_scene_types) or '<li>No scene type hashtags found</li>'
        other_hashtag_list = ''.join(f'<li class="hashtag-item">{h}: {hashtag_counts[h]}</li>\n' for h in sorted(hashtag_counts) if h not in special_ratings and h not in special_scene_types) or '<li>No other hashtags found</li>'

        scene_type_count = sum(hashtag_counts.get(h, 0) for h in special_scene_types)
        date_diff_text = f'{date_diff} days' if date_diff is not None else 'N/A'

        # Titles with serial numbers
        titles = []
        media_extensions = ['.mp4', '.webm', '.ogg', '.gif']
        group_subfolder = os.path.join(docs_photos_folder, group_name)
        thumbs_subfolder = os.path.join(group_subfolder, 'thumbs')
        media_files = [f for f in os.listdir(thumbs_subfolder) if f.lower().endswith(tuple(media_extensions))] if os.path.exists(thumbs_subfolder) else []
        fallback_photos = [f for f in os.listdir(group_subfolder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')) and os.path.isfile(os.path.join(group_subfolder, f))] if os.path.exists(group_subfolder) else []
        print(f"Group {group_name}: Thumbs media files = {media_files}, Fallback photos = {fallback_photos}")
        serial_number = 1
        for message in messages:
            if message.get('action') == 'topic_created':
                title = message.get('title', '')
                message_id = message.get('id')
                date_str = message.get('date', '')
                if title.strip() and message_id and date_str:
                    try:
                        date = datetime.fromisoformat(date_str).strftime('%Y-%m-%d')
                        media_path = 'https://via.placeholder.com/600x300'
                        is_gif = False
                        if media_files:
                            serial_match = find_serial_match_media(serial_number, media_files)
                            if serial_match:
                                media_path = f"../Photos/{group_name}/thumbs/{serial_match}"
                                is_gif = serial_match.lower().endswith('.gif')
                                print(f"Group {group_name}, Title '{title}' (S.No {serial_number}): Matched media '{serial_match}', selected path {media_path}")
                        else:
                            print(f"Group {group_name}, Title '{title}' (S.No {serial_number}): No media files in {thumbs_subfolder}")
                            if fallback_photos:
                                random_photo = random.choice(fallback_photos)
                                media_path = f"../Photos/{group_name}/{random_photo}"
                                is_gif = random_photo.lower().endswith('.gif')
                                print(f"  Using fallback photo: {media_path}")
                        titles.append({
                            'title': title,
                            'message_id': message_id,
                            'date': date,
                            'media_path': media_path,
                            'is_gif': is_gif,
                            'serial_number': serial_number
                        })
                        serial_number += 1
                    except ValueError:
                        continue
        titles.sort(key=lambda x: x['date'], reverse=True)  # Sort by date, newest first
        titles_count = len(titles)

        # Titles grid
        titles_grid = f"<p>Total Titles: {titles_count}</p><div class='titles-grid' id='titlesGrid'>"
        for t in titles:
            media_element = (
                f"<img src='{t['media_path']}' alt='Media for {t['title']}' style='width:100%;max-width:600px;height:300px;object-fit:cover;'>"
                if t['is_gif'] or t['media_path'] == 'https://via.placeholder.com/600x300'
                else f"<video src='{t['media_path']}' style='width:100%;max-width:600px;height:300px;object-fit:cover;' loop muted playsinline></video>"
            )
            titles_grid += f"""
                <div class='grid-item'>
                    {media_element}
                    <p class='title'><a href='https://t.me/c/{telegram_group_id}/{t['message_id']}' target='_blank'>{t['title']}</a></p>
                    <p class='date'>S.No: {t['serial_number']} | {t['date']}</p>
                </div>
            """
        titles_grid += f"</div>" if titles else f"<p>No titles found (Total: {titles_count})</p>"

        # Titles table
        titles_table = f"<table class='titles-table' id='titlesTable'><thead><tr><th onclick='sortTitlesTable(0)'>S.No</th><th onclick='sortTitlesTable(1)'>Items</th><th onclick='sortTitlesTable(2)'>Date</th></tr></thead><tbody id='titlesTableBody'>"
        for t in titles:
            titles_table += f"<tr><td>{t['serial_number']}</td><td><a href='https://t.me/c/{telegram_group_id}/{t['message_id']}' target='_blank'>{t['title']}</a></td><td>{t['date']}</td></tr>"
        titles_table += f"</tbody></table>" if titles else f"<p>No titles found</p>"

        # Photos for slideshow
        photo_paths = []
        if os.path.exists(group_subfolder):
            photo_paths = [f"../Photos/{group_name}/{f}" for f in os.listdir(group_subfolder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')) and os.path.isfile(os.path.join(group_subfolder, f))]
            print(f"Group {group_name}: Found {len(photo_paths)} photos in {group_subfolder}: {photo_paths}")
        if not photo_paths:
            photo_paths = ['https://via.placeholder.com/1920x800']
            print(f"Group {group_name}: Using placeholder for slideshow")

        slideshow_content = '<div class="container">\n' + ''.join(f'<div class="mySlides"><div class="numbertext">{i} / {len(photo_paths)}</div><img src="{p}" style="width:100%;height:auto;"></div>' for i, p in enumerate(photo_paths, 1)) + """
            <a class="prev" onclick="plusSlides(-1)">❮</a>
            <a class="next" onclick="plusSlides(1)">❯</a>
            <div class="caption-container"><p id="caption"></p></div>
            <div class="row">
        """ + ''.join(f'<div class="column"><img class="demo cursor" src="{p}" style="width:100%" onclick="currentSlide({i})" alt="{group_name} Photo {i}"></div>' for i, p in enumerate(photo_paths, 1)) + '</div></div>'

        photo_file_name = next((f"{group_name}{ext}" for ext in ('.jpg', '.jpeg', '.png', '.gif', '.webp') if os.path.exists(os.path.join(docs_photos_folder, f"{group_name}{ext}"))), None)
        if photo_file_name:
            print(f"Group {group_name}: Found single photo at {docs_photos_folder}/{photo_file_name}")
        else:
            print(f"Group {group_name}: No single photo found in {docs_photos_folder}/")

        if group_name not in history_data:
            history_data[group_name] = []

        # Pre-compute JSON for history data to avoid f-string issue
        history_data_json = json.dumps(history_data.get(group_name, []))

        # HTML content for group pages
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{group_name}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background-color: #1e2a44; color: #ffffff; text-align: center; }}
        h1, h2 {{ color: #e6b800; width: 80%; margin: 20px auto; text-align: center; font-size: 36px; }}
        .info {{ background-color: #2a3a5c; padding: 10px; border-radius: 5px; margin-bottom: 20px; }}
        .hashtags {{ list-style-type: none; padding: 0; }}
        .hashtag-item {{ background-color: #3b4a6b; margin: 5px 0; padding: 5px; border-radius: 3px; display: inline-block; width: 200px; color: #ffffff; }}
        .rank-container {{ 
            width: 80%; 
            margin: 20px auto; 
            display: flex; 
            justify-content: center; 
            align-items: center; 
            gap: 20px; 
            flex-wrap: wrap; 
        }}
        .rank-number {{ font-size: 48px; font-weight: bold; color: #e6b800; display: inline-block; }}
        @keyframes countUp {{ from {{ content: "0"; }} to {{ content: attr(data-rank); }} }}
        .rank-number::before {{ content: "0"; animation: countUp 2s ease-out forwards; display: inline-block; min-width: 60px; }}
        .chart-container {{ max-width: 400px; width: 100%; background-color: #2a3a5c; padding: 10px; border-radius: 5px; }}
        canvas {{ width: 100% !important; height: auto !important; }}
        .titles-grid {{ 
            display: grid; 
            grid-template-columns: repeat(3, 1fr); 
            gap: 20px; 
            margin: 20px auto; 
            max-width: 1800px; 
            width: 80%; 
            box-sizing: border-box; 
        }}
        .grid-item {{ 
            background-color: #2a3a5c; 
            padding: 10px; 
            border-radius: 5px; 
            text-align: center; 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            width: 100%; 
            box-sizing: border-box; 
        }}
        .grid-item video, .grid-item img {{ 
            width: 100%; 
            max-width: 600px; 
            height: 300px; 
            object-fit: cover; 
            border-radius: 5px; 
        }}
        .grid-item .title {{ 
            margin: 10px 0 5px; 
            font-size: 16px; 
            font-weight: bold; 
            color: #e6b800; 
        }}
        .grid-item .date {{ 
            margin: 0; 
            font-size: 14px; 
            color: #cccccc; 
        }}
        .titles-table {{ 
            width: 80%; 
            margin: 20px auto; 
            border-collapse: collapse; 
            background-color: #2a3a5c; 
        }}
        .titles-table th, .titles-table td {{ 
            padding: 10px; 
            border: 1px solid #3b4a6b; 
            text-align: left; 
            vertical-align: middle; 
            color: #ffffff; 
        }}
        .titles-table th {{ 
            background-color: #e6b800; 
            color: #1e2a44; 
            cursor: pointer; 
        }}
        .titles-table th:hover {{ 
            background-color: #b30000; 
        }}
        a {{ color: #e6b800; text-decoration: none; }}
        a:hover {{ color: #b30000; text-decoration: underline; }}
        .container {{ 
            position: relative; 
            width: 80%; 
            margin: 20px auto; 
            height: auto; 
            max-height: 600px; 
            display: block; 
            overflow: hidden; 
            background-color: #2a3a5c; 
        }}
        .mySlides {{ 
            display: none; 
            width: 100%; 
            height: auto; 
            aspect-ratio: 16/9; 
        }}
        .mySlides img {{ 
            width: 100%; 
            height: auto; 
            object-fit: contain; 
        }}
        .cursor {{ cursor: pointer; }}
        .prev, .next {{ 
            cursor: pointer; 
            position: absolute; 
            top: 50%; 
            transform: translateY(-50%); 
            width: auto; 
            padding: 16px; 
            color: #e6b800; 
            font-weight: bold; 
            font-size: 20px; 
            border-radius: 0 3px 3px 0; 
            user-select: none; 
            -webkit-user-select: none; 
            z-index: 10; 
        }}
        .prev {{ left: 0; }}
        .next {{ right: 0; border-radius: 3px 0 0 3px; }}
        .prev:hover, .next:hover {{ background-color: #b30000; }}
        .numbertext {{ 
            color: #e6b800; 
            font-size: 12px; 
            padding: 8px 12px; 
            position: absolute; 
            top: 0; 
            z-index: 10; 
        }}
        .caption-container {{ 
            text-align: center; 
            background-color: #1e2a44; 
            padding: 2px 16px; 
            color: #e6b800; 
        }}
        .row {{ 
            display: flex; 
            flex-wrap: wrap; 
            justify-content: center; 
            margin-top: 10px; 
        }}
        .column {{ 
            flex: 0 0 {100 / len(photo_paths) if photo_paths else 100}%; 
            max-width: 100px; 
            padding: 5px; 
        }}
        .demo {{ 
            opacity: 0.6; 
            width: 100%; 
            height: auto; 
            object-fit: cover; 
        }}
        .active, .demo:hover {{ opacity: 1; }}
        .tab {{ 
            overflow: hidden; 
            margin: 20px auto; 
            width: 80%; 
            background-color: #2a3a5c; 
            border-radius: 5px 5px 0 0; 
        }}
        .tab button {{ 
            background-color: #2a3a5c; 
            color: #e6b800; 
            float: left; 
            border: none; 
            outline: none; 
            cursor: pointer; 
            padding: 14px 16px; 
            transition: 0.3s; 
            font-size: 17px; 
            width: 50%; 
        }}
        .tab button:hover {{ background-color: #b30000; }}
        .tab button.active {{ background-color: #3b4a6b; }}
        .tabcontent {{ 
            display: none; 
            padding: 6px 12px; 
            border-top: none; 
            background-color: #2a3a5c; 
            margin: 0 auto; 
            width: 80%; 
            border-radius: 0 0 5px 5px; 
        }}
        #Videos {{ display: block; }}
        @media only screen and (max-width: 1800px) {{ 
            .titles-grid {{ grid-template-columns: repeat(2, 1fr); }} 
        }}
        @media only screen and (max-width: 1200px) {{ 
            .titles-grid {{ grid-template-columns: 1fr; }} 
        }}
        @media only screen and (max-width: 768px) {{ 
            .container {{ width: 80%; max-height: 400px; }} 
            h1 {{ width: 80%; margin: 10px auto; font-size: 30px; }}
            .rank-container {{ width: 80%; flex-direction: column; gap: 10px; }} 
            .chart-container {{ max-width: 100%; }} 
            .column {{ flex: 0 0 80px; max-width: 80px; }} 
            .mySlides img {{ object-fit: contain; }} 
            .tab button {{ font-size: 14px; padding: 10px; }}
        }}
    </style>
</head>
<body>
    <h1>{group_name}</h1>
    <div class="rank-container">
        <div class="chart-container"><h2>Rank History</h2><canvas id="rankChart"></canvas></div>
        <p>Rank: <span class="rank-number" data-rank="RANK_PLACEHOLDER"></span></p>
    </div>
    {slideshow_content}
    <div class="info"><p>Scenes: {total_messages}</p><p>Last Scene: {date_diff_text}</p></div>
    <div class="info">
        <h2>Rating Hashtag Counts (#FIVE, #FOUR, #Three)</h2><ul class="hashtags">{ratings_hashtag_list}</ul>
        <h2>Scene Type Hashtag Counts</h2><ul class="hashtags">{scene_types_hashtag_list}</ul>
        <h2>Other Hashtag Counts</h2><ul class="hashtags">{other_hashtag_list}</ul>
    </div>
    <div class="info">
        <h2>Titles</h2>
        <div class="tab">
            <button class="tablinks active" onclick="openTab(event, 'Videos')">Videos</button>
            <button class="tablinks" onclick="openTab(event, 'Table')">Table</button>
        </div>
        <div id="Videos" class="tabcontent">
            {titles_grid}
        </div>
        <div id="Table" class="tabcontent">
            {titles_table}
        </div>
    </div>
    <script>
        let slideIndex = 1;
        showSlides(slideIndex);
        function plusSlides(n) {{ 
            clearInterval(autoSlide); 
            showSlides(slideIndex += n); 
            autoSlide = setInterval(() => plusSlides(1), 3000); 
        }}
        function currentSlide(n) {{ 
            clearInterval(autoSlide); 
            showSlides(slideIndex = n); 
            autoSlide = setInterval(() => plusSlides(1), 3000); 
        }}
        function showSlides(n) {{
            let i;
            let slides = document.getElementsByClassName("mySlides");
            let dots = document.getElementsByClassName("demo");
            let captionText = document.getElementById("caption");
            if (n > slides.length) {{ slideIndex = 1 }}
            if (n < 1) {{ slideIndex = slides.length }}
            for (i = 0; i < slides.length; i++) {{ 
                slides[i].style.display = "none"; 
            }}
            for (i = 0; i < dots.length; i++) {{ 
                dots[i].className = dots[i].className.replace(" active", ""); 
            }}
            slides[slideIndex-1].style.display = "block";
            dots[slideIndex-1].className += " active";
            captionText.innerHTML = dots[slideIndex-1].alt;
        }}
        let autoSlide = setInterval(() => plusSlides(1), 3000);

        function openTab(evt, tabName) {{
            let i, tabcontent, tablinks;
            tabcontent = document.getElementsByClassName("tabcontent");
            for (i = 0; i < tabcontent.length; i++) {{
                tabcontent[i].style.display = "none";
            }}
            tablinks = document.getElementsByClassName("tablinks");
            for (i = 0; i < tablinks.length; i++) {{
                tablinks[i].className = tablinks[i].className.replace(" active", "");
            }}
            document.getElementById(tabName).style.display = "block";
            evt.currentTarget.className += " active";
        }}

        // Chart.js for rank history
        document.addEventListener('DOMContentLoaded', function() {{
            const ctx = document.getElementById('rankChart').getContext('2d');
            const historyData = {history_data_json};
            const dates = historyData.map(entry => entry.date);
            const ranks = historyData.map(entry => entry.rank);
            new Chart(ctx, {{
                type: 'line',
                data: {{ 
                    labels: dates, 
                    datasets: [{{
                        label: 'Rank Over Time', 
                        data: ranks, 
                        borderColor: '#e6b800', 
                        backgroundColor: 'rgba(230, 184, 0, 0.2)', 
                        fill: true, 
                        tension: 0.4 
                    }}] 
                }},
                options: {{ 
                    scales: {{ 
                        y: {{ 
                            beginAtZero: true, 
                            title: {{ display: true, text: 'Rank', color: '#e6b800' }}, 
                            ticks: {{ stepSize: 1, color: '#ffffff' }}, 
                            suggestedMax: {len(chats) + 1},
                            grid: {{ color: '#3b4a6b' }}
                        }}, 
                        x: {{ 
                            title: {{ display: true, text: 'Date', color: '#e6b800' }},
                            ticks: {{ color: '#ffffff' }},
                            grid: {{ color: '#3b4a6b' }}
                        }} 
                    }}, 
                    plugins: {{ 
                        legend: {{ display: true, labels: {{ color: '#e6b800' }} }} 
                    }} 
                }}
            }});

            // Add hover-to-play for videos in titles grid
            const videos = document.querySelectorAll('.grid-item video');
            videos.forEach(video => {{
                video.addEventListener('mouseover', () => {{
                    video.play().catch(error => {{
                        console.error('Error playing video:', error);
                    }});
                }});
                video.addEventListener('mouseout', () => {{
                    video.pause();
                }});
            }});

            // Initialize titles table sorted by S.No descending (highest ID at top)
            sortTitlesTable(0, -1); // Sort by S.No column, highest first
        }});

        // Titles table and grid sorting
        let titlesSortDirections = [-1, 0, 0]; // S.No starts descending
        function sortTitlesTable(columnIndex, forceDirection) {{
            const tbody = document.getElementById('titlesTableBody');
            const rows = Array.from(tbody.getElementsByTagName('tr'));
            const direction = forceDirection !== undefined ? forceDirection : (titlesSortDirections[columnIndex] === 1 ? -1 : 1);
            rows.sort((a, b) => {{
                let aValue = a.cells[columnIndex].innerText;
                let bValue = b.cells[columnIndex].innerText;
                if (columnIndex === 0) {{ // S.No column
                    aValue = parseInt(aValue);
                    bValue = parseInt(bValue);
                    return direction * (aValue - bValue);
                }} else if (columnIndex === 2) {{ // Date column
                    aValue = new Date(aValue);
                    bValue = new Date(bValue);
                    return direction * (aValue - bValue);
                }} else if (columnIndex === 1) {{ // Items column
                    return direction * aValue.localeCompare(bValue);
                }}
                return 0;
            }});
            while (tbody.firstChild) {{ 
                tbody.removeChild(tbody.firstChild); 
            }}
            rows.forEach(row => tbody.appendChild(row));
            titlesSortDirections[columnIndex] = direction;
            titlesSortDirections = titlesSortDirections.map((d, i) => i === columnIndex ? d : 0);
            // Sync grid with table
            sortTitlesGrid(columnIndex, direction);
        }}

        function sortTitlesGrid(columnIndex, direction) {{
            const grid = document.getElementById('titlesGrid');
            const items = Array.from(grid.getElementsByClassName('grid-item'));
            items.sort((a, b) => {{
                let aValue, bValue;
                if (columnIndex === 0) {{ // S.No
                    aValue = parseInt(a.querySelector('.date').innerText.split('S.No: ')[1].split(' | ')[0]);
                    bValue = parseInt(b.querySelector('.date').innerText.split('S.No: ')[1].split(' | ')[0]);
                    return direction * (aValue - bValue);
                }} else if (columnIndex === 1) {{ // Items
                    aValue = a.querySelector('.title').innerText;
                    bValue = b.querySelector('.title').innerText;
                    return direction * aValue.localeCompare(bValue);
                }} else if (columnIndex === 2) {{ // Date
                    aValue = new Date(a.querySelector('.date').innerText.split(' | ')[1]);
                    bValue = new Date(b.querySelector('.date').innerText.split(' | ')[1]);
                    return direction * (aValue - bValue);
                }}
                return 0;
            }});
            while (grid.firstChild) {{ 
                grid.removeChild(grid.firstChild); 
            }}
            items.forEach(item => grid.appendChild(item));
        }}
    </script>
</body>
</html>
"""

        sanitized_name = sanitize_filename(group_name)
        html_file = f"{sanitized_name}_{group_id}.html"
        html_filename = os.path.join(html_subfolder, html_file)

        all_data.append({
            'date': current_date,
            'group name': group_name,
            'total messages': total_messages,
            'Datedifference': date_diff if date_diff is not None else 'N/A',
            'count of the hashtag "#FIVE"': hashtag_counts.get('#FIVE', 0),
            'count of the hashtag "#FOUR"': hashtag_counts.get('#FOUR', 0),
            'count of the hashtag "#Three"': hashtag_counts.get('#THREE', 0),
            'count of the hashtag "#SceneType"': scene_type_count,
            'score': 0,
            'rank': 0,
            'total titles': titles_count,
            'html_file': html_file,
            'html_content': html_content,
            'photo_file_name': f"Photos/{photo_file_name}" if photo_file_name else None
        })

# Calculate scores
min_date_diff = min(date_diffs) if date_diffs else 0
max_date_diff_denom = max(date_diffs) - min_date_diff if date_diffs and max(date_diffs) > min_date_diff else 1

for entry in all_data:
    five_count = entry['count of the hashtag "#FIVE"']
    four_count = entry['count of the hashtag "#FOUR"']
    three_count = entry['count of the hashtag "#Three"']
    messages = entry['total messages']
    diff = entry['Datedifference']

    hashtag_score = (10 * five_count) + (5 * four_count) + (1 * three_count)
    messages_score = (messages / max_messages) * 10 if max_messages > 0 else 0
    date_score = 0
    if diff != 'N/A' and date_diffs:
        date_score = 10 * (1 - (diff - min_date_diff) / max_date_diff_denom) if max_date_diff_denom > 0 else 10
    entry['score'] = hashtag_score + messages_score + date_score

# Sort by score and assign ranks
sorted_data = sorted(all_data, key=lambda x: x['score'], reverse=True)
for i, entry in enumerate(sorted_data, 1):
    entry['rank'] = i
    history_data[entry['group name']].append({'date': current_date, 'rank': i})
    html_content_with_rank = entry['html_content'].replace('RANK_PLACEHOLDER', str(i))
    html_path = os.path.join(html_subfolder, entry['html_file'])
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content_with_rank)
    print(f"Wrote HTML file: {html_path}")

# Write current run to output.csv
csv_data = [{k: v for k, v in entry.items() if k in csv_columns} for entry in sorted_data]
with open(csv_file, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=csv_columns)
    writer.writeheader()
    writer.writerows(csv_data)
print(f"\nWrote CSV file: {csv_file}")

# Append new history entries to history.csv
new_history_rows = [{'date': current_date, 'group name': entry['group name'], 'rank': entry['rank']} for entry in sorted_data]
new_history_rows = [row for row in new_history_rows if row.get('group name') and row.get('rank') is not None]
if new_history_rows:
    write_header = not os.path.exists(history_csv_file)
    with open(history_csv_file, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=history_columns)
        if write_header:
            writer.writeheader()
        writer.writerows(new_history_rows)
    print(f"\nAppended {len(new_history_rows)} rows to {history_csv_file}")
else:
    print(f"No new history entries to append to {history_csv_file}")

# Generate ranking HTML
total_groups = len(sorted_data)
table_rows = ''
for entry in sorted_data:
    group_name = escape(entry['group name'])
    photo_src = entry['photo_file_name'] if entry['photo_file_name'] else 'https://via.placeholder.com/300'
    html_link = f"HTML/{entry['html_file']}"
    last_scene = f"{entry['Datedifference']} days" if entry['Datedifference'] != 'N/A' else 'N/A'
    table_rows += f"""
    <tr>
        <td>{entry['rank']}</td>
        <td><a href="{html_link}" target="_blank">{group_name}</a></td>
        <td><div class="flip-card"><div class="flip-card-inner"><div class="flip-card-front"><img src="{photo_src}" alt="{group_name}" style="width:300px;height:300px;object-fit:cover;"></div><div class="flip-card-back"><a href="{html_link}" target="_blank" style="color: #e6b800; text-decoration: none;"><h1>{group_name}</h1></a></div></div></div></td>
        <td>{last_scene}</td>
        <td>{entry['total titles']}</td>
        <td>{entry['count of the hashtag "#FIVE"']}</td>
        <td>{entry['count of the hashtag "#FOUR"']}</td>
        <td>{entry['count of the hashtag "#Three"']}</td>
        <td>{entry['count of the hashtag "#SceneType"']}</td>
        <td>{entry['score']:.2f}</td>
    </tr>
    """

ranking_html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PS Ranking - {current_date}</title>
    <style>
        body {{ font-family: Arial, sans-serif; background-color: #1e2a44; color: #ffffff; margin: 20px; text-align: center; }}
        h1, h2 {{ color: #e6b800; }}
        table {{ width: 80%; margin: 20px auto; border-collapse: collapse; background-color: #2a3a5c; box-shadow: 0 0 10px rgba(0, 0, 0, 0.3); }}
        th, td {{ border: 1px solid #3b4a6b; text-align: center; vertical-align: middle; padding: 15px; color: #ffffff; }}
        th {{ background-color: #e6b800; color: #1e2a44; cursor: pointer; }}
        th:hover {{ background-color: #b30000; }}
        tr:hover {{ background-color: #3b4a6b; }}
        a {{ text-decoration: none; color: #e6b800; }}
        a:hover {{ color: #b30000; text-decoration: underline; }}
        .flip-card {{ background-color: transparent; width: 300px; height: 300px; perspective: 1000px; margin: auto; }}
        .flip-card-inner {{ position: relative; width: 100%; height: 100%; text-align: center; transition: transform 0.6s; transform-style: preserve-3d; box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2); }}
        .flip-card:hover .flip-card-inner {{ transform: rotateY(180deg); }}
        .flip-card-front, .flip-card-back {{ position: absolute; width: 100%; height: 100%; backface-visibility: hidden; border-radius: 5px; }}
        .flip-card-front {{ background-color: #2a3a5c; color: #ffffff; }}
        .flip-card-back {{ background-color: #3b4a6b; color: #e6b800; transform: rotateY(180deg); display: flex; justify-content: center; align-items: center; flex-direction: column; }}
        .flip-card-back h1 {{ margin: 0; font-size: 24px; word-wrap: break-word; padding: 10px; }}
        @media only screen and (max-width: 1200px) {{ 
            table {{ width: 90%; }} 
            .flip-card {{ width: 200px; height: 200px; }} 
            .flip-card-back h1 {{ font-size: 18px; }}
            th, td {{ font-size: 14px; padding: 10px; }}
        }}
        @media only screen and (max-width: 768px) {{ 
            table {{ width: 95%; }} 
            .flip-card {{ width: 150px; height: 150px; }} 
            .flip-card-back h1 {{ font-size: 16px; }}
            th, td {{ font-size: 12px; padding: 8px; }}
        }}
    </style>
</head>
<body>
    <h1>PS Ranking - {current_date}</h1>
    <h2>Total Number of Groups: {total_groups}</h2>
    <table id="rankingTable">
        <thead>
            <tr>
                <th onclick="sortTable(0)">Rank</th>
                <th onclick="sortTable(1)">Group Name</th>
                <th>Photo</th>
                <th onclick="sortTable(3)">Last Scene</th>
                <th onclick="sortTable(4)">Total Titles</th>
                <th onclick="sortTable(5)">#FIVE</th>
                <th onclick="sortTable(6)">#FOUR</th>
                <th onclick="sortTable(7)">#Three</th>
                <th onclick="sortTable(8)">Thumbnails</th>
                <th onclick="sortTable(9)">Score</th>
            </tr>
        </thead>
        <tbody id="tableBody">
            {table_rows}
        </tbody>
    </table>
    <script>
        let sortDirections = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        function sortTable(columnIndex) {{
            if (columnIndex === 2) return; // Skip Photo column
            const tbody = document.getElementById('tableBody');
            const rows = Array.from(tbody.getElementsByTagName('tr'));
            const isNumeric = [true, false, false, true, true, true, true, true, true, true];
            const direction = sortDirections[columnIndex] === 1 ? -1 : 1;

            rows.sort((a, b) => {{
                let aValue = a.cells[columnIndex].innerText;
                let bValue = b.cells[columnIndex].innerText;

                if (columnIndex === 3) {{ // Last Scene column
                    if (aValue === 'N/A' && bValue === 'N/A') return 0;
                    if (aValue === 'N/A') return direction * 1;
                    if (bValue === 'N/A') return direction * -1;
                    aValue = parseInt(aValue);
                    bValue = parseInt(bValue);
                    return direction * (aValue - bValue);
                }}

                if (isNumeric[columnIndex]) {{ 
                    aValue = parseFloat(aValue) || aValue; 
                    bValue = parseFloat(bValue) || bValue; 
                    return direction * (aValue - bValue); 
                }}
                return direction * aValue.localeCompare(bValue);
            }});

            while (tbody.firstChild) {{ 
                tbody.removeChild(tbody.firstChild); 
            }}
            rows.forEach(row => tbody.appendChild(row));
            sortDirections[columnIndex] = direction;
            sortDirections = sortDirections.map((d, i) => i === columnIndex ? d : 0);
        }}
    </script>
</body>
</html>
"""

# Write ranking HTML file
ranking_html_file = os.path.join(output_folder, 'index.html')
with open(ranking_html_file, 'w', encoding='utf-8') as f:
    f.write(ranking_html_content)
print(f"\nWrote ranking HTML file: {ranking_html_file}")

print(f"\nProcessed {len(chats)} groups. Output written to {output_folder}")