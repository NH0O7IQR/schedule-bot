import os
import re
import sys
import json
import requests
import vk_api
import pandas as pd
from datetime import datetime, timedelta, timezone
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo
import hashlib

# НАСТРОЙКИ
VK_TOKEN = os.environ.get('VK_TOKEN')
GROUP_ID = int(os.environ.get('GROUP_ID', 29490619))
TOPIC_ID = int(os.environ.get('TOPIC_ID', 68268274))
SHEET_NAME = os.environ.get('SHEET_NAME', 'ПЕРК-243')
TRAVEL_MINUTES = 90 
TIMEZONE = ZoneInfo("Europe/Moscow")

MIN_STUDY_EVENTS = int(os.environ.get('MIN_STUDY_EVENTS', 10))
FORCE_UPDATE = os.environ.get('FORCE_UPDATE') == '1'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Регулярки
TITLE_RE = re.compile(r'перк[\s\-_]*(3|243)(?!\d)', re.I)
TIME_RE = re.compile(r'(\d{1,2})[.:](\d{2})\s*[-–—]\s*(\d{1,2})[.:](\d{2})')
ROOM_RE = re.compile(r'^(?P<teacher>.+?)\s*[-–—]\s*(?P<room>[\dА-Яа-яA-Za-z]*-?\d+[А-Яа-яA-Za-z]?)$')

def split_teacher_room(s: str) -> tuple[str, str]:
    info = re.split(r'\s+[-–—]\s+', s.strip())
    if len(info) >= 2:
        return info[0].strip(), info[-1].strip()
    m = ROOM_RE.match(s.strip())
    if m:
        return m['teacher'].strip(), m['room'].strip()
    return s.strip(), ""

def download_excel():
    print("Ищем файл с расписанием в ВК...")
    if not VK_TOKEN:
        sys.exit("Ошибка: VK_TOKEN не задан в переменных окружения!")
        
    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()
    
    try:
        comments = vk.board.getComments(group_id=GROUP_ID, topic_id=TOPIC_ID, count=100, sort='desc')
    except (vk_api.exceptions.ApiError, requests.exceptions.RequestException) as e:
        sys.exit(f"Ошибка VK API: {e}")

    candidates = []
    for item in comments.get('items', []):
        for attach in item.get('attachments', []):
            if attach.get('type') == 'doc':
                doc = attach['doc']
                if doc['ext'] in ('xls', 'xlsx') and TITLE_RE.search(doc['title']):
                    candidates.append(doc)

    if not candidates:
        sys.exit("Excel файл не найден в обсуждении.")

    doc = max(candidates, key=lambda d: d['date'])
    print(f"Найден файл: {doc['title']} (от {datetime.fromtimestamp(doc['date'], tz=TIMEZONE).strftime('%d.%m.%Y %H:%M')})")
    
    try:
        response = requests.get(doc['url'], timeout=15)
        response.raise_for_status()
        ext = doc['ext']
        filename = os.path.join(BASE_DIR, f'schedule.{ext}')
        with open(filename, 'wb') as f:
            f.write(response.content)
        
        return filename, datetime.fromtimestamp(doc['date'], tz=timezone.utc)
    except requests.RequestException as e:
        sys.exit(f"Ошибка скачивания файла: {e}")

def parse_date(v):
    if pd.isna(v): return None
    if isinstance(v, (datetime, pd.Timestamp)): return v.date()
    s = str(v).strip()
    
    m = re.search(r'(\d{4}-\d{2}-\d{2})', s)
    if m and len(s) - len(m.group(1)) <= 5:
        return datetime.strptime(m.group(1), "%Y-%m-%d").date()
            
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})', s)
    if m and len(s) - len(m.group(1)) <= 5:
        return datetime.strptime(m.group(1), "%d.%m.%Y").date()
    return None

def parse_excel(filename):
    print("Парсим расписание...")
    df = pd.read_excel(filename, sheet_name=SHEET_NAME, header=None)
    events = []
    current_dates = []
    skipped_dates = 0
    total_dates_checked = 0
    skipped_times = 0
    
    for _, row in df.iterrows():
        parsed_dates = [parse_date(row[c]) for c in range(3, 21)]
        if sum(d is not None for d in parsed_dates) >= 5:
            current_dates = parsed_dates
            continue
            
        time_cell = str(row[2]).strip()
        time_match = TIME_RE.search(time_cell)
        
        if time_match and current_dates:
            if int(time_match.group(1)) == 0 and int(time_match.group(2)) == 0:
                continue
                
            try:
                start_time = datetime.strptime(f"{time_match.group(1)}:{time_match.group(2)}", "%H:%M").time()
                end_time = datetime.strptime(f"{time_match.group(3)}:{time_match.group(4)}", "%H:%M").time()
                if end_time <= start_time:
                    skipped_times += 1
                    continue
            except ValueError:
                skipped_times += 1
                continue 

            for i, col_idx in enumerate(range(3, 21)):
                cell_value = str(row[col_idx]).strip()
                
                if cell_value and cell_value.lower() not in ['nan', '-', '—', ''] and 'самостоятельн' not in cell_value.lower():
                    total_dates_checked += 1
                    
                    event_date = current_dates[i]
                    if event_date is None:
                        skipped_dates += 1
                        continue
                        
                    parts = cell_value.replace('<br>', '\n').split('\n')
                    subject = parts[0].strip()
                    if not subject: continue
                        
                    teacher, room = "", ""
                    if len(parts) > 1:
                        teacher, room = split_teacher_room(parts[1])
                    
                    extra_info = " ".join(parts[2:]).strip()
                    desc = f"Преподаватель: {teacher}" if teacher else ""
                    if extra_info:
                        desc = f"{desc}\n{extra_info}".strip()

                    start_dt = datetime.combine(event_date, start_time, tzinfo=TIMEZONE)
                    end_dt = datetime.combine(event_date, end_time, tzinfo=TIMEZONE)
                    
                    events.append({
                        'summary': subject,
                        'location': room,
                        'description': desc,
                        'start': start_dt,
                        'end': end_dt
                    })
                    
    if skipped_times:
        print(f"⚠️ Строк с ошибкой времени пропущено: {skipped_times}")
        
    if events:
        min_date = min(e['start'] for e in events).strftime('%d.%m.%Y')
        max_date = max(e['start'] for e in events).strftime('%d.%m.%Y')
        print(f"Диапазон дат в календаре: {min_date} — {max_date}")
        
    return events, skipped_dates, total_dates_checked

def check_changes(new_events, report=True):
    print("Проверяем изменения...")
    json_path = os.path.join(BASE_DIR, 'last_schedule.json')
    old_events = []
    
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                old_events = json.load(f)
        except json.JSONDecodeError:
            print("⚠️ Файл last_schedule.json поврежден. Создаем заново.")
            old_events = []
            
    serialized_new = []
    for ev in new_events:
        serialized_new.append({
            'date': ev['start'].strftime('%Y-%m-%d'),
            'time': ev['start'].strftime('%H:%M'),
            'end_time': ev['end'].strftime('%H:%M'),
            'subject': ev['summary'],
            'room': ev['location'],
            'teacher': ev['description'].replace('Преподаватель: ', '').strip()
        })
    serialized_new.sort(key=lambda e: (e['date'], e['time'], e['subject']))

    changes = []
    
    if old_events and report:
        # Фильтруем прошедшие пары (строго больше сегодняшнего дня)
        today_str = datetime.now(TIMEZONE).strftime('%Y-%m-%d')
        old_events_filtered = [e for e in old_events if e['date'] > today_str]
        
        old_dates = set(e['date'] for e in old_events_filtered)
        new_dates = set(e['date'] for e in serialized_new)
        
        if len(old_dates) > 0 and len(new_dates) > 0:
            intersection = old_dates.intersection(new_dates)
            min_len = min(len(old_dates), len(new_dates))
            # Смена семестра, если пересечений меньше 20%
            if (len(intersection) / min_len) < 0.2:
                print("⚠️ Обнаружена смена семестра (даты почти не пересекаются). Отчет не создаем.")
                changes = []
            else:
                old_dict = {(e['date'], e['time'], e['subject']): e for e in old_events_filtered}
                new_dict = {(e['date'], e['time'], e['subject']): e for e in serialized_new}
                
                for key, new_ev in new_dict.items():
                    if key not in old_dict:
                        changes.append(f"✨ **Добавлена новая пара:**\n**Дата:** {new_ev['date']}\n**Время:** {new_ev['time']}\n**Предмет:** {new_ev['subject']}\n**Аудитория:** {new_ev['room']}")
                    else:
                        old_ev = old_dict[key]
                        if old_ev['room'] != new_ev['room']:
                            changes.append(f"🚪 **Изменена аудитория:**\n**Дата:** {new_ev['date']} ({new_ev['time']})\n**Предмет:** {new_ev['subject']}\n**Было:** {old_ev['room']}\n**Стало:** {new_ev['room']}")
                        if old_ev['teacher'] != new_ev['teacher']:
                            changes.append(f"👨‍🏫 **Изменен преподаватель:**\n**Дата:** {new_ev['date']} ({new_ev['time']})\n**Предмет:** {new_ev['subject']}\n**Было:** {old_ev['teacher']}\n**Стало:** {new_ev['teacher']}")
                        if old_ev['end_time'] != new_ev['end_time']:
                            changes.append(f"⏰ **Изменено время окончания:**\n**Дата:** {new_ev['date']} ({new_ev['time']})\n**Предмет:** {new_ev['subject']}\n**Было:** {old_ev['end_time']}\n**Стало:** {new_ev['end_time']}")
                            
                for key, old_ev in old_dict.items():
                    if key not in new_dict:
                        changes.append(f"❌ **Пара отменена:**\n**Дата:** {old_ev['date']}\n**Время:** {old_ev['time']}\n**Предмет:** {old_ev['subject']}")

    # 1. СНАЧАЛА записываем отчет
    if changes:
        changes_dir = os.path.join(BASE_DIR, 'changes')
        os.makedirs(changes_dir, exist_ok=True)
        
        timestamp = datetime.now(TIMEZONE).strftime('%Y-%m-%d_%H-%M-%S')
        filename = os.path.join(changes_dir, f"{timestamp}.md")
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("# Изменения расписания\n\n" + "\n\n---\n\n".join(changes) + "\n")
        print(f"Найдено изменений: {len(changes)}. Отчет сохранен: {filename}")
    elif not report:
         print("Принудительное обновление: отчет об изменениях пропущен.")
    else:
        print("Изменений нет (расписание актуально).")
                    
    # 2. ТОЛЬКО ПОТОМ сохраняем слепок атомарно
    temp_json = json_path + '.tmp'
    with open(temp_json, 'w', encoding='utf-8') as f:
        json.dump(serialized_new, f, ensure_ascii=False, indent=2)
    os.replace(temp_json, json_path)

def create_travel_events(study_events):
    print("Считаем время на дорогу...")
    travel_events = []
    days = {}
    for ev in study_events:
        day = ev['start'].date()
        if day not in days: days[day] = []
        days[day].append(ev)
        
    for day, evs in days.items():
        earliest_start = min(ev['start'] for ev in evs)
        latest_end = max(ev['end'] for ev in evs)
        
        travel_events.append({
            'summary': '🚇 Дорога в вуз', 'location': '', 'description': '',
            'start': earliest_start - timedelta(minutes=TRAVEL_MINUTES),
            'end': earliest_start
        })
        travel_events.append({
            'summary': '🚇 Дорога домой', 'location': '', 'description': '',
            'start': latest_end,
            'end': latest_end + timedelta(minutes=TRAVEL_MINUTES)
        })
    return travel_events

def count_events(path):
    try:
        with open(path, 'rb') as f:
            return sum(1 for _ in Calendar.from_ical(f.read()).walk('VEVENT'))
    except FileNotFoundError:
        return 0
    except Exception as e:
        print(f"⚠️ ПРЕДУПРЕЖДЕНИЕ: Ошибка чтения старого файла {path}: {e}")
        return 0

def make_ics(events, filename, calendar_name, file_stamp):
    print(f"Создаем файл {filename}...")
    cal = Calendar()
    cal.add('prodid', '-//VK Schedule Bot//RU')
    cal.add('version', '2.0')
    cal.add('X-WR-CALNAME', calendar_name)
    cal.add('X-WR-TIMEZONE', 'Europe/Moscow')
    cal.add('X-PUBLISHED-TTL', 'PT6H')
    cal.add('REFRESH-INTERVAL', timedelta(hours=6), parameters={'VALUE': 'DURATION'})

    seen = {}
    for ev in sorted(events, key=lambda e: (e['start'], e['summary'], e['location'], e['description'])):
        event = Event()
        event.add('summary', ev['summary'])
        if ev.get('location'): event.add('location', ev['location'])
        if ev.get('description'): event.add('description', ev['description'])
            
        event.add('dtstart', ev['start'].astimezone(timezone.utc))
        event.add('dtend', ev['end'].astimezone(timezone.utc))
        event.add('dtstamp', file_stamp)
        event.add('last-modified', file_stamp) 
        
        unique_str = f"{ev['start'].astimezone(timezone.utc).isoformat()}-{ev['summary']}-{ev['location']}"
        base_uid = hashlib.sha256(unique_str.encode('utf-8')).hexdigest()[:32]
        seen[base_uid] = seen.get(base_uid, 0) + 1
        uid = f"{base_uid}-{seen[base_uid]}@schedule-bot" if seen[base_uid] > 1 else f"{base_uid}@schedule-bot"
        event.add('uid', uid)
        
        cal.add_component(event)

    temp_file = os.path.join(BASE_DIR, filename + '.tmp')
    final_file = os.path.join(BASE_DIR, filename)
    with open(temp_file, 'wb') as f:
        f.write(cal.to_ical())
    os.replace(temp_file, final_file)
    print(f"Файл {filename} готов. Событий: {len(events)}")

if __name__ == '__main__':
    file_path, file_stamp = download_excel()
    study_events, skipped_dates, total_checked = parse_excel(file_path)
    
    if total_checked > 0:
        print(f"Статистика ячеек с парами: всего {total_checked}, без даты {skipped_dates}")
    
    if len(study_events) < MIN_STUDY_EVENTS:
        sys.exit(f"КРИТИЧЕСКАЯ ОШИБКА: Слишком мало событий ({len(study_events)} < {MIN_STUDY_EVENTS}). Файлы не обновлены.")
        
    if not FORCE_UPDATE and total_checked > 0 and (skipped_dates / total_checked) > 0.3:
        sys.exit(f"КРИТИЧЕСКАЯ ОШИБКА: Больше 30% пар без дат ({skipped_dates} из {total_checked}). Чтобы обновить принудительно, используйте FORCE_UPDATE=1.")
        
    old_file_path = os.path.join(BASE_DIR, 'study.ics')
    old_count = count_events(old_file_path)
    if not FORCE_UPDATE and old_count > 0 and len(study_events) < old_count * 0.5:
        sys.exit(f"КРИТИЧЕСКАЯ ОШИБКА: Событий стало слишком мало ({old_count} -> {len(study_events)}). Чтобы обновить принудительно, используйте FORCE_UPDATE=1.")
        
    make_ics(study_events, 'study.ics', 'Study', file_stamp)
    travel_events = create_travel_events(study_events)
    make_ics(travel_events, 'life.ics', '🚇', file_stamp)
    
    try:
        check_changes(study_events, report=not FORCE_UPDATE)
    except Exception as e:
        print(f"⚠️ Ошибка при проверке изменений: {e}")