import os
import requests
import vk_api
import pandas as pd
from datetime import datetime, timedelta
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo
import hashlib

# НАСТРОЙКИ
VK_TOKEN = os.environ.get('VK_TOKEN')
GROUP_ID = 29490619
TOPIC_ID = 68268274
SHEET_NAME = 'ПЕРК-243'

# Время на дорогу в минутах (45 минут туда, 45 обратно)
TRAVEL_MINUTES = 90 
TIMEZONE = ZoneInfo("Europe/Moscow")

def download_excel():
    print("Ищем файл с расписанием в ВК...")
    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()
    
    comments = vk.board.getComments(group_id=GROUP_ID, topic_id=TOPIC_ID, count=30, sort='desc')
    
    for item in comments['items']:
        if 'attachments' in item:
            for attach in item['attachments']:
                if attach['type'] == 'doc':
                    doc = attach['doc']
                    title = doc['title'].lower()
                    
                    if doc['ext'] in ['xls', 'xlsx']:
                        if ('перк' in title or 'perk' in title) and '3' in title:
                            print(f"Найден подходящий файл: {doc['title']}")
                            response = requests.get(doc['url'])
                            with open('schedule.xls', 'wb') as f:
                                f.write(response.content)
                            return True
    print("Excel файл не найден.")
    return False

def parse_excel():
    print("Парсим расписание...")
    df = pd.read_excel('schedule.xls', sheet_name=SHEET_NAME, header=None)
    
    events = []
    current_dates = []
    
    for idx, row in df.iterrows():
        val_3 = str(row[3]).strip()
        
        # 1. Ищем строку с датами
        if '2026' in val_3 or '2027' in val_3:
            current_dates = []
            for col_idx in range(3, 21):
                cell_val = row[col_idx]
                if pd.isna(cell_val):
                    current_dates.append(None)
                else:
                    try:
                        d = pd.to_datetime(cell_val).date()
                        current_dates.append(d)
                    except:
                        current_dates.append(None)
            continue
            
        time_cell = str(row[2]).strip()
        
        # 2. Ищем строку со временем
        if '-' in time_cell and '.' in time_cell and current_dates:
            try:
                start_time_str, end_time_str = time_cell.split('-')
                if start_time_str.strip() == "00.00":
                    continue
                    
                start_time = datetime.strptime(start_time_str.strip(), "%H.%M").time()
                end_time = datetime.strptime(end_time_str.strip(), "%H.%M").time()
            except Exception:
                continue 

            # 3. Извлекаем данные из колонок
            for i, col_idx in enumerate(range(3, 21)):
                cell_value = str(row[col_idx]).strip()
                
                if cell_value and cell_value != 'nan' and 'самостоятельной' not in cell_value.lower():
                    parts = cell_value.replace('<br>', '\n').split('\n')
                    subject = parts[0].strip()
                    
                    teacher = ""
                    room = ""
                    
                    if len(parts) > 1:
                        # Разделяем по последнему дефису (чтобы не сломать двойные фамилии)
                        info = parts[1].rsplit('-', 1)
                        if len(info) == 2:
                            teacher = info[0].strip()
                            room = info[1].strip()
                        else:
                            teacher = parts[1].strip()

                    event_date = current_dates[i]
                    if event_date is None:
                        continue
                        
                    start_dt = datetime.combine(event_date, start_time, tzinfo=TIMEZONE)
                    end_dt = datetime.combine(event_date, end_time, tzinfo=TIMEZONE)
                    
                    events.append({
                        'summary': subject,
                        'location': room,
                        'description': teacher,  # <--- Теперь преподаватель сохраняется
                        'start': start_dt,
                        'end': end_dt
                    })
                    
    return events

def create_travel_events(study_events):
    print("Считаем время на дорогу...")
    travel_events = []
    
    # Группируем занятия по дням
    days = {}
    for ev in study_events:
        day = ev['start'].date()
        if day not in days:
            days[day] = []
        days[day].append(ev)
        
    for day, evs in days.items():
        # Находим самое раннее начало и самый поздний конец в этот день
        earliest_start = min(ev['start'] for ev in evs)
        latest_end = max(ev['end'] for ev in evs)
        
        # Дорога ТУДА (заканчивается ровно к началу первой пары)
        travel_to_start = earliest_start - timedelta(minutes=TRAVEL_MINUTES)
        travel_to_end = earliest_start
        
        travel_events.append({
            'summary': '🚇 Дорога в вуз',
            'location': '',
            'start': travel_to_start,
            'end': travel_to_end
        })
        
        # Дорога ОБРАТНО (начинается ровно после окончания последней пары)
        travel_from_start = latest_end
        travel_from_end = latest_end + timedelta(minutes=TRAVEL_MINUTES)
        
        travel_events.append({
            'summary': '🚇 Дорога домой',
            'location': '',
            'start': travel_from_start,
            'end': travel_from_end
        })
        
    return travel_events

def make_ics(events, filename, calendar_name):
    print(f"Создаем файл {filename}...")
    cal = Calendar()
    cal.add('prodid', '-//VK Schedule Bot//RU')
    cal.add('version', '2.0')
    cal.add('X-WR-CALNAME', calendar_name)
    cal.add('X-WR-TIMEZONE', 'Europe/Moscow')

    for ev in events:
        event = Event()
        event.add('summary', ev['summary'])
        
        if ev.get('location'):
            event.add('location', ev['location'])
            
        if ev.get('description'):  # <--- Добавляем преподавателя в описание
            event.add('description', ev['description'])
            
        event.add('dtstart', ev['start'])
        event.add('dtend', ev['end'])
        
        # Генерируем постоянный UID
        unique_str = f"{ev['start'].isoformat()}-{ev['summary']}"
        uid_hash = hashlib.md5(unique_str.encode('utf-8')).hexdigest()
        event.add('uid', f"{uid_hash}@schedule-bot")
        
        cal.add_component(event)

    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    print(f"Файл {filename} готов. Событий: {len(events)}")

if __name__ == '__main__':
    if download_excel():
        study_events = parse_excel()
        if study_events:
            # 1. Генерируем календарь для учебы
            make_ics(study_events, 'study.ics', 'Study')
            
            # 2. Генерируем календарь для дороги
            travel_events = create_travel_events(study_events)
            make_ics(travel_events, 'life.ics', '🚇')
        else:
            print("Не удалось найти занятия.")
