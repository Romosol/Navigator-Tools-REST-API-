"""
Navigator App: Автономное настольное приложение с графическим интерфейсом (GUI)
для автоматизации Навигатора дополнительного образования (NavAdd и NavConfirm).

Особенности:
1. Сохранение логина/пароля в config.json с возможностью быстрого входа под прошлым пользователем.
2. Авто-создание и открытие готовой таблицы list.xlsx с тестовыми данными.
3. Кнопки быстрого открытия config.json и list.xlsx в проводнике/Excel.
4. Понятный графический интерфейс на Tkinter (без сторонних тяжелых GUI библиотек).
5. Строгая проверка 'is_approved': true и даты рождения (дд.мм.гг / дд.мм.гггг).
"""

import os
import sys
import json
import time
import shutil
import subprocess
import threading
import webbrowser
import platform
import urllib.parse
from typing import Optional, Tuple, Dict, Any, List, Callable, Union
import requests

APP_VERSION = "2.1.0"
GITHUB_REPO = "Romosol/Navigator-Tools-REST-API-"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"
GITHUB_ISSUES_URL = f"https://github.com/{GITHUB_REPO}/issues"
GITHUB_NEW_ISSUE_URL = f"https://github.com/{GITHUB_REPO}/issues/new"
GITHUB_ISSUES_API = f"https://api.github.com/repos/{GITHUB_REPO}/issues"

CONFIG_FILE = "config.json"
EVENT_EXCEL_FILE = "event_list.xlsx"
PROGRAM_EXCEL_FILE = "programm_list.xlsx"
STUDY_EXCEL_FILE = "study_list.xlsx"
LEGACY_EXCEL_FILE = "list.xlsx"
EXCEL_FILE = EVENT_EXCEL_FILE  # Алиас для обратной совместимости
LOGS_DIR = "logs"
SESSION_START_TIME_STR = time.strftime("%Y-%m-%d_%H-%M-%S")
CURRENT_LOG_FILE = os.path.join(LOGS_DIR, f"results_log_{SESSION_START_TIME_STR}.txt")
LOG_FILE = CURRENT_LOG_FILE  # Указывает на текущий файл запуска внутри папки logs/


def parse_version_tuple(ver_str: str) -> Tuple[int, ...]:
    """Преобразует строку версии (например 'v2.2.0', '2.1') в кортеж чисел (2, 2, 0)."""
    if not ver_str:
        return (0,)
    s = str(ver_str).strip().lstrip("vV")
    for sep in ["-", "+", " ", "_"]:
        if sep in s:
            s = s.split(sep, 1)[0]
    parts = []
    for p in s.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        if digits:
            parts.append(int(digits))
    return tuple(parts) if parts else (0,)


def is_newer_version(remote_ver: str, local_ver: str) -> bool:
    """Сравнивает версии и возвращает True, если remote_ver строго новее local_ver."""
    rem = parse_version_tuple(remote_ver)
    loc = parse_version_tuple(local_ver)
    max_len = max(len(rem), len(loc))
    rem_padded = rem + (0,) * (max_len - len(rem))
    loc_padded = loc + (0,) * (max_len - len(loc))
    return rem_padded > loc_padded


def normalize_academic_year_id(val: Any) -> str:
    """
    Преобразует ввод учебного года в числовой academic_year_id.
    Поддерживает ввод обоих форматов:
      - '2026/2027' -> '2026'
      - '2026-2027' -> '2026'
      - '2026 / 2027' -> '2026'
      - '2026' -> '2026'
      - 2026 -> '2026'
    """
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    for sep in ["/", "-", "—", "–", "\\"]:
        if sep in s:
            part = s.split(sep)[0].strip()
            digits = "".join([c for c in part if c.isdigit()])
            if len(digits) == 4:
                return digits
    digits = "".join([c for c in s if c.isdigit()])
    if len(digits) >= 4:
        return digits[:4]
    return s


def normalize_date_input(raw: Any, default_val: str = "") -> str:
    """
    Нормализует ввод даты в формат YYYY-MM-DD.
    Поддерживает форматы: дд.мм.гг, дд.мм.гггг, гггг-мм-дд, с любыми разделителями (/ или . или -)
    """
    if raw is None:
        return default_val
    if hasattr(raw, 'strftime'):
        return raw.strftime('%Y-%m-%d')
    s = str(raw).strip()
    if not s or s.lower() in ['none', 'nan', 'nat', '']:
        return default_val
    if ' ' in s:
        s = s.split(' ')[0]
    s_clean = s.replace('/', '.').replace('-', '.')
    parts = [p.strip() for p in s_clean.split('.') if p.strip()]
    if len(parts) == 3:
        if len(parts[0]) == 4 and parts[0].isdigit():
            return f'{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}'
        if parts[0].isdigit() and parts[1].isdigit() and parts[2].isdigit():
            d = parts[0].zfill(2)
            m = parts[1].zfill(2)
            y_raw = parts[2]
            if len(y_raw) == 2:
                y_num = int(y_raw)
                full_y = 2000 + y_num if y_num <= 35 else 1900 + y_num
                return f'{full_y}-{m}-{d}'
            elif len(y_raw) == 4:
                return f'{y_raw}-{m}-{d}'
    return s


def normalize_fio_tokens(fio: Any) -> List[str]:
    """Разбивает ФИО на чистые токены (буквенные слова в нижнем регистре с заменой 'ё' на 'е')."""
    if not fio:
        return []
    import re
    s = str(fio).strip().lower().replace("ё", "е")
    clean = re.sub(r"[^a-zа-я0-9\s\-]", " ", s)
    return [w.strip("-") for w in clean.split() if w.strip("-")]


def check_fio_format(fio: Any) -> Tuple[bool, str]:
    """
    Проверяет формат записи ФИО ребенка из таблицы Excel:
    Возвращает (is_valid: bool, error_reason: str).
    - Если в Excel указаны Фамилия и инициалы (например: 'Иванов И.И.', 'Иванов И. И.', 'Иванов И.', 'Иванов И И', 'И.И. Иванов'),
      отмечает это как неправильный формат (требуются Фамилия и полное Имя).
    - Если указана только фамилия без имени ('Иванов'), возвращает ошибку отсутствия имени.
    """
    if not fio:
        return False, "пустая строка ФИО"

    raw = str(fio).strip()
    if not raw:
        return False, "пустая строка ФИО"

    import re

    # 1. Проверка на инициалы с точкой (например 'И.', 'И.И.', 'А. С.', 'И.В.')
    if re.search(r"\b[a-zA-Zа-яА-ЯёЁ]\.", raw):
        return False, "неправильный формат: указаны инициалы с точкой вместо полного имени (например: 'Иванов И.И.')"

    # 2. Очищаем от знаков пунктуации и делим на слова
    clean = re.sub(r"[^a-zA-Zа-яА-ЯёЁ0-9\-]", " ", raw)
    words = [w.strip("-") for w in clean.split() if w.strip("-")]

    if len(words) < 2:
        return False, "неправильный формат: указана только фамилия без имени"

    # 3. Проверка на однобуквенные токены (инициалы без точки, например 'Иванов И И' или 'Иванов И')
    for w in words:
        if len(w) == 1:
            return False, f"неправильный формат: однобуквенный инициал '{w}' вместо полного имени"

    return True, ""


def extract_last_first(val: Any) -> Optional[Tuple[str, str]]:
    """
    Извлекает нормализованные (Фамилия, Имя) из строки ФИО или словаря заявки Навигатора.
    Сайт Навигатора выдает Фамилию и Имя (поля kid_last_name, kid_first_name или строка site_user_fio).
    В таблице Excel может быть Фамилия Имя Отчество, либо Фамилия Имя.
    Отчество игнорируется, берутся только Фамилия и Имя ребенка.
    Возвращает (last_name, first_name) в нижнем регистре с заменой 'ё' на 'е'.
    Если передан некорректный формат (инициалы, < 2 слов), возвращает None.
    """
    if isinstance(val, dict):
        last_raw = str(val.get("kid_last_name") or "").strip()
        first_raw = str(val.get("kid_first_name") or "").strip()
        if last_raw and first_raw:
            l_toks = normalize_fio_tokens(last_raw)
            f_toks = normalize_fio_tokens(first_raw)
            if l_toks and f_toks:
                if len(l_toks[0]) >= 2 and len(f_toks[0]) >= 2:
                    return l_toks[0], f_toks[0]
        # Резервный источник: server_fio или site_user_fio
        fio_fallback = val.get("server_fio") or val.get("site_user_fio") or ""
        return extract_last_first(fio_fallback)

    if not val:
        return None

    s = str(val).strip()
    is_valid, _ = check_fio_format(s)
    if not is_valid:
        return None

    toks = normalize_fio_tokens(s)
    if len(toks) >= 2:
        # В русском ФИО первые два слова обычно Фамилия и Имя
        # (или Имя и Фамилия, перестановка поддерживается в match_fio)
        return toks[0], toks[1]

    return None


def match_fio(val_a: Any, val_b: Any) -> bool:
    """
    Сравнивает ребенка с сайта Навигатора со списком в Excel:
    1. Сравниваются именно Фамилия и Имя:
       - Сайт выдает Фамилию и Имя (в заявках: kid_last_name, kid_first_name или site_user_fio).
       - В таблице Excel может быть: Фамилия Имя Отчество, либо Фамилия Имя.
       - Отчество в Excel игнорируется при сверке.
    2. Если в Excel указаны Фамилия и инициалы ('Иванов И.И.', 'Иванов И.') — они считаются
       неправильным форматом и совпадение отклоняется до исправления на полное имя.
    3. Поддерживает перестановку: Фамилия Имя == Имя Фамилия.
    4. Защищен от разного регистра, пробелов и различий е/ё (Артём == Артем).
    Возвращает True, если Фамилия и Имя совпадают.
    """
    lf_a = extract_last_first(val_a)
    lf_b = extract_last_first(val_b)
    if not lf_a or not lf_b:
        return False

    a_last, a_first = lf_a
    b_last, b_first = lf_b

    # 1. Прямое совпадение: Фамилия == Фамилия, Имя == Имя
    if a_last == b_last and a_first == b_first:
        return True

    # 2. Совпадение с перестановкой: Фамилия Имя == Имя Фамилия
    if a_last == b_first and a_first == b_last:
        return True

    # 3. Проверка вхождений токенов Фамилии и Имени (для нестандартного порядка слов)
    str_a = val_a.get("server_fio") if isinstance(val_a, dict) else str(val_a)
    str_b = val_b.get("server_fio") if isinstance(val_b, dict) else str(val_b)
    toks_a = normalize_fio_tokens(str_a)
    toks_b = normalize_fio_tokens(str_b)
    if len(toks_a) >= 2 and len(toks_b) >= 2:
        if set(toks_a[:2]) == set(toks_b[:2]):
            return True

    return False


# =========================================================================
# 1. СЕРВИСНЫЕ ФУНКЦИИ (КОНФИГ, ЛОГИ И EXCEL)
# =========================================================================
def init_session_log_file() -> str:
    """
    Инициализирует новый лог-файл при каждом запуске/перезапуске программы.
    Имя файла включает текущую дату и время (например: results_log_2026-09-14_11-15-30.txt).
    Все логи сохраняются ИСКЛЮЧИТЕЛЬНО в папке 'logs/'. В корневой директории файлы не создаются.
    """
    global CURRENT_LOG_FILE, LOG_FILE, SESSION_START_TIME_STR
    SESSION_START_TIME_STR = time.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"results_log_{SESSION_START_TIME_STR}.txt"

    try:
        os.makedirs(LOGS_DIR, exist_ok=True)
    except Exception:
        pass

    # Если в корневой директории остались старые файлы логов от прошлых версий — переносим их в logs/
    try:
        for item in os.listdir("."):
            if item.startswith("results_log") and item.endswith(".txt") and os.path.isfile(item):
                dst = os.path.join(LOGS_DIR, item)
                try:
                    if not os.path.exists(dst):
                        shutil.move(item, dst)
                    else:
                        os.remove(item)
                except Exception:
                    pass
    except Exception:
        pass

    CURRENT_LOG_FILE = os.path.join(LOGS_DIR, filename)
    LOG_FILE = CURRENT_LOG_FILE

    header = (
        "=" * 85 + "\n"
        "НАВИГАТОР ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ — ЖУРНАЛ СЕССИИ\n"
        f"Файл сессии:          {filename}\n"
        f"Папка сохранения:     {os.path.abspath(LOGS_DIR)}\n"
        f"Дата и время запуска: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"Рабочая директория:   {os.path.abspath('.')}\n"
        "Логирование активности всех 4-х вкладок:\n"
        "  • Вкладка 1: Пакетная запись на мероприятие (event_list.xlsx)\n"
        "  • Вкладка 2: Подтверждение заявок и отметка участия в мероприятии\n"
        "  • Вкладка 3: Зачисление на учебные программы (programm_list.xlsx)\n"
        "  • Вкладка 4: Сверка со study_list.xlsx, подтверждение и зачисление на обучение\n"
        + "=" * 85 + "\n\n"
    )

    # Сохраняем ТОЛЬКО в папке logs/
    targets = [CURRENT_LOG_FILE, os.path.join(LOGS_DIR, "results_log.txt")]
    for p in targets:
        try:
            with open(p, "w", encoding="utf-8") as f:
                f.write(header)
        except Exception as e:
            print(f"Ошибка создания лога {p}: {e}")

    return CURRENT_LOG_FILE


def write_to_log_file(text: str, with_timestamp: bool = False) -> None:
    """Безопасно дописывает запись в файлы журнала ИСКЛЮЧИТЕЛЬНО внутри папки logs/ в UTF-8."""
    global CURRENT_LOG_FILE
    if not CURRENT_LOG_FILE or not os.path.exists(CURRENT_LOG_FILE):
        init_session_log_file()

    line = str(text)
    if with_timestamp and line.strip():
        stripped = line.strip()
        # Добавляем временную метку [ЧЧ:ММ:СС] только если строка еще не начинается с метки или разделителя
        if not (stripped.startswith("[") or stripped.startswith("=") or stripped.startswith("-") or stripped.startswith("!")):
            line = f"[{time.strftime('%H:%M:%S')}] {line}"

    targets = [CURRENT_LOG_FILE, os.path.join(LOGS_DIR, "results_log.txt")]
    for p in targets:
        try:
            with open(p, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


def ensure_log_file() -> None:
    """Создает новый уникальный файл журнала текущей сессии в папке logs/ при запуске."""
    global CURRENT_LOG_FILE
    if not CURRENT_LOG_FILE or not os.path.exists(CURRENT_LOG_FILE):
        init_session_log_file()
DEFAULT_CONFIG: Dict[str, Any] = {
    "email": "",
    "password": "",
    "saved_at": "",
    "activity_name": "Название мероприятия",
    "activity_datetime": "01.09.2026 11:00",
    "confirm_activity_name": "Название мероприятия",
    "program_name": "Название программы",
    "program_event_id": "",
    "program_group_name": "Название группы",
    "program_group_id": "",
    "program_academic_year_id": "2026/2027",
    "program_use_certificate": False,
    "program_create_certificate": True,
    "study_program_name": "Название программы",
    "study_program_group_name": "Название группы",
    "study_academic_year_id": "2026/2027",
    "study_decree_number": "183",
    "study_date_signing": "31.08.2026",
    "study_date_start": "01.09.2026",
    "study_financing_source": "1",
    "study_verify_excel": True,
    "check_updates": True,
    "github_token": "",
    "history": {
        "email": [],
        "activity_name": [],
        "activity_datetime": ["01.09.2026 11:00", "2026-08-31 10:00:00", "2026-09-01 12:00:00"],
        "program_name": [],
        "program_group_name": [],
        "academic_year_id": ["2026/2027", "2025/2026", "2026"],
        "decree_number": ["183", "124", "1"],
        "date_signing": ["31.08.2026", "2026-08-31", "01.09.2026"],
        "date_start": ["01.09.2026", "2026-09-01", "15.09.2026"],
    }
}


def load_config() -> Dict[str, Any]:
    """Загружает конфиг, объединяя сохраненные параметры с дефолтными значениями-примерами и историей."""
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    loaded_hist = loaded.get("history", {})
                    base_hist = dict(DEFAULT_CONFIG.get("history", {}))
                    if isinstance(loaded_hist, dict):
                        base_hist.update(loaded_hist)
                    cfg.update(loaded)
                    cfg["history"] = base_hist
        except Exception:
            pass
    else:
        # Если конфига еще нет — сразу создаем файл с валидными примерами
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return cfg


def get_field_history(field_key: str, max_items: int = 30) -> List[str]:
    """Возвращает упорядоченный список ранее использованных значений для поля из config.json."""
    cfg = load_config()
    history = cfg.get("history", {})
    if not isinstance(history, dict):
        history = {}
    items = history.get(field_key, [])
    if not isinstance(items, list):
        items = []

    res: List[str] = []

    # 1. Сначала добавляем элементы из истории в порядке недавнего использования (MRU)
    for it in items:
        s = str(it).strip()
        if s and s not in res and not s.startswith("Название "):
            res.append(s)

    # 2. Подтягиваем текущие значения из config.json (включая альтернативные ключи старых версий)
    field_to_cfg_keys = {
        "activity_name": ["activity_name", "confirm_activity_name"],
        "activity_datetime": ["activity_datetime"],
        "program_name": ["study_program_name", "program_name"],
        "program_group_name": ["study_program_group_name", "program_group_name"],
        "academic_year_id": ["study_academic_year_id", "program_academic_year_id"],
        "decree_number": ["study_decree_number"],
        "date_signing": ["study_date_signing"],
        "date_start": ["study_date_start"],
        "email": ["email"]
    }
    check_keys = field_to_cfg_keys.get(field_key, [field_key])
    for ck in check_keys:
        curr_val = str(cfg.get(ck, "")).strip()
        if curr_val and curr_val not in res and not curr_val.startswith("Название "):
            res.append(curr_val)

    # 3. Если список пуст или мал, подмешиваем дефолтные подсказки
    def_hist = DEFAULT_CONFIG.get("history", {}).get(field_key, [])
    for d in def_hist:
        sd = str(d).strip()
        if sd and sd not in res and not sd.startswith("Название "):
            res.append(sd)

    return res[:max_items]


def add_to_field_history(field_key: str, value: str, max_items: int = 30) -> List[str]:
    """Добавляет новое значение в историю поля в config.json (ставит его наверх списка MRU)."""
    val = str(value).strip()
    if not val or val.startswith("Название "):
        return get_field_history(field_key, max_items)

    cfg = load_config()
    history = cfg.get("history", {})
    if not isinstance(history, dict):
        history = {}

    curr_list = history.get(field_key, [])
    if not isinstance(curr_list, list):
        curr_list = []

    # Перемещаем значение на первое место без дублирования
    new_list = [val] + [str(x).strip() for x in curr_list if str(x).strip() != val and str(x).strip() and not str(x).strip().startswith("Название ")]
    history[field_key] = new_list[:max_items]
    cfg["history"] = history
    cfg["saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Ошибка сохранения истории в config.json: {e}")

    return history[field_key]


def save_config(email: str, password: str) -> None:
    """Сохраняет учетные данные пользователя в config.json, сохраняя остальные поля."""
    update_config(email=email, password=password)


def update_config(**kwargs) -> None:
    """Обновляет и сохраняет указанные поля в config.json."""
    cfg = load_config()
    cfg.update(kwargs)
    cfg["saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Ошибка сохранения config.json: {e}")


def open_file_in_os(file_path: str):
    """Открывает файл в ассоциированной программе ОС (Excel, Блокнот)."""
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        return False, f"Файл {file_path} не найден"
    try:
        if sys.platform.startswith("win"):
            os.startfile(abs_path)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", abs_path])
        else:
            subprocess.Popen(["xdg-open", abs_path])
        return True, "Файл открыт"
    except Exception as e:
        return False, f"Не удалось открыть файл: {e}"


def ensure_default_excel(file_path: str = EVENT_EXCEL_FILE):
    """Создает шаблонный файл Excel, если его нет, либо добавляет третий столбец статуса."""
    try:
        import openpyxl
        # Если целевого файла нет, но есть старый list.xlsx — копируем его
        if not os.path.exists(file_path):
            if os.path.exists(LEGACY_EXCEL_FILE):
                try:
                    import shutil
                    shutil.copy(LEGACY_EXCEL_FILE, file_path)
                except Exception:
                    pass

        if not os.path.exists(file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            # Заголовки (3 столбца)
            ws.cell(row=1, column=1, value="ФИО")
            ws.cell(row=1, column=2, value="Дата рождения")
            ws.cell(row=1, column=3, value="Статус заявки")
            # Пример строки
            ws.cell(row=2, column=1, value="Фамилия Имя Отчество")
            ws.cell(row=2, column=2, value="22.08.1998")
            ws.cell(row=2, column=3, value="")  # Пустое: ребенок еще не обработан
            ws.column_dimensions["A"].width = 38
            ws.column_dimensions["B"].width = 20
            ws.column_dimensions["C"].width = 28
            wb.save(file_path)
            wb.close()
        else:
            # Если файл существует, проверяем наличие столбца со статусом
            try:
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active
                if ws.max_column < 3 or not ws.cell(1, 3).value:
                    ws.cell(row=1, column=3, value="Статус заявки")
                    ws.column_dimensions["C"].width = 28
                    wb.save(file_path)
                wb.close()
            except Exception:
                pass
    except Exception as e:
        print(f"Не удалось инициализировать {file_path}: {e}")


def ensure_study_excel(file_path: str = STUDY_EXCEL_FILE):
    """
    Создает или проверяет отдельную таблицу для подтверждения и отметки на обучение (study_list.xlsx).
    Содержит ровно 2 столбца:
      1. ФИО (Полное имя ребенка)
      2. Флаг зачисления (переключается при подтверждении и зачислении на программу)
    """
    try:
        import openpyxl
        if not os.path.exists(file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Обучение"
            # Ровно 2 столбца
            ws.cell(row=1, column=1, value="ФИО")
            ws.cell(row=1, column=2, value="Флаг зачисления")

            # Если в programm_list.xlsx уже есть дети, переносим их ФИО для удобства пользователя
            seeded = False
            if os.path.exists(PROGRAM_EXCEL_FILE):
                try:
                    p_wb = openpyxl.load_workbook(PROGRAM_EXCEL_FILE, data_only=True)
                    p_ws = p_wb.active
                    r_out = 2
                    for r_in in range(2, p_ws.max_row + 1):
                        val = p_ws.cell(r_in, 1).value
                        if val and str(val).strip():
                            ws.cell(row=r_out, column=1, value=str(val).strip())
                            ws.cell(row=r_out, column=2, value="")
                            r_out += 1
                            seeded = True
                    p_wb.close()
                except Exception:
                    pass

            if not seeded:
                ws.cell(row=2, column=1, value="")
                ws.cell(row=2, column=2, value="")

            ws.column_dimensions["A"].width = 42
            ws.column_dimensions["B"].width = 34
            wb.save(file_path)
            wb.close()
        else:
            # Если файл существует, проверяем наличие 2-го столбца с флагом
            try:
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active
                if ws.max_column < 2 or not ws.cell(1, 2).value:
                    ws.cell(row=1, column=2, value="Флаг зачисления")
                    ws.column_dimensions["B"].width = 34
                    wb.save(file_path)
                wb.close()
            except Exception:
                pass
    except Exception as e:
        print(f"Не удалось инициализировать {file_path}: {e}")


def ensure_all_excel_files():
    """Гарантирует существование всех таблиц: event_list.xlsx, programm_list.xlsx и study_list.xlsx."""
    ensure_default_excel(EVENT_EXCEL_FILE)
    ensure_default_excel(PROGRAM_EXCEL_FILE)
    ensure_study_excel(STUDY_EXCEL_FILE)


def update_excel_cell_status(file_path: str, row_num: int, status_text: str, col_num: int = 3) -> Tuple[bool, str]:
    """
    Записывает флаг/статус успешного создания заявки в ячейку Excel.
    Автоматически сохраняет и закрывает книгу.
    В случае блокировки файла (например, открыт в Microsoft Excel) возвращает предупреждение без падения.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        # Проверяем заголовок столбца статуса
        col_letter = openpyxl.utils.get_column_letter(col_num)
        if not ws.cell(1, col_num).value:
            ws.cell(1, col_num, value="Статус заявки")
            ws.column_dimensions[col_letter].width = 28

        ws.cell(row=row_num, column=col_num, value=status_text)
        wb.save(file_path)
        wb.close()
        return True, "Статус записан в Excel"
    except PermissionError:
        fname = os.path.basename(file_path)
        return False, f"Файл {fname} открыт в Microsoft Excel или другой программе (файл заблокирован)"
    except Exception as e:
        return False, str(e)


# =========================================================================
# 2. АДАПТИВНАЯ ЗАЩИТА ОТ ПЕРЕГРУЗКИ СЕРВЕРА (SMART RATE LIMITER)
# =========================================================================
class AdaptiveRateLimiter:
    """
    Адаптивный контроллер частоты запросов с защитой от перегрузки сервера (Smart Backoff).
    - Базовая задержка: не менее 1.15 - 1.25 секунды (строго <= 1 запрос в секунду).
    - При росте времени отклика сервера задержка динамически увеличивается (до 2.5 - 4.5 сек).
    - При получении HTTP 429/502/503/504 включается аварийная защитная пауза (6 - 10 сек).
    """
    def __init__(self, min_interval: float = 1.2, abort_checker: Optional[Callable[[], bool]] = None):
        # Гарантируем, что интервал не может быть меньше 1.0 с (не более 1 запроса в секунду)
        self.min_interval = max(1.0, float(min_interval))
        self.current_delay = self.min_interval
        self.last_request_time = 0.0
        self.last_latency_ms: Optional[float] = None
        self.last_status_code: int = 200
        self.abort_checker = abort_checker

    def wait_before_request(self) -> float:
        """Блокирует поток на необходимое время, гарантируя соблюдение паузы между запросами (с мгновенным выходом при прерывании)."""
        now = time.time()
        elapsed = now - self.last_request_time
        needed_wait = self.current_delay - elapsed
        if needed_wait > 0:
            step = 0.05
            while needed_wait > 0:
                if self.abort_checker and self.abort_checker():
                    break
                time.sleep(min(step, needed_wait))
                needed_wait -= step
        self.last_request_time = time.time()
        return self.current_delay

    def record_response(self, latency_ms: float, status_code: int = 200) -> Tuple[str, float, str]:
        """
        Анализирует отклик сервера и корректирует задержку.
        Возвращает: (описание_нагрузки, новая_задержка, цвет_индикатора)
        """
        self.last_latency_ms = latency_ms
        self.last_status_code = status_code

        # 1. Сервер вернул ошибку перегрузки / rate limit
        if status_code in (429, 502, 503, 504):
            self.current_delay = min(10.0, max(6.0, self.current_delay * 2.0))
            return f"ПЕРЕГРУЗКА (HTTP {status_code})", round(self.current_delay, 2), "#ef4444"

        # 2. Адаптация по времени отклика (Latency)
        if latency_ms < 600:
            # Сервер свободен: постепенно возвращаемся к базовой бережной задержке (1.2 с)
            self.current_delay = max(self.min_interval, self.current_delay * 0.95)
            return "Низкая (Сервер свободен)", round(self.current_delay, 2), "#10b981"
        elif latency_ms < 1500:
            # Умеренная нагрузка: держим задержку ~1.2 - 1.5 с
            self.current_delay = max(self.min_interval, 1.35)
            return "Умеренная (Норма)", round(self.current_delay, 2), "#3b82f6"
        elif latency_ms < 3000:
            # Повышенная нагрузка: сервер замедляется, плавно увеличиваем интервал
            self.current_delay = max(2.0, min(4.0, self.current_delay + 0.4))
            return "Повышенная (Замедление)", round(self.current_delay, 2), "#f59e0b"
        else:
            # Сервер перегружен: задержка увеличивается до 3.5 - 5.0 с
            self.current_delay = max(3.5, min(6.0, self.current_delay + 0.8))
            return "Критическая (Высокая нагрузка)", round(self.current_delay, 2), "#ef4444"


# =========================================================================
# 3. REST-КЛИЕНТ НАВИГАТОРА
# =========================================================================
class NavigatorClient:
    def __init__(self, base_url: str = "https://xn--02-6kcatyook.xn--80aafey1amqq.xn--d1acj3b"):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "*/*",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Content-Type": "application/json",
            "Origin": self.base_url,
            "Referer": f"{self.base_url}/admin/",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
            "X-Skip-400-Error-Window": "1",
        })
        self.access_token: Optional[str] = None
        self.user_info: Dict[str, Any] = {}

        # Адаптивный контроллер (строго <= 1 запрос в секунду, базовая задержка 1.2 с)
        self.limiter = AdaptiveRateLimiter(min_interval=1.2)
        self.on_ping_update: Optional[Callable[[float, str, str, float], None]] = None
        self.on_log_message: Optional[Callable[[str, Optional[str]], None]] = None
        self.abort_checker: Optional[Callable[[], bool]] = None

    def is_aborted(self) -> bool:
        """Проверяет, запросил ли пользователь остановку операции."""
        if hasattr(self, "abort_checker") and self.abort_checker and self.abort_checker():
            return True
        return False

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Выполняет запрос с соблюдением адаптивного лимита скорости, перехватом задержки и проверкой прерывания."""
        if self.is_aborted():
            raise InterruptedError("Операция прервана пользователем")

        max_retries = 2
        for attempt in range(max_retries + 1):
            if self.is_aborted():
                raise InterruptedError("Операция прервана пользователем")

            # Соблюдаем безопасную паузу: не более 1 запроса в секунду
            cur_delay = self.limiter.wait_before_request()

            if self.is_aborted():
                raise InterruptedError("Операция прервана пользователем")

            start_t = time.perf_counter()
            try:
                resp = self.session.request(method, url, **kwargs)
                lat_ms = (time.perf_counter() - start_t) * 1000.0
                label, new_delay, color = self.limiter.record_response(lat_ms, resp.status_code)

                if self.on_ping_update:
                    try:
                        self.on_ping_update(lat_ms, label, color, new_delay)
                    except Exception:
                        pass

                # Если сервер перегружен (429 Too Many Requests или 502/503/504)
                if resp.status_code in (429, 502, 503, 504) and attempt < max_retries:
                    wait_time = max(6.0, new_delay)
                    if self.on_log_message:
                        self.on_log_message(
                            f"⚠️ [Защита сервера] Код {resp.status_code}. Пауза {wait_time:.1f}с перед повторной попыткой...",
                            "yellow"
                        )
                    step = 0.1
                    rem = wait_time
                    while rem > 0:
                        if self.is_aborted():
                            raise InterruptedError("Операция прервана пользователем")
                        time.sleep(min(step, rem))
                        rem -= step
                    continue

                return resp
            except Exception as exc:
                if isinstance(exc, InterruptedError):
                    raise exc
                lat_ms = (time.perf_counter() - start_t) * 1000.0
                label, new_delay, color = self.limiter.record_response(lat_ms, 504)
                if self.on_ping_update:
                    try:
                        self.on_ping_update(lat_ms, label, color, new_delay)
                    except Exception:
                        pass
                if attempt < max_retries:
                    step = 0.1
                    rem = 3.0
                    while rem > 0:
                        if self.is_aborted():
                            raise InterruptedError("Операция прервана пользователем")
                        time.sleep(min(step, rem))
                        rem -= step
                    continue
                raise exc
        return resp

    def check_ping(self) -> Tuple[bool, float, str, str, float]:
        """
        Замеряет отклик (RTT/latency) сервера Навигатора в реальном времени.
        Возвращает: (успех, latency_ms, статус_нагрузки, цвет_индикатора, задержка)
        """
        url = f"{self.base_url}/api/rest/activity/"
        start_t = time.perf_counter()
        try:
            # Легкий запрос с длиной 1 для минимальной нагрузки
            resp = self.session.get(url, params={"length": 1, "page": 1}, timeout=10)
            lat_ms = (time.perf_counter() - start_t) * 1000.0
            label, delay, color = self.limiter.record_response(lat_ms, resp.status_code)
            if self.on_ping_update:
                try:
                    self.on_ping_update(lat_ms, label, color, delay)
                except Exception:
                    pass
            return True, lat_ms, label, color, delay
        except Exception as e:
            lat_ms = (time.perf_counter() - start_t) * 1000.0
            label, delay, color = self.limiter.record_response(lat_ms, 504)
            if self.on_ping_update:
                try:
                    self.on_ping_update(lat_ms, f"Сбой связи ({e})", color, delay)
                except Exception:
                    pass
            return False, lat_ms, f"Сбой связи ({e})", color, delay

    def login(self, email: str, password: str) -> Tuple[bool, str]:
        url = f"{self.base_url}/api/user/login"
        payload = {"email": email.strip(), "password": password.strip()}
        try:
            resp = self.session.post(url, json=payload, timeout=15)
            data = resp.json()
            if data.get("success") and "access_token" in data.get("data", {}):
                self.access_token = data["data"]["access_token"]
                self.user_info = data["data"].get("user", {})
                self.session.headers["Authorization"] = f"Bearer {self.access_token}"
                name = self.user_info.get("name") or self.user_info.get("email") or email
                return True, f"Успешный вход: {name}"
            msg = data.get("message") or data.get("errors") or "Неверный логин или пароль"
            return False, str(msg)
        except Exception as e:
            return False, f"Ошибка сети при авторизации: {e}"

    def search_activity(self, activity_name: str) -> Tuple[Optional[Dict[str, Any]], str]:
        url = f"{self.base_url}/api/rest/activity/"
        params = {"query": activity_name.strip(), "page": 1, "start": 0, "length": 25}
        try:
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            if not data.get("success") or not data.get("data"):
                return None, f"Мероприятие '{activity_name}' не найдено"
            acts = data["data"]
            for a in acts:
                if a.get("name", "").strip().lower() == activity_name.strip().lower():
                    return a, "Найдено точное совпадение"
            return acts[0], f"Найдено: '{acts[0].get('name')}'"
        except Exception as e:
            return None, f"Сбой поиска мероприятия: {e}"

    def find_kid(self, fio: str, birth_date: Optional[Any] = None) -> Tuple[Optional[Dict[str, Any]], str]:
        clean_fio = " ".join(str(fio).strip().split())
        url = f"{self.base_url}/api/activity/rest/kid"
        params = {"search[value]": clean_fio, "page": 1, "start": 0, "length": 25}
        try:
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            if not data.get("success") or not data.get("data"):
                return None, f"Ребенок не найден в базе"

            candidates: List[Dict[str, Any]] = data["data"]
            target_dob = self._normalize_dob(birth_date)

            # Фильтрация по совпадению ДР
            matched_by_dob = []
            if target_dob:
                for c in candidates:
                    if self._normalize_dob(c.get("birthday")) == target_dob:
                        matched_by_dob.append(c)
                if not matched_by_dob:
                    found_dobs = [str(c.get("birthday")) for c in candidates if c.get("birthday")]
                    return None, f"Найдено {len(candidates)} тезок, но ДР '{target_dob}' не совпала ({', '.join(found_dobs)})"
            else:
                matched_by_dob = candidates

            # Фильтрация по обязательному подтвержденному аккаунту 'is_approved' == True
            approved = [
                c for c in matched_by_dob
                if c.get("is_approved") is True or str(c.get("is_approved")).lower() in ["true", "1"]
            ]
            if not approved:
                return None, f"Ребенок найден, но аккаунт НЕ подтвержден ('is_approved': false)"

            chosen = approved[0]
            return chosen, f"OK (аккаунт подтвержден, ДР {chosen.get('birthday')})"
        except Exception as e:
            return None, f"Сбой поиска ребенка: {e}"

    def create_order(self, activity_id: int, event_datetime: str, kid: Dict[str, Any]) -> Tuple[bool, str, Optional[int]]:
        url = f"{self.base_url}/api/rest/activityOrder"
        norm_dt = self._normalize_datetime(event_datetime)
        payload = {
            "data": {
                "activity_id": int(activity_id),
                "date": norm_dt,
                "site_user_id": int(kid["site_user_id"]),
                "kid_id": kid["id"],
                "state": "initial"
            }
        }
        try:
            resp = self._request("POST", url, json=payload, timeout=15)
            data = resp.json()
            if data.get("success"):
                oid = data.get("data", {}).get("id")
                return True, f"Заявка создана (#{oid})", oid
            return False, str(data.get("message") or "Отказ сервера"), None
        except Exception as e:
            return False, f"Ошибка создания заявки: {e}", None

    def create_program_order(
        self,
        event_id: Any,
        group_id: Any,
        academic_year_id: Any,
        kid: Dict[str, Any],
        use_certificate: bool = False,
        create_certificate: bool = True
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Создает заявку на зачисление на учебную программу (POST /api/rest/order).
        Payload:
        {
          "data": {
            "use_certificate": false,
            "create_certificate": true,
            "site_user_id": "791100",
            "event_id": "12345",
            "group_id": "67890",
            "academic_year_id": "2026",
            "kid_id": "f6c7c6d2-6b40-4767-a77c-5d0889139026"
          }
        }
        """
        url = f"{self.base_url}/api/rest/order"
        payload = {
            "data": {
                "use_certificate": bool(use_certificate),
                "create_certificate": bool(create_certificate),
                "site_user_id": str(kid.get("site_user_id", "")).strip(),
                "event_id": str(event_id).strip(),
                "group_id": str(group_id).strip(),
                "academic_year_id": str(academic_year_id).strip(),
                "kid_id": str(kid.get("id", "")).strip()
            }
        }
        try:
            resp = self._request("POST", url, json=payload, timeout=20)
            data = resp.json()
            if data.get("success") or data.get("err_code") == 0:
                order_data = data.get("data", {})
                oid = order_data.get("id")
                return True, f"Заявка на зачисление создана (#{oid})", oid
            msg = data.get("message") or data.get("errors") or data.get("warnings") or "Отказ сервера при создании заявки"
            return False, str(msg), None
        except Exception as e:
            return False, f"Ошибка создания заявки на программу: {e}", None

    def check_existing_program_order(self, kid_id: str, event_id: Union[int, str]) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """
        Проверяет в Навигаторе, не была ли уже подана заявка на зачисление для данного ребенка на эту программу.
        Использует GET /api/rest/order с фильтрацией:
          - state_grid in ["initial", "approve", "study"]
          - kid_id eq <kid_id>
          - event_id eq <event_id>
        Возвращает:
          (True, order_data, f"Заявка #{order_id} уже существует в статусе '{state}'") если найдена,
          (False, None, "Активных заявок не найдено") если не найдена.
        """
        clean_kid_id = str(kid_id).strip()
        clean_event_id = str(event_id).strip()
        if not clean_kid_id or not clean_event_id:
            return False, None, "Не указан kid_id или event_id для проверки существующих заявок"

        url = f"{self.base_url}/api/rest/order"
        filters = [
            {"property": "state_grid", "value": ["initial", "approve", "study"], "comparison": "in"},
            {"property": "kid_id", "value": clean_kid_id, "comparison": "eq"},
            {"property": "event_id", "value": clean_event_id, "comparison": "eq"}
        ]
        params = {
            "start": 0,
            "page": 1,
            "length": 25,
            "pagination": "arrows",
            "extFilters": json.dumps(filters, ensure_ascii=False)
        }
        try:
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            if (data.get("success") or data.get("err_code") == 0) and data.get("data"):
                items = data.get("data", [])
                if items and isinstance(items, list) and len(items) > 0:
                    order = items[0]
                    oid = order.get("id")
                    st = order.get("state", order.get("state_grid", "активна"))
                    return True, order, f"Заявка #{oid} уже существует в статусе '{st}'"
            return False, None, "Активных заявок не найдено"
        except Exception as e:
            return False, None, f"Ошибка проверки существующих заявок: {e}"

    def search_program(self, program_name: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Ищет учебную программу (курс/мероприятие программы) по названию через GET /api/rest/events.
        Фильтры: is_deleted == false, name like <program_name>.
        """
        clean_name = program_name.strip()
        if not clean_name:
            return None, "Название программы не указано"

        url = f"{self.base_url}/api/rest/events"

        def _do_query(search_val: str) -> List[Dict[str, Any]]:
            filters = [
                {"type": "boolean", "property": "is_deleted", "value": False, "comparison": "eq"},
                {"property": "name", "value": search_val, "comparison": "like"}
            ]
            params = {
                "_dc": int(time.time() * 1000),
                "page": 1,
                "start": 0,
                "length": 25,
                "extFilters": json.dumps(filters, ensure_ascii=False)
            }
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            if data.get("success") and data.get("data"):
                return data.get("data", [])
            return []

        try:
            items = _do_query(clean_name)

            # Если с исходным названием ничего не найдено, пробуем без кавычек
            if not items:
                unquoted = clean_name.replace('"', '').replace('«', '').replace('»', '').strip()
                if unquoted and unquoted != clean_name:
                    items = _do_query(unquoted)

            if not items:
                return None, f"Программа '{clean_name}' не найдена"

            def _norm(s: str) -> str:
                return " ".join(s.lower().replace('"', '').replace('«', '').replace('»', '').split())

            target_norm = _norm(clean_name)

            # 1. Точное совпадение
            for it in items:
                if _norm(it.get("name", "")) == target_norm:
                    return it, f"Найдено точное совпадение: '{it.get('name')}' (ID: {it.get('id')})"

            # 2. Подстрока
            for it in items:
                it_norm = _norm(it.get("name", ""))
                if target_norm in it_norm or it_norm in target_norm:
                    return it, f"Найдено совпадение: '{it.get('name')}' (ID: {it.get('id')})"

            # 3. Первое в списке
            first = items[0]
            return first, f"Найдено: '{first.get('name')}' (ID: {first.get('id')})"
        except Exception as e:
            return None, f"Сбой поиска программы: {e}"

    def get_program_groups(self, event_id: Union[int, str]) -> Tuple[List[Dict[str, Any]], str]:
        """
        Получает список групп, прикрепленных к указанной программе (event_id),
        через GET /api/rest/eventGroups?extFilters=[{"property":"event_id","value":<event_id>},{"property":"is_deleted","value":"0","comparison":"eq"}]
        """
        try:
            url = f"{self.base_url}/api/rest/eventGroups"
            ev_val = int(event_id) if str(event_id).isdigit() else event_id
            ext_filters = [
                {"property": "event_id", "value": ev_val},
                {"property": "is_deleted", "value": "0", "comparison": "eq"}
            ]
            params = {
                "_dc": int(time.time() * 1000),
                "page": 1,
                "start": 0,
                "length": 100,
                "extFilters": json.dumps(ext_filters)
            }
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            if data.get("success") or data.get("err_code") == 0:
                raw_groups = data.get("data", [])
                active_groups = [g for g in raw_groups if not g.get("is_deleted")]
                return active_groups, ""
            return [], data.get("message", "Не удалось загрузить группы программы")
        except Exception as e:
            return [], f"Ошибка загрузки групп программы: {e}"

    def validate_group_for_program(
        self,
        group_input: str,
        event_id: Union[int, str],
        program_name: str = ""
    ) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]], str]:
        """
        Проверяет, что группа принадлежит указанной программе (event_id).
        1. Запрашивает группы программы через GET /api/rest/eventGroups с extFilters=[{"property":"event_id","value":...}].
        2. Если group_input - числовой ID:
           Проверяет точное совпадение id в списке групп программы.
           Если нет - возвращает ошибку о несовпадении программы и список доступных групп этой программы.
        3. Если group_input - текстовое название:
           Ищет группу среди групп этой программы.
           - Если найдена ровно 1 группа с таким названием в программе: возвращает её.
           - Если несколько групп с таким названием в этой же программе: возвращает предупреждение со списком ID.
           - Если в этой программе такой группы нет: возвращает ошибку и список доступных групп.
        """
        clean_input = group_input.strip()
        if not clean_input:
            return None, [], "Группа не указана"

        prog_groups, err_msg = self.get_program_groups(event_id)
        if err_msg:
            return None, [], f"Не удалось получить группы программы (ID: {event_id}): {err_msg}"

        p_label = f"'{program_name}' (ID: {event_id})" if program_name else f"ID {event_id}"

        if not prog_groups:
            return None, [], f"У программы {p_label} нет активных групп в Навигаторе."

        def _format_available(groups: List[Dict[str, Any]]) -> str:
            lines = []
            for g in groups:
                lines.append(f"• ID {g.get('id')}: «{g.get('name')}» (педагог: {g.get('teacher', 'н/д')}, мест: {g.get('size', 'н/д')})")
            return "\n".join(lines)

        # 1. Если введен точный числовой ID группы
        if clean_input.isdigit():
            for g in prog_groups:
                if str(g.get("id")) == clean_input:
                    return g, prog_groups, f"Группа ID {clean_input} подтверждена: «{g.get('name')}» (принадлежит программе {p_label})"

            avail_str = _format_available(prog_groups)
            other_info = ""
            try:
                single_g, _, _ = self.search_group(clean_input)
                if single_g and single_g.get("event_id"):
                    other_info = f"\n(Внимание: группа с ID {clean_input} «{single_g.get('name')}» привязана к другой программе: ID {single_g.get('event_id')})\n"
            except Exception:
                pass

            msg = (
                f"Группа с ID {clean_input} НЕ принадлежит программе {p_label}!{other_info}\n"
                f"Группы, доступные для этой программы ({len(prog_groups)} шт.):\n"
                f"{avail_str}\n\n"
                f"Пожалуйста, выберите ID группы из списка выше."
            )
            return None, prog_groups, msg

        # 2. Если введено текстовое название группы
        def _norm(s: str) -> str:
            return " ".join(s.lower().replace('"', '').replace('«', '').replace('»', '').split())

        target_norm = _norm(clean_input)
        exact_matches = [g for g in prog_groups if _norm(g.get("name", "")) == target_norm]
        partial_matches = [g for g in prog_groups if target_norm in _norm(g.get("name", ""))] if not exact_matches else []
        matched = exact_matches if exact_matches else partial_matches

        if len(matched) == 1:
            g = matched[0]
            return g, prog_groups, f"Группа «{g.get('name')}» (ID: {g.get('id')}) найдена и принадлежит программе {p_label}"

        if len(matched) > 1:
            details = _format_available(matched)
            msg = (
                f"В программе {p_label} найдено несколько групп с названием '{clean_input}' ({len(matched)} шт.):\n"
                f"{details}\n\n"
                f"Пожалуйста, укажите точный ID нужной группы в поле «Группа»."
            )
            return None, matched, msg

        # Группа с таким названием не найдена в этой программе
        avail_str = _format_available(prog_groups)
        msg = (
            f"Группа с названием '{clean_input}' НЕ найдена в программе {p_label}!\n\n"
            f"Группы, доступные для этой программы ({len(prog_groups)} шт.):\n"
            f"{avail_str}\n\n"
            f"Пожалуйста, укажите точное название или ID группы из списка выше."
        )
        return None, prog_groups, msg

    def search_group(self, group_query: str) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]], str]:
        """
        Ищет группу в Навигаторе по общему REST API GET /api/rest/eventGroups?query=...
        Используется для глобального поиска или разрешения информации об отдельной группе.
        """
        clean_query = group_query.strip()
        if not clean_query:
            return None, [], "Название или ID группы не указаны"

        url = f"{self.base_url}/api/rest/eventGroups"
        if clean_query.isdigit():
            try:
                params = {
                    "_dc": int(time.time() * 1000),
                    "page": 1,
                    "start": 0,
                    "length": 10,
                    "query": clean_query
                }
                resp = self._request("GET", url, params=params, timeout=15)
                data = resp.json()
                items = data.get("data", [])
                for it in items:
                    if str(it.get("id")) == clean_query:
                        return it, [it], f"Группа ID {clean_query}: '{it.get('name')}' (педагог: {it.get('teacher', 'н/д')})"
            except Exception:
                pass
            stub = {"id": clean_query, "name": f"Группа #{clean_query}"}
            return stub, [stub], f"ID группы: {clean_query}"

        try:
            params = {
                "_dc": int(time.time() * 1000),
                "page": 1,
                "start": 0,
                "length": 50,
                "sort": json.dumps([{"property": "id", "direction": "DESC"}]),
                "query": clean_query
            }
            resp = self._request("GET", url, params=params, timeout=15)
            data = resp.json()
            raw_items = data.get("data", []) if (data.get("success") or data.get("err_code") == 0) else []
            active_items = [g for g in raw_items if not g.get("is_deleted")]
            if not active_items:
                return None, [], f"Группа '{clean_query}' не найдена в Навигаторе"

            def _norm(s: str) -> str:
                return " ".join(s.lower().replace('"', '').replace('«', '').replace('»', '').split())

            target_norm = _norm(clean_query)
            exact_matches = [g for g in active_items if _norm(g.get("name", "")) == target_norm]
            matches = exact_matches if exact_matches else active_items

            if len(matches) == 1:
                g = matches[0]
                return g, matches, f"Найдена группа: '{g.get('name')}' (ID: {g.get('id')})"
            return None, matches, f"Найдено {len(matches)} групп с названием '{clean_query}'"
        except Exception as e:
            return None, [], f"Ошибка поиска группы: {e}"

    def get_activity_orders_by_state(self, activity_id: int, states: List[str], limit: int = 100) -> Tuple[List[Dict[str, Any]], int]:
        """Возвращает заявки мероприятия с фильтрацией по статусам (state)."""
        url = f"{self.base_url}/api/rest/activityOrder"
        filters = [
            {"property": "activity_id", "value": str(activity_id), "comparison": "eq"},
            {"property": "state", "value": states, "comparison": "in"}
        ]
        params = {"page": 1, "start": 0, "length": limit, "extFilters": json.dumps(filters)}
        try:
            resp = self._request("GET", url, params=params, timeout=20)
            data = resp.json()
            if data.get("success"):
                return data.get("data", []), int(data.get("recordsFiltered", 0))
            return [], 0
        except Exception:
            return [], 0

    def get_pending_orders(self, activity_id: int, limit: int = 100) -> Tuple[List[Dict[str, Any]], int]:
        """Возвращает неподтвержденные заявки (state: initial)."""
        return self.get_activity_orders_by_state(activity_id, ["initial"], limit)

    def get_approved_orders(self, activity_id: int, limit: int = 100) -> Tuple[List[Dict[str, Any]], int]:
        """Возвращает подтвержденные заявки, ожидающие отметки участия (state: approve)."""
        return self.get_activity_orders_by_state(activity_id, ["approve"], limit)

    def set_activity_order_state(self, order_id: int, state: str, comment: str = "") -> Tuple[bool, str]:
        """Устанавливает статус заявки (approve, participant и др.)."""
        url = f"{self.base_url}/api/setActivityOrderState"
        data_body: Dict[str, Any] = {"id": int(order_id), "state": state}
        if state == "approve" and comment is not None:
            data_body["approve_comment"] = comment
        payload = {"data": data_body}
        try:
            resp = self._request("POST", url, json=payload, timeout=10)
            data = resp.json()
            if data.get("success"):
                return True, f"Заявка #{order_id} переведена в статус '{state}'"
            return False, data.get("message", "Отказ")
        except Exception as e:
            return False, f"Ошибка: {e}"

    def approve_order(self, order_id: int) -> Tuple[bool, str]:
        """Подтверждает заявку (state -> approve)."""
        return self.set_activity_order_state(order_id, "approve", comment="")

    def mark_participant_order(self, order_id: int) -> Tuple[bool, str]:
        """Отмечает участие ребенка по заявке (state -> participant)."""
        return self.set_activity_order_state(order_id, "participant")

    def get_program_orders(
        self,
        event_id: Union[int, str],
        academic_year_id: str = "2026",
        group_id: Optional[Union[int, str]] = None,
        states: Optional[List[str]] = None,
        page: int = 1,
        start: int = 0,
        length: int = 50
    ) -> Tuple[List[Dict[str, Any]], bool, int, str]:
        """
        Запрашивает заявки на учебную программу через:
        GET /api/rest/order?start=0&page=1&length=50&pagination=arrows&extFilters=[...]
        Фильтры:
          - fact_academic_year_id = <academic_year_id>
          - event_id = <event_id>
          - (опционально) fact_group_id = <group_id> (если указан)
          - (опционально) state_grid in <states> (если указан)
        Возвращает: (список_заявок, nextPage_bool, recordsFiltered, error_message)
        """
        url = f"{self.base_url}/api/rest/order"
        filters: List[Dict[str, Any]] = [
            {"property": "fact_academic_year_id", "value": str(academic_year_id).strip(), "comparison": "eq"},
            {"property": "event_id", "value": str(event_id).strip(), "comparison": "eq"}
        ]
        if group_id is not None and str(group_id).strip():
            filters.append({"property": "fact_group_id", "value": str(group_id).strip(), "comparison": "eq"})
        if states:
            filters.append({"property": "state_grid", "value": states, "comparison": "in"})

        params = {
            "start": start,
            "page": page,
            "length": length,
            "pagination": "arrows",
            "extFilters": json.dumps(filters, ensure_ascii=False)
        }
        try:
            resp = self._request("GET", url, params=params, timeout=20)
            data = resp.json()
            if data.get("success") or data.get("err_code") == 0:
                orders = data.get("data", [])
                next_page = bool(data.get("nextPage", False))
                total = int(data.get("recordsFiltered", len(orders)))
                return orders, next_page, total, ""
            return [], False, 0, str(data.get("message", "Не удалось получить список заявок программы"))
        except Exception as e:
            return [], False, 0, f"Ошибка запроса заявок программы: {e}"

    def get_all_program_orders(
        self,
        event_id: Union[int, str],
        academic_year_id: str = "2026",
        group_id: Optional[Union[int, str]] = None,
        states: Optional[List[str]] = None,
        max_records: int = 500
    ) -> Tuple[List[Dict[str, Any]], str]:
        """
        Запрашивает ВСЕ страницы заявок на программу с авто-пагинацией (по 50 штук на страницу).
        """
        all_orders: List[Dict[str, Any]] = []
        page = 1
        page_size = 50
        while True:
            start = (page - 1) * page_size
            orders, next_page, total, err = self.get_program_orders(
                event_id=event_id,
                academic_year_id=academic_year_id,
                group_id=group_id,
                states=states,
                page=page,
                start=start,
                length=page_size
            )
            if err:
                return all_orders, err
            if not orders:
                break
            all_orders.extend(orders)
            if not next_page or len(all_orders) >= max_records or len(orders) < page_size:
                break
            page += 1
        return all_orders, ""

    def approve_program_order(self, order_id: Union[int, str], comment: str = "") -> Tuple[bool, str]:
        """
        Подтверждает заявку на программу через POST /api/approveRequest.
        Заявка должна быть в статусе initial.
        curl: POST /api/approveRequest
        payload: {"data":{"id":"<order_id>","comment":"","entry_exams":false,"send_to_rpgu":true}}
        """
        url = f"{self.base_url}/api/approveRequest"
        payload = {
            "data": {
                "id": str(order_id).strip(),
                "comment": comment,
                "entry_exams": False,
                "send_to_rpgu": True
            }
        }
        try:
            resp = self._request("POST", url, json=payload, timeout=15)
            data = resp.json()
            if data.get("success") or data.get("err_code") == 0:
                return True, f"Заявка #{order_id} подтверждена (state -> approve)"
            msg = data.get("message") or data.get("errors") or data.get("warnings") or "Отказ сервера"
            return False, str(msg)
        except Exception as e:
            return False, f"Ошибка подтверждения заявки #{order_id}: {e}"

    def study_program_order(
        self,
        order_id: Union[int, str],
        financing_source: str = "1",
        date_start: str = "2026-09-01",
        decree_number: str = "183",
        date_signing: str = "2026-08-31",
        comment: str = ""
    ) -> Tuple[bool, str]:
        """
        Отмечает обучение ребенка по программе через POST /api/studyRequest.
        Заявка должна быть в статусе approve.
        curl: POST /api/studyRequest
        headers: X-Skip-400-Error-Window: 1
        payload: {"data":{"id":"<order_id>","financing_source":"1","date_start":"2026-09-01","decree_number":"183","date_signing":"2026-08-31","comment":"","reasons":[]}}
        """
        url = f"{self.base_url}/api/studyRequest"
        payload = {
            "data": {
                "id": str(order_id).strip(),
                "financing_source": str(financing_source).strip() or "1",
                "date_start": date_start.strip(),
                "decree_number": str(decree_number).strip(),
                "date_signing": date_signing.strip(),
                "comment": comment,
                "reasons": []
            }
        }
        headers = {"X-Skip-400-Error-Window": "1"}
        try:
            resp = self._request("POST", url, json=payload, headers=headers, timeout=15)
            data = resp.json()
            if data.get("success") or data.get("err_code") == 0:
                return True, f"Заявка #{order_id} зачислена на обучение (state -> study)"
            msg = data.get("message") or data.get("errors") or data.get("warnings") or "Отказ сервера"
            return False, str(msg)
        except Exception as e:
            return False, f"Ошибка зачисления на обучение заявки #{order_id}: {e}"

    def _normalize_dob(self, raw: Optional[Any]) -> Optional[str]:
        if raw is None:
            return None
        if hasattr(raw, "strftime"):
            return raw.strftime("%Y-%m-%d")
        s = str(raw).strip()
        if not s or s.lower() in ["none", "nan", "nat", ""]:
            return None
        if " " in s:
            s = s.split(" ")[0]
        s_clean = s.replace("/", ".").replace("-", ".")
        parts = [p.strip() for p in s_clean.split(".") if p.strip()]
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
            day = parts[0].zfill(2)
            month = parts[1].zfill(2)
            y = parts[2]
            if len(y) == 2:
                y_num = int(y)
                full_y = 2000 + y_num if y_num <= 35 else 1900 + y_num
                return f"{full_y}-{month}-{day}"
            elif len(y) == 4:
                return f"{y}-{month}-{day}"
        return None

    def _normalize_datetime(self, dt_str: str) -> str:
        s = dt_str.strip()
        if "." in s and " " in s:
            d_part, t_part = s.split(" ", 1)
            parts = d_part.split(".")
            t_clean = t_part if len(t_part.split(":")) == 3 else f"{t_part}:00"
            return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)} {t_clean}"
        if len(s.split(":")) == 2:
            return f"{s}:00"
        return s

    def update_child_status(self, file_path: str, row_num: int, status_text: str, col_num: int = 3) -> Tuple[bool, str]:
        """Обновляет статус/флаг обработки ребенка в указанной ячейке Excel."""
        return update_excel_cell_status(file_path, row_num, status_text, col_num)

    def read_excel(self, file_path: str) -> List[Dict[str, Any]]:
        try:
            import openpyxl
        except ImportError:
            raise ImportError("Необходимо установить openpyxl: pip install openpyxl")

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        fio_col = 1
        dob_col = 2
        status_col = 3

        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(1, col).value or "").strip().lower()
            if "фио" in val or "ребенок" in val or "фамилия" in val:
                fio_col = col
            elif "рожд" in val or "дат" in val or "др" in val:
                dob_col = col
            elif "статус" in val or "обработ" in val or "флаг" in val or "заявк" in val:
                status_col = col

        # Маркеры успешной обработки заявки (включая запись на мероприятие и зачисление на программу)
        processed_keywords = ["добавлен", "обработан", "успешно", "заявка", "зачислен", "да", "ок", "ok", "+", "true", "1"]

        dirty = False
        records = []
        for r in range(2, ws.max_row + 1):
            fio = ws.cell(r, fio_col).value
            dob = ws.cell(r, dob_col).value
            raw_status = ws.cell(r, status_col).value if ws.max_column >= status_col else None
            status_str = str(raw_status).strip() if raw_status is not None else ""

            if not fio or not str(fio).strip():
                continue
            fio_str = str(fio).strip()

            # Проверка формата ФИО (защита от инициалов вроде 'Иванов И.И.')
            is_valid_fmt, fmt_err = check_fio_format(fio_str)
            if not is_valid_fmt:
                err_marker = "⚠️ Неправильный формат (инициалы: укажите Фамилию и полное Имя)" if "инициал" in fmt_err else f"⚠️ {fmt_err}"
                if not status_str or ("неправильный формат" not in status_str.lower() and "инициал" not in status_str.lower()):
                    status_str = err_marker
                    ws.cell(r, status_col, value=status_str)
                    dirty = True

            # Если в статусе уже зафиксировано успешное добавление
            is_processed = bool(
                status_str and any(kw in status_str.lower() for kw in processed_keywords)
            )

            records.append({
                "row": r,
                "fio": fio_str,
                "dob": dob,
                "status": status_str,
                "is_processed": is_processed,
                "is_valid_format": is_valid_fmt,
                "format_error": fmt_err if not is_valid_fmt else "",
                "status_col": status_col,
            })

        if dirty:
            try:
                wb.save(file_path)
            except Exception as e:
                print(f"Ошибка сохранения отметки формата в {file_path}: {e}")

        wb.close()
        return records

    def read_study_excel(self, file_path: str = STUDY_EXCEL_FILE) -> List[Dict[str, Any]]:
        """
        Считывает отдельную 2-колоночную таблицу (study_list.xlsx).
        Столбец 1: ФИО ребенка
        Столбец 2: Флаг зачисления
        Возвращает список записей с информацией о строке, ФИО, текущем флаге и маркерах подтверждения/зачисления.
        """
        try:
            import openpyxl
        except ImportError:
            raise ImportError("Необходимо установить openpyxl: pip install openpyxl")

        if not os.path.exists(file_path):
            ensure_study_excel(file_path)

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        fio_col = 1
        flag_col = 2

        for col in range(1, min(ws.max_column + 1, 10)):
            val = str(ws.cell(1, col).value or "").strip().lower()
            if "фио" in val or "ребенок" in val or "фамилия" in val or "ученик" in val:
                fio_col = col
            elif "флаг" in val or "статус" in val or "зачисл" in val or "подтвержд" in val:
                flag_col = col

        records = []
        dirty = False
        for r in range(2, ws.max_row + 1):
            fio_val = ws.cell(r, fio_col).value
            if not fio_val or not str(fio_val).strip():
                continue
            fio_str = str(fio_val).strip()
            raw_flag = ws.cell(r, flag_col).value if ws.max_column >= flag_col else None
            flag_str = str(raw_flag).strip() if raw_flag is not None else ""

            # Проверка формата ФИО (защита от инициалов вроде 'Иванов И.И.')
            is_valid_fmt, fmt_err = check_fio_format(fio_str)
            if not is_valid_fmt:
                # Если в ячейке флага еще не зафиксировано предупреждение о неправильном формате,
                # отмечаем в столбце флага Excel как неправильный формат
                err_marker = "⚠️ Неправильный формат (инициалы: укажите Фамилию и полное Имя)" if "инициал" in fmt_err else f"⚠️ {fmt_err}"
                if not flag_str or ("неправильный формат" not in flag_str.lower() and "инициал" not in flag_str.lower()):
                    flag_str = err_marker
                    ws.cell(r, flag_col, value=flag_str)
                    dirty = True

            flag_lower = flag_str.lower()
            is_confirmed = any(kw in flag_lower for kw in ["подтвержд", "approve", "подтвержден"])
            is_enrolled = any(kw in flag_lower for kw in ["зачислен", "study", "приказ", "обучение"])

            records.append({
                "row": r,
                "fio": fio_str,
                "flag": flag_str,
                "is_valid_format": is_valid_fmt,
                "format_error": fmt_err if not is_valid_fmt else "",
                "is_confirmed": is_confirmed,
                "is_enrolled": is_enrolled,
                "flag_col": flag_col,
            })

        if dirty:
            try:
                wb.save(file_path)
            except Exception as e:
                print(f"Ошибка сохранения отметки формата в {file_path}: {e}")

        wb.close()
        return records

    def update_study_child_flag(self, file_path: str, row_num: int, flag_text: str, col_num: int = 2) -> Tuple[bool, str]:
        """Обновляет флаг подтверждения/зачисления ребенка в таблице study_list.xlsx."""
        return update_excel_cell_status(file_path, row_num, flag_text, col_num)


# =========================================================================
# 3. ГРАФИЧЕСКИЙ ИНТЕРФЕЙС (TKINTER GUI)
# =========================================================================
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkfont


def enable_high_dpi():
    """
    Включает аппаратную поддержку High-DPI в Windows (100%, 125%, 150%, 200%, 4K).
    Устраняет системное мыльное сглаживание Windows, делая текст и элементы кристально четкими.
    """
    if sys.platform.startswith("win"):
        try:
            import ctypes
            # Per-Monitor DPI Awareness v2 (Windows 10 1703+ и Windows 11)
            # Значение 2 = DPI_AWARENESS_PER_MONITOR_AWARE_V2
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                # Per-Monitor DPI (Windows 8.1+)
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                try:
                    # Fallback для старых сборок Windows
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception:
                    pass


def setup_universal_clipboard(root: tk.Tk):
    """
    Глобально включает работу Ctrl+C, Ctrl+V, Ctrl+X, Ctrl+A для ВСЕХ инпутов (Entry, Text),
    включая русскую раскладку клавиатуры (где Tkinter по умолчанию блокирует хоткеи).
    Также добавляет контекстное меню правой кнопки мыши (ПКМ).
    """

    def get_focused_entry(event=None):
        w = root.focus_get()
        if isinstance(w, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text)):
            return w
        if event and hasattr(event, "widget") and isinstance(event.widget, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text)):
            return event.widget
        return None

    def copy_text(w=None):
        widget = w or get_focused_entry()
        if not widget:
            return "break"
        try:
            if isinstance(widget, tk.Text):
                if widget.tag_ranges("sel"):
                    selected = widget.get("sel.first", "sel.last")
                    root.clipboard_clear()
                    root.clipboard_append(selected)
            else:
                if widget.selection_present():
                    selected = widget.selection_get()
                    root.clipboard_clear()
                    root.clipboard_append(selected)
        except Exception:
            pass
        return "break"

    def cut_text(w=None):
        widget = w or get_focused_entry()
        if not widget:
            return "break"
        try:
            copy_text(widget)
            if isinstance(widget, tk.Text):
                if widget.tag_ranges("sel"):
                    widget.delete("sel.first", "sel.last")
            else:
                if widget.selection_present():
                    first = widget.index("sel.first")
                    last = widget.index("sel.last")
                    widget.delete(first, last)
        except Exception:
            pass
        return "break"

    def paste_text(w=None):
        widget = w or get_focused_entry()
        if not widget:
            return "break"
        try:
            clipboard = root.clipboard_get()
            if isinstance(widget, tk.Text):
                if widget.tag_ranges("sel"):
                    widget.delete("sel.first", "sel.last")
                widget.insert("insert", clipboard)
            else:
                if widget.selection_present():
                    first = widget.index("sel.first")
                    last = widget.index("sel.last")
                    widget.delete(first, last)
                widget.insert("insert", clipboard)
        except Exception:
            pass
        return "break"

    def select_all(w=None):
        widget = w or get_focused_entry()
        if not widget:
            return "break"
        try:
            if isinstance(widget, tk.Text):
                widget.tag_add("sel", "1.0", "end")
                widget.mark_set("insert", "end")
            else:
                widget.selection_range(0, tk.END)
                widget.icursor(tk.END)
        except Exception:
            pass
        return "break"

    # Создаем контекстное меню ПКМ
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="Вставить (Ctrl+V)", command=paste_text)
    menu.add_command(label="Копировать (Ctrl+C)", command=copy_text)
    menu.add_command(label="Вырезать (Ctrl+X)", command=cut_text)
    menu.add_separator()
    menu.add_command(label="Выделить всё (Ctrl+A)", command=select_all)

    def show_context_menu(event):
        widget = event.widget
        if isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text)):
            widget.focus_set()
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    root.bind_all("<Button-3>", show_context_menu)
    # На Mac правая кнопка мыши или Control+Click
    root.bind_all("<Button-2>", show_context_menu)

    # Обработчик нажатия клавиш с проверкой Control / Command
    def on_global_key(event):
        widget = event.widget
        if not isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text)):
            return

        # Маска Control: 0x4 (Windows/Linux) или Command (Mac)
        ctrl_pressed = bool(event.state & 0x0004 or event.state & 0x0008)
        if not ctrl_pressed:
            return

        keysym = event.keysym.lower()
        char = event.char.lower() if event.char else ""

        # Ctrl + V (русская раскладка: клавиша 'м' / cyrillic_em)
        if keysym in ["v", "cyrillic_em"] or char in ["v", "м", "\x16"]:
            paste_text(widget)
            return "break"

        # Ctrl + C (русская раскладка: клавиша 'с' / cyrillic_es)
        if keysym in ["c", "cyrillic_es"] or char in ["c", "с", "\x03"]:
            copy_text(widget)
            return "break"

        # Ctrl + X (русская раскладка: клавиша 'ч' / cyrillic_che)
        if keysym in ["x", "cyrillic_che"] or char in ["x", "ч", "\x18"]:
            cut_text(widget)
            return "break"

        # Ctrl + A (русская раскладка: клавиша 'ф' / cyrillic_ef)
        if keysym in ["a", "cyrillic_ef"] or char in ["a", "ф", "\x01"]:
            select_all(widget)
            return "break"

    # Привязываем ко всем событиям нажатия клавиш
    root.bind_all("<KeyPress>", on_global_key, add="+")


class NavigatorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Навигатор: Автоматизация v{APP_VERSION} (Мероприятия, Подтверждение & Зачисление)")
        self.root.geometry("880x730")
        self.root.minsize(820, 640)

        # Подключаем глобальную поддержку буфера обмена (Ctrl+C, Ctrl+V в RU/EN + ПКМ)
        setup_universal_clipboard(self.root)

        self.client = NavigatorClient()
        self.saved_cfg = load_config()

        # Флаги выполнения фоновых задач и аварийной остановки
        self.stop_requested = False
        self.is_task_running = False
        self.current_task_name = ""

        # Привязываем проверку прерывания к REST-клиенту и регулятору частоты
        self.client.abort_checker = lambda: getattr(self, "stop_requested", False)
        self.client.limiter.abort_checker = lambda: getattr(self, "stop_requested", False)

        # Настраиваем мониторинг отклика сервера в реальном времени
        self.is_monitoring_active = True
        self.client.on_ping_update = self._on_server_ping_update
        self.client.on_log_message = self.log

        # Создаем шаблонные таблицы (event_list.xlsx и programm_list.xlsx) и файл журнала при первом запуске
        ensure_all_excel_files()
        ensure_log_file()

        # Стили интерфейса
        self._setup_styles()

        # Запускаем фоновый цикл периодического замера отклика
        self._start_bg_ping_loop()

        # Стартовое состояние: окно авторизации
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.show_login_screen()

        # Фоновая проверка обновлений на GitHub через 1.5 сек после старта программы
        if self.saved_cfg.get("check_updates", True):
            self.root.after(1500, lambda: self.check_updates_async(silent_if_latest=True))

    def _on_server_ping_update(self, latency_ms: float, load_label: str, color: str, delay: float):
        """Обновляет статус отклика сервера и скорости в GUI в реальном времени."""
        def update_ui():
            try:
                if hasattr(self, "lbl_server_ping") and self.lbl_server_ping.winfo_exists():
                    lat_str = f"{latency_ms:.0f} мс" if latency_ms < 1000 else f"{latency_ms/1000:.2f} сек"
                    self.lbl_server_ping.config(
                        text=f"● Отклик сервера: {lat_str} | Нагрузка: {load_label}",
                        foreground=color
                    )
                if hasattr(self, "lbl_server_speed") and self.lbl_server_speed.winfo_exists():
                    rps = round(1.0 / max(0.1, delay), 2)
                    self.lbl_server_speed.config(
                        text=f"Скорость отправки: {rps} запр/сек (пауза {delay:.1f}с) | Режим: Защита от перегрузки (строго ≤ 1 запр/сек)"
                    )
            except Exception:
                pass
        self.root.after(0, update_ui)

    def manual_check_ping(self):
        """Ручная проверка отклика сервера по нажатию кнопки."""
        if hasattr(self, "btn_ping_now") and self.btn_ping_now.winfo_exists():
            self.btn_ping_now.config(state=tk.DISABLED, text="⏳ Замер...")

        def worker():
            self.client.check_ping()
            def finish():
                if hasattr(self, "btn_ping_now") and self.btn_ping_now.winfo_exists():
                    self.btn_ping_now.config(state=tk.NORMAL, text="🔄 Проверить отклик")
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _start_bg_ping_loop(self):
        """Фоновый цикл проверки отклика сервера каждые 20 секунд."""
        def loop():
            time.sleep(2)
            while getattr(self, "is_monitoring_active", True):
                if hasattr(self, "lbl_server_ping") and self.lbl_server_ping.winfo_exists():
                    try:
                        self.client.check_ping()
                    except Exception:
                        pass
                time.sleep(20)
        threading.Thread(target=loop, daemon=True).start()

    def _setup_styles(self):
        style = ttk.Style()
        available = style.theme_names()
        # Для кристальной четкости в Windows используем нативную тему vista или winnative
        if sys.platform.startswith("win"):
            if "vista" in available:
                style.theme_use("vista")
            elif "winnative" in available:
                style.theme_use("winnative")
        elif "clam" in available:
            style.theme_use("clam")

        # Настраиваем чистый, четкий системный шрифт (Segoe UI в Windows, SF Pro/Helvetica на Mac)
        font_family = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"
        for font_name in ["TkDefaultFont", "TkTextFont", "TkMenuFont"]:
            try:
                f = tkfont.nametofont(font_name)
                f.configure(family=font_family, size=10)
            except Exception:
                pass

        try:
            h_font = tkfont.nametofont("TkHeadingFont")
            h_font.configure(family=font_family, size=11, weight="bold")
        except Exception:
            pass

    def setup_history_combobox(
        self,
        combobox: ttk.Combobox,
        history_key: str,
        initial_val: str,
        on_change_callback: Optional[Callable[[], None]] = None
    ) -> None:
        """
        Настраивает Combobox для поддержки истории ввода:
        при нажатии на инпут выпадает список ранее использованных данных.
        При выборе или завершении ввода значение автоматически сохраняется в историю.
        """
        items = get_field_history(history_key)
        combobox["values"] = items
        if initial_val:
            combobox.set(initial_val)

        def _open_dropdown(event=None):
            # Если клик был по стрелке, ttk сам откроет/закроет выпадающий список
            if event and hasattr(event, "x") and hasattr(event, "y"):
                try:
                    elem = combobox.identify(event.x, event.y)
                    if "arrow" in elem:
                        return
                except Exception:
                    pass

            def _post():
                try:
                    fresh_items = get_field_history(history_key)
                    if fresh_items:
                        combobox["values"] = fresh_items
                        try:
                            combobox.tk.eval(f"{combobox} post")
                        except Exception:
                            try:
                                combobox.event_generate("<Alt-Down>")
                            except Exception:
                                combobox.event_generate("<Down>")
                except Exception:
                    pass

            combobox.after(15, _post)

        # При клике на инпут раскрываем список ранее использованных данных
        combobox.bind("<Button-1>", _open_dropdown, add="+")
        combobox.bind("<Down>", lambda e: combobox.after(10, _open_dropdown), add="+")

        def _commit(event=None):
            val = combobox.get().strip()
            if val and not val.startswith("Название "):
                add_to_field_history(history_key, val)
                combobox["values"] = get_field_history(history_key)
            if on_change_callback:
                try:
                    on_change_callback()
                except Exception:
                    pass

        combobox.bind("<FocusOut>", _commit, add="+")
        combobox.bind("<Return>", _commit, add="+")
        combobox.bind("<<ComboboxSelected>>", _commit, add="+")

    # ---------------------------------------------------------------------
    # ПРОВЕРКА И УВЕДОМЛЕНИЯ О НОВЫХ РЕЛИЗАХ (GITHUB RELEASES)
    # ---------------------------------------------------------------------
    def check_updates_async(self, silent_if_latest: bool = True):
        """
        Проверяет наличие нового релиза на GitHub в фоновом потоке,
        чтобы не блокировать графический интерфейс при медленном интернете.
        """
        def worker():
            try:
                headers = {
                    "Accept": "application/vnd.github.v3+json",
                    "User-Agent": f"NavigatorApp/{APP_VERSION}"
                }
                resp = requests.get(GITHUB_RELEASES_API, headers=headers, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    tag_name = data.get("tag_name", "").strip()
                    release_url = data.get("html_url", GITHUB_RELEASES_PAGE)
                    release_title = data.get("name") or tag_name
                    release_notes = (data.get("body") or "").strip()

                    if is_newer_version(tag_name, APP_VERSION):
                        self.root.after(0, lambda: self._show_update_available_dialog(
                            latest_ver=tag_name,
                            url=release_url,
                            title=release_title,
                            notes=release_notes
                        ))
                    elif not silent_if_latest:
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Проверка обновлений",
                            f"У вас установлена последняя версия!\n\n"
                            f"Текущая версия: v{APP_VERSION}\n"
                            f"Релиз на GitHub: {tag_name or ('v' + APP_VERSION)}"
                        ))
                elif not silent_if_latest:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Проверка обновлений",
                        f"Не удалось получить информацию о релизах с GitHub (Код ответа: {resp.status_code})."
                    ))
            except Exception as e:
                if not silent_if_latest:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Проверка обновлений",
                        f"Ошибка подключения к GitHub при проверке обновлений:\n{e}\n\nПроверьте подключение к интернету."
                    ))

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_available_dialog(self, latest_ver: str, url: str, title: str, notes: str):
        """Отображает модальное окно с предложением скачать новый релиз NavigatorApp.exe."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Доступно обновление программы")
        dlg.geometry("540x410")
        dlg.minsize(480, 340)
        dlg.transient(self.root)
        dlg.grab_set()

        # Центрирование окна относительно главного окна приложения
        try:
            dlg.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - 540) // 2)
            y = ry + max(0, (rh - 410) // 2)
            dlg.geometry(f"+{x}+{y}")
        except Exception:
            pass

        pad = ttk.Frame(dlg, padding=18)
        pad.pack(fill=tk.BOTH, expand=True)

        # Верхняя информационная шапка
        top_frame = ttk.Frame(pad)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        lbl_icon = ttk.Label(top_frame, text="🚀", font=("Segoe UI", 26))
        lbl_icon.pack(side=tk.LEFT, padx=(0, 12))

        header_box = ttk.Frame(top_frame)
        header_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(
            header_box,
            text=f"Доступна новая версия: {latest_ver}",
            font=("Segoe UI", 12, "bold"),
            foreground="#1e40af"
        ).pack(anchor=tk.W)

        ttk.Label(
            header_box,
            text=f"Установленная версия: v{APP_VERSION}",
            font=("Segoe UI", 9),
            foreground="#64748b"
        ).pack(anchor=tk.W, pady=(2, 0))

        # Описание релиза
        box_notes = ttk.LabelFrame(pad, text=f"Описание релиза ({title})", padding=8)
        box_notes.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        txt = tk.Text(box_notes, wrap=tk.WORD, height=7, font=("Segoe UI", 9), relief=tk.FLAT)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(box_notes, orient=tk.VERTICAL, command=txt.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.config(yscrollcommand=sb.set)

        summary_body = notes if notes else "Новый исполняемый файл NavigatorApp.exe готов к загрузке на странице релизов GitHub."
        txt.insert(tk.END, summary_body)
        txt.config(state=tk.DISABLED)

        # Опция автопроверки
        var_auto = tk.BooleanVar(value=bool(self.saved_cfg.get("check_updates", True)))
        def _toggle_auto():
            v = var_auto.get()
            self.saved_cfg["check_updates"] = v
            update_config(check_updates=v)

        chk = ttk.Checkbutton(
            pad,
            text="Автоматически проверять новые релизы при запуске",
            variable=var_auto,
            command=_toggle_auto
        )
        chk.pack(anchor=tk.W, pady=(0, 12))

        # Кнопки действий
        btns = ttk.Frame(pad)
        btns.pack(fill=tk.X)

        def _download_now():
            try:
                webbrowser.open(url)
            except Exception as ex:
                messagebox.showerror("Ошибка", f"Не удалось открыть браузер:\n{url}\n\n{ex}")
            dlg.destroy()

        def _close():
            dlg.destroy()

        btn_go = ttk.Button(
            btns,
            text="📥 Скачать обновление (Открыть GitHub)",
            command=_download_now
        )
        btn_go.pack(side=tk.RIGHT, padx=(6, 0))

        btn_later = ttk.Button(
            btns,
            text="Напомнить позже",
            command=_close
        )
        btn_later.pack(side=tk.RIGHT)

    # ---------------------------------------------------------------------
    # ДИАЛОГ СОЗДАНИЯ БАГ-РЕПОРТА (GITHUB ISSUES)
    # ---------------------------------------------------------------------
    def open_bug_report_dialog(self, initial_tab: str = ""):
        """
        Открывает интерактивное окно создания отчета об ошибке (баг-репорта).
        Позволяет пользователю ввести описание проблемы и отправить её
        в систему GitHub Issues (напрямую по токену либо через предзаполненную веб-форму).
        """
        dlg = tk.Toplevel(self.root)
        dlg.title("Создать баг-репорт (GitHub Issues)")
        dlg.geometry("640x660")
        dlg.minsize(560, 520)
        dlg.transient(self.root)
        dlg.grab_set()

        # Центрирование диалога относительно главного окна
        try:
            dlg.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - 640) // 2)
            y = ry + max(0, (rh - 660) // 2)
            dlg.geometry(f"+{x}+{y}")
        except Exception:
            pass

        pad = ttk.Frame(dlg, padding=16)
        pad.pack(fill=tk.BOTH, expand=True)

        # Верхняя информационная шапка
        top_frame = ttk.Frame(pad)
        top_frame.pack(fill=tk.X, pady=(0, 12))

        lbl_icon = ttk.Label(top_frame, text="🐞", font=("Segoe UI", 26))
        lbl_icon.pack(side=tk.LEFT, padx=(0, 12))

        header_box = ttk.Frame(top_frame)
        header_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(
            header_box,
            text="Сообщить об ошибке (Баг-репорт)",
            font=("Segoe UI", 12, "bold"),
            foreground="#b91c1c"
        ).pack(anchor=tk.W)

        ttk.Label(
            header_box,
            text="Ваш отчёт будет направлен в GitHub Issues для скорейшего исправления.",
            font=("Segoe UI", 8),
            foreground="#64748b"
        ).pack(anchor=tk.W, pady=(2, 0))

        # Поле: Тема ошибки
        box_title = ttk.Frame(pad)
        box_title.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(box_title, text="Краткая тема ошибки (Заголовок):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        entry_title = ttk.Entry(box_title, font=("Segoe UI", 9))
        entry_title.pack(fill=tk.X)
        entry_title.insert(0, "Ошибка при выполнении операции")

        # Поле: В какой вкладке произошла
        box_tab = ttk.Frame(pad)
        box_tab.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(box_tab, text="Раздел программы, в котором возникла ошибка:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))

        tab_choices = [
            "Вкладка 1: Пакетная запись на мероприятие (event_list.xlsx)",
            "Вкладка 2: Подтверждение и участие в мероприятии",
            "Вкладка 3: Зачисление на учебную программу (programm_list.xlsx)",
            "Вкладка 4: Сверка со study_list, подтверждение и обучение",
            "Экран авторизации / Вход в систему",
            "Проверка обновлений или настройки",
            "Другое / Общая ошибка приложения"
        ]
        cmb_section = ttk.Combobox(box_tab, values=tab_choices, state="readonly", font=("Segoe UI", 9))
        cmb_section.pack(fill=tk.X)
        if initial_tab and any(initial_tab in c for c in tab_choices):
            for c in tab_choices:
                if initial_tab in c:
                    cmb_section.set(c)
                    break
        else:
            cmb_section.set(tab_choices[0])

        # Поле: Описание и шаги воспроизведения
        box_desc = ttk.Frame(pad)
        box_desc.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        ttk.Label(box_desc, text="Подробное описание проблемы и шаги для воспроизведения:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))

        txt_desc = tk.Text(box_desc, wrap=tk.WORD, height=7, font=("Segoe UI", 9))
        txt_desc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_desc = ttk.Scrollbar(box_desc, orient=tk.VERTICAL, command=txt_desc.yview)
        sb_desc.pack(side=tk.RIGHT, fill=tk.Y)
        txt_desc.config(yscrollcommand=sb_desc.set)

        placeholder_text = (
            "1. Что делали (какие кнопки нажимали, какие данные вводили):\n"
            "   \n"
            "2. Что ожидали получить:\n"
            "   \n"
            "3. Что произошло на самом деле (сообщение или поведение):\n"
            "   \n"
        )
        txt_desc.insert(tk.END, placeholder_text)

        # Чекбокс: Прикрепить диагностические данные
        var_attach_diag = tk.BooleanVar(value=True)
        chk_diag = ttk.Checkbutton(
            pad,
            text=f"Прикрепить технические данные (ОС, версия v{APP_VERSION}, последние строки лога сессии)",
            variable=var_attach_diag
        )
        chk_diag.pack(anchor=tk.W, pady=(0, 8))

        # Опциональный GitHub Personal Access Token
        box_token = ttk.LabelFrame(pad, text="Прямая отправка через GitHub Token (необязательно)", padding=8)
        box_token.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            box_token,
            text="Укажите токен только если хотите отправить репорт напрямую без открытия браузера. Иначе оставьте пустым.",
            font=("Segoe UI", 8),
            foreground="#64748b"
        ).pack(anchor=tk.W, pady=(0, 4))

        token_row = ttk.Frame(box_token)
        token_row.pack(fill=tk.X)

        var_token = tk.StringVar(value=str(self.saved_cfg.get("github_token", "")))
        entry_token = ttk.Entry(token_row, textvariable=var_token, show="*", font=("Segoe UI", 9))
        entry_token.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        var_save_token = tk.BooleanVar(value=bool(self.saved_cfg.get("github_token")))
        ttk.Checkbutton(token_row, text="Сохранить в config.json", variable=var_save_token).pack(side=tk.RIGHT)

        # Сборка текста репорта в формате Markdown
        def _build_markdown_report() -> Tuple[str, str]:
            t_input = entry_title.get().strip()
            title_clean = t_input if t_input else "Сообщение об ошибке"
            section_val = cmb_section.get()
            desc_val = txt_desc.get("1.0", tk.END).strip()

            body_lines = [
                f"### Описание проблемы\n{desc_val}\n",
                f"### Раздел программы\n- {section_val}\n"
            ]

            if var_attach_diag.get():
                os_name = f"{platform.system()} {platform.release()} ({platform.version()})"
                py_ver = sys.version.split()[0]
                exec_mode = "EXE (автономная сборка PyInstaller)" if getattr(sys, "frozen", False) else "Python скрипт (исходный код)"
                
                body_lines.append(
                    f"### Окружение (Диагностика)\n"
                    f"- **Версия NavigatorApp:** v{APP_VERSION}\n"
                    f"- **Операционная система:** {os_name}\n"
                    f"- **Режим запуска:** {exec_mode}\n"
                    f"- **Версия Python:** {py_ver}\n"
                )

                # Безопасное извлечение последних 15 строк журнала сессии
                tail_logs = []
                log_candidate = CURRENT_LOG_FILE if CURRENT_LOG_FILE and os.path.exists(CURRENT_LOG_FILE) else LOG_FILE
                if log_candidate and os.path.exists(log_candidate):
                    try:
                        import re
                        with open(log_candidate, "r", encoding="utf-8", errors="ignore") as lf:
                            raw_lines = lf.readlines()
                            for line in raw_lines[-15:]:
                                l_clean = re.sub(r'("password"|"token"|"pwd"):\s*"[^"]+"', r'\1: "***"', line)
                                tail_logs.append(l_clean)
                    except Exception as ex:
                        tail_logs.append(f"Не удалось прочитать лог: {ex}\n")

                if tail_logs:
                    body_lines.append(f"### Журнал событий (последние строки)\n```text\n{''.join(tail_logs).strip()}\n```\n")

            full_body = "\n".join(body_lines)
            return title_clean, full_body

        # Нижняя панель действий
        btn_bar = ttk.Frame(pad)
        btn_bar.pack(fill=tk.X)

        status_lbl = ttk.Label(btn_bar, text="", font=("Segoe UI", 9))
        status_lbl.pack(side=tk.LEFT)

        def _copy_markdown():
            title_val, body_val = _build_markdown_report()
            full_text = f"# [BUG] {title_val}\n\n{body_val}"
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(full_text)
                status_lbl.config(text="✓ Скопировано в буфер обмена!", foreground="#16a34a")
            except Exception as ex:
                messagebox.showerror("Ошибка", f"Не удалось скопировать: {ex}")

        def _submit_report():
            title_val, body_val = _build_markdown_report()
            token = var_token.get().strip()

            # Сохранение токена при запросе пользователя
            if var_save_token.get():
                self.saved_cfg["github_token"] = token
                update_config(github_token=token)
            elif "github_token" in self.saved_cfg and not token:
                self.saved_cfg["github_token"] = ""
                update_config(github_token="")

            # Если указан токен, отправляем напрямую через GitHub REST API
            if token:
                status_lbl.config(text="Отправка в GitHub Issues...", foreground="#2563eb")
                dlg.update_idletasks()

                def send_worker():
                    try:
                        headers = {
                            "Authorization": f"Bearer {token}",
                            "Accept": "application/vnd.github.v3+json",
                            "User-Agent": f"NavigatorApp/{APP_VERSION}"
                        }
                        payload = {
                            "title": f"[BUG] {title_val}",
                            "body": body_val,
                            "labels": ["bug"]
                        }
                        resp = requests.post(GITHUB_ISSUES_API, headers=headers, json=payload, timeout=10)
                        if resp.status_code == 201:
                            issue_data = resp.json()
                            issue_url = issue_data.get("html_url", GITHUB_ISSUES_URL)
                            issue_num = issue_data.get("number", "")

                            def on_success():
                                status_lbl.config(text=f"✓ Успешно опубликован issue #{issue_num}!", foreground="#16a34a")
                                if messagebox.askyesno(
                                    "Баг-репорт создан",
                                    f"Баг-репорт успешно отправлен в GitHub Issues (Issue #{issue_num})!\n\nОткрыть созданный тикет в браузере?"
                                ):
                                    webbrowser.open(issue_url)
                                dlg.destroy()

                            self.root.after(0, on_success)
                        else:
                            def on_fail():
                                status_lbl.config(text="Ошибка отправки по API", foreground="#dc2626")
                                err_msg = f"GitHub API вернул код {resp.status_code}."
                                if messagebox.askyesno(
                                    "Ошибка отправки",
                                    f"{err_msg}\n\nОткрыть веб-форму GitHub в браузере с заполненными данными?"
                                ):
                                    _open_in_browser(title_val, body_val)
                                    dlg.destroy()

                            self.root.after(0, on_fail)
                    except Exception as ex:
                        def on_exc():
                            status_lbl.config(text="Ошибка соединения", foreground="#dc2626")
                            if messagebox.askyesno(
                                "Ошибка сети",
                                f"Не удалось связаться с GitHub:\n{ex}\n\nОткрыть страницу создания Issue в браузере?"
                            ):
                                _open_in_browser(title_val, body_val)
                                dlg.destroy()

                        self.root.after(0, on_exc)

                threading.Thread(target=send_worker, daemon=True).start()
            else:
                # Если токен не указан — открываем GitHub New Issue с предзаполненными полями
                _open_in_browser(title_val, body_val)
                dlg.destroy()

        def _open_in_browser(title_val: str, body_val: str):
            try:
                encoded_title = urllib.parse.quote_plus(f"[BUG] {title_val}")
                encoded_body = urllib.parse.quote_plus(body_val)
                issue_create_url = f"{GITHUB_NEW_ISSUE_URL}?title={encoded_title}&body={encoded_body}&labels=bug"
                webbrowser.open(issue_create_url)
            except Exception as ex:
                messagebox.showerror("Ошибка", f"Не удалось открыть браузер:\n{ex}")

        btn_copy = ttk.Button(
            btn_bar,
            text="📋 Скопировать текст",
            command=_copy_markdown
        )
        btn_copy.pack(side=tk.RIGHT, padx=(4, 0))

        btn_send = ttk.Button(
            btn_bar,
            text="🚀 Отправить на GitHub",
            command=_submit_report
        )
        btn_send.pack(side=tk.RIGHT, padx=(4, 0))

        btn_cancel = ttk.Button(
            btn_bar,
            text="Закрыть",
            command=dlg.destroy
        )
        btn_cancel.pack(side=tk.RIGHT)

    # ---------------------------------------------------------------------
    # ЭКРАН 1: АВТОРИЗАЦИЯ (с выбором: прошлый или новый пользователь)
    # ---------------------------------------------------------------------
    def show_login_screen(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

        container = ttk.Frame(self.main_frame)
        container.place(relx=0.5, rely=0.45, anchor=tk.CENTER)

        ttk.Label(container, text="Вход в Навигатор ДО", font=("Segoe UI", 16, "bold")).pack(pady=(0, 15))

        has_saved = bool(self.saved_cfg.get("email") and self.saved_cfg.get("password"))

        if has_saved:
            saved_box = ttk.LabelFrame(container, text="Сохраненный аккаунт", padding=15)
            saved_box.pack(fill=tk.X, pady=(0, 15))

            user_mail = self.saved_cfg.get("email")
            ttk.Label(saved_box, text=f"Пользователь: {user_mail}", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
            ttk.Label(saved_box, text="Вход будет выполнен с сохраненным паролем.", font=("Segoe UI", 9), foreground="gray").pack(anchor=tk.W, pady=(2, 8))

            btn_login_prev = ttk.Button(
                saved_box,
                text=f"➜ Войти как {user_mail}",
                command=lambda: self.do_login(self.saved_cfg["email"], self.saved_cfg["password"])
            )
            btn_login_prev.pack(fill=tk.X, pady=4)

            ttk.Separator(container, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # Форма нового входа
        form_title = "Войти как другой пользователь" if has_saved else "Введите учетные данные"
        new_box = ttk.LabelFrame(container, text=form_title, padding=15)
        new_box.pack(fill=tk.X)

        ttk.Label(new_box, text="Email / Логин:").pack(anchor=tk.W)
        self.entry_email = ttk.Combobox(new_box, width=38, font=("Segoe UI", 10))
        self.entry_email.pack(fill=tk.X, pady=(2, 8))
        self.setup_history_combobox(self.entry_email, "email", self.saved_cfg.get("email", ""))

        ttk.Label(new_box, text="Пароль:").pack(anchor=tk.W)
        self.entry_password = ttk.Entry(new_box, width=38, show="•", font=("Segoe UI", 10))
        self.entry_password.pack(fill=tk.X, pady=(2, 8))

        self.var_save_pass = tk.BooleanVar(value=True)
        ttk.Checkbutton(new_box, text="Запомнить логин и пароль в config.json", variable=self.var_save_pass).pack(anchor=tk.W, pady=(2, 10))

        btn_new_login = ttk.Button(
            new_box,
            text="Войти в систему",
            command=self.on_new_login_click
        )
        btn_new_login.pack(fill=tk.X)

        self.login_status_lbl = ttk.Label(container, text="", font=("Segoe UI", 9))
        self.login_status_lbl.pack(pady=(10, 0))

        # Нижняя плашка: версия программы и ручная проверка обновлений
        footer_box = ttk.Frame(container)
        footer_box.pack(fill=tk.X, pady=(15, 0))

        lbl_v = ttk.Label(footer_box, text=f"Версия: v{APP_VERSION}", font=("Segoe UI", 8), foreground="gray")
        lbl_v.pack(side=tk.LEFT)

        btn_chk = ttk.Button(
            footer_box,
            text="🔄 Проверить обновления",
            command=lambda: self.check_updates_async(silent_if_latest=False)
        )
        btn_chk.pack(side=tk.RIGHT, padx=(4, 0))

        btn_bug = ttk.Button(
            footer_box,
            text="🐞 Создать баг-репорт",
            command=self.open_bug_report_dialog
        )
        btn_bug.pack(side=tk.RIGHT)

    def on_new_login_click(self):
        email = self.entry_email.get().strip()
        pwd = self.entry_password.get().strip()
        if not email or not pwd:
            messagebox.showwarning("Внимание", "Пожалуйста, заполните email и пароль.")
            return
        add_to_field_history("email", email)
        if self.var_save_pass.get():
            save_config(email, pwd)
            self.saved_cfg = load_config()
        self.do_login(email, pwd)

    def do_login(self, email: str, pwd: str):
        self.login_status_lbl.config(text="Авторизация...", foreground="blue")
        self.root.update()

        def worker():
            ok, msg = self.client.login(email, pwd)
            self.root.after(0, lambda: self._on_login_result(ok, msg, email, pwd))

        threading.Thread(target=worker, daemon=True).start()

    def _on_login_result(self, ok: bool, msg: str, email: str, pwd: str):
        if ok:
            save_config(email, pwd)
            u_info = getattr(self.client, "user_info", {}) or {}
            user_title = u_info.get("name") or u_info.get("email") or email
            write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🔑 УСПЕШНАЯ АВТОРИЗАЦИЯ: {user_title} ({email})")
            self.show_main_dashboard()
        else:
            write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ❌ ОШИБКА АВТОРИЗАЦИИ: {email} | {msg}")
            self.login_status_lbl.config(text=f"Ошибка: {msg}", foreground="red")
            messagebox.showerror("Ошибка входа", msg)

    # ---------------------------------------------------------------------
    # ЭКРАН 2: ГЛАВНАЯ ПАНЕЛЬ С КНОПКАМИ И ИНПУТАМИ
    # ---------------------------------------------------------------------
    def show_main_dashboard(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

        # Верхняя информационная панель
        top_bar = ttk.Frame(self.main_frame)
        top_bar.pack(fill=tk.X, pady=(0, 10))

        user_name = self.client.user_info.get("name") or self.client.user_info.get("email") or "Администратор"
        ttk.Label(top_bar, text=f"Пользователь: {user_name}", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)

        btn_box = ttk.Frame(top_bar)
        btn_box.pack(side=tk.RIGHT)

        log_btn_title = f"📋 Лог ({os.path.basename(CURRENT_LOG_FILE)})" if CURRENT_LOG_FILE else "📋 Журнал сессии"
        ttk.Button(btn_box, text="📊 event_list.xlsx", command=lambda: open_file_in_os(EVENT_EXCEL_FILE)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="🎓 programm_list.xlsx", command=lambda: open_file_in_os(PROGRAM_EXCEL_FILE)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="📑 study_list.xlsx", command=lambda: open_file_in_os(STUDY_EXCEL_FILE)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text=log_btn_title, command=lambda: open_file_in_os(CURRENT_LOG_FILE or LOG_FILE)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="📁 Папка логов", command=lambda: open_file_in_os(os.path.abspath(LOGS_DIR))).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="⚙ config.json", command=lambda: open_file_in_os(CONFIG_FILE)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="🔄 Обновления", command=lambda: self.check_updates_async(silent_if_latest=False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box, text="🐞 Баг-репорт", command=self.open_bug_report_dialog).pack(side=tk.LEFT, padx=2)
        
        def _on_logout():
            write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🚪 ВЫХОД ИЗ УЧЕТНОЙ ЗАПИСИ: {user_name}")
            self.show_login_screen()

        ttk.Button(btn_box, text="Выйти", command=_on_logout).pack(side=tk.LEFT, padx=2)

        # Разделитель
        ttk.Separator(self.main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 8))

        # Виджет мониторинга нагрузки на сервер в реальном времени
        mon_frame = ttk.LabelFrame(self.main_frame, text=" ⚡ Монитор сервера и защита от перегрузки (в реальном времени) ", padding=(10, 6))
        mon_frame.pack(fill=tk.X, pady=(0, 10))

        mon_left = ttk.Frame(mon_frame)
        mon_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.lbl_server_ping = ttk.Label(
            mon_left,
            text="● Отклик сервера: замер связи... | Нагрузка: вычисление...",
            font=("Segoe UI", 9, "bold"),
            foreground="#2563eb"
        )
        self.lbl_server_ping.pack(anchor=tk.W)

        self.lbl_server_speed = ttk.Label(
            mon_left,
            text="Скорость отправки: 0.8 запр/сек (пауза 1.2 с) | Режим: Защита от перегрузки (строго ≤ 1 запр/сек)",
            font=("Segoe UI", 8),
            foreground="#64748b"
        )
        self.lbl_server_speed.pack(anchor=tk.W, pady=(2, 0))

        mon_right = ttk.Frame(mon_frame)
        mon_right.pack(side=tk.RIGHT)

        self.btn_ping_now = ttk.Button(mon_right, text="🔄 Проверить отклик", command=self.manual_check_ping)
        self.btn_ping_now.pack(side=tk.RIGHT, padx=4)

        # Сразу запускаем первый быстрый замер при открытии дашборда
        self.manual_check_ping()

        # Загружаем актуальные данные из config.json
        self.saved_cfg = load_config()
        initial_activity = self.saved_cfg.get("activity_name") or DEFAULT_CONFIG["activity_name"]
        initial_datetime = self.saved_cfg.get("activity_datetime") or DEFAULT_CONFIG["activity_datetime"]
        initial_confirm_act = self.saved_cfg.get("confirm_activity_name") or initial_activity

        # Основной блок с режимами работы (Мероприятия, Подтверждение, Зачисление на программу, Подтверждение и обучение)
        tabs = ttk.Notebook(self.main_frame)
        tabs.pack(fill=tk.X, pady=(0, 10))

        def _on_tab_changed(event):
            try:
                sel_id = tabs.select()
                if sel_id:
                    tab_name = tabs.tab(sel_id, "text")
                    write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 📌 ПЕРЕКЛЮЧЕНИЕ НА ВКЛАДКУ: {tab_name.strip()}")
            except Exception:
                pass

        tabs.bind("<<NotebookTabChanged>>", _on_tab_changed)

        # --- ВКЛАДКА 1: ПАКЕТНАЯ ЗАПИСЬ (NavAdd) ---
        tab_add = ttk.Frame(tabs, padding=12)
        tabs.add(tab_add, text=" 📝 1. Запись на мероприятие ")
        tab_add.columnconfigure(1, weight=1)

        ttk.Label(tab_add, text="Название мероприятия:").grid(row=0, column=0, sticky=tk.W, pady=(4, 0))
        
        add_event_box = ttk.Frame(tab_add)
        add_event_box.grid(row=0, column=1, columnspan=2, sticky=tk.W + tk.E, pady=(4, 0))
        add_event_box.columnconfigure(0, weight=1)

        self.entry_add_event = ttk.Combobox(add_event_box, font=("Segoe UI", 10))
        self.entry_add_event.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_add_event = ttk.Button(
            add_event_box,
            text="🔍 Проверить",
            command=self.check_add_event_info
        )
        self.btn_check_add_event.grid(row=0, column=1, sticky=tk.E)

        # Подсказка-пример валидного ввода для мероприятия
        ttk.Label(
            tab_add,
            text="💡 Нажмите на поле для выбора из истории. Пример: Мастер-класс по робототехнике",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(tab_add, text="Дата и время участия:").grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        self.entry_add_dt = ttk.Combobox(tab_add, font=("Segoe UI", 10))
        self.entry_add_dt.grid(row=2, column=1, columnspan=2, sticky=tk.W + tk.E, pady=(4, 0))

        # Привязываем авто-сохранение в config.json
        def save_add_fields(event=None):
            try:
                name = self.entry_add_event.get().strip() if hasattr(self, "entry_add_event") else ""
                dt = self.entry_add_dt.get().strip() if hasattr(self, "entry_add_dt") else ""
                if name or dt:
                    update_config(activity_name=name, activity_datetime=dt)
                    self.saved_cfg = load_config()
            except Exception:
                pass

        self.setup_history_combobox(self.entry_add_event, "activity_name", initial_activity, save_add_fields)
        self.setup_history_combobox(self.entry_add_dt, "activity_datetime", initial_datetime, save_add_fields)

        # Подсказка-пример валидного формата даты/времени
        ttk.Label(
            tab_add,
            text="Формат: ГГГГ-ММ-ДД ЧЧ:ММ:СС (пример: 2026-08-31 11:00:00, сохраняется в config.json)",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=3, column=1, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(tab_add, text="Файл таблицы:", foreground="gray").grid(row=4, column=0, sticky=tk.W, pady=4)
        tab_add_file_box = ttk.Frame(tab_add)
        tab_add_file_box.grid(row=4, column=1, columnspan=2, sticky=tk.W + tk.E, pady=4)
        ttk.Label(tab_add_file_box, text=f"Используется '{EVENT_EXCEL_FILE}' (мероприятия)", foreground="#2563eb", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(tab_add_file_box, text="📂 Открыть event_list.xlsx", command=lambda: open_file_in_os(EVENT_EXCEL_FILE)).pack(side=tk.LEFT)

        self.btn_run_add = ttk.Button(tab_add, text="▶ Запустить запись детей на мероприятие из таблицы", command=self.run_batch_add)
        self.btn_run_add.grid(row=5, column=0, columnspan=3, sticky=tk.W + tk.E, pady=(8, 4))

        # --- ВКЛАДКА 2: ПОДТВЕРЖДЕНИЕ ЗАЯВОК И ОТМЕТКА УЧАСТИЯ (NavConfirm) ---
        tab_conf = ttk.Frame(tabs, padding=12)
        tabs.add(tab_conf, text=" ✓ 2. Подтверждение и отметка участия ")
        tab_conf.columnconfigure(1, weight=1)

        ttk.Label(tab_conf, text="Название мероприятия:").grid(row=0, column=0, sticky=tk.W, pady=(4, 0))
        
        conf_event_box = ttk.Frame(tab_conf)
        conf_event_box.grid(row=0, column=1, sticky=tk.W + tk.E, pady=(4, 0))
        conf_event_box.columnconfigure(0, weight=1)

        self.entry_conf_event = ttk.Combobox(conf_event_box, font=("Segoe UI", 10))
        self.entry_conf_event.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_conf_event = ttk.Button(
            conf_event_box,
            text="🔍 Проверить",
            command=self.check_conf_event_info
        )
        self.btn_check_conf_event.grid(row=0, column=1, sticky=tk.E)

        def save_conf_fields(event=None):
            try:
                name = self.entry_conf_event.get().strip() if hasattr(self, "entry_conf_event") else ""
                if name:
                    update_config(confirm_activity_name=name)
                    self.saved_cfg = load_config()
            except Exception:
                pass

        self.setup_history_combobox(self.entry_conf_event, "activity_name", initial_confirm_act, save_conf_fields)

        ttk.Label(
            tab_conf,
            text="💡 Нажмите на поле для выбора из истории. Пример: Мастер-класс по робототехнике",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=1, column=1, sticky=tk.W, pady=(0, 6))

        self.btn_check_orders = ttk.Button(
            tab_conf,
            text="🔍 Проверить статус заявок мероприятия",
            command=self.check_pending_orders
        )
        self.btn_check_orders.grid(row=2, column=0, columnspan=2, sticky=tk.W + tk.E, pady=(6, 4))

        # Блок с действиями (Подтверждение и Отметка участия отдельными кнопками)
        actions_box = ttk.Frame(tab_conf)
        actions_box.grid(row=3, column=0, columnspan=2, sticky=tk.W + tk.E, pady=4)
        actions_box.columnconfigure(0, weight=1)
        actions_box.columnconfigure(1, weight=1)

        self.btn_run_confirm = ttk.Button(
            actions_box,
            text="✓ 1. Подтвердить заявки (state: approve)",
            command=self.run_batch_confirm,
            state=tk.DISABLED
        )
        self.btn_run_confirm.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 4), pady=2)

        self.btn_run_participant = ttk.Button(
            actions_box,
            text="🎖️ 2. Отметить участие (state: participant)",
            command=self.run_batch_participant,
            state=tk.DISABLED
        )
        self.btn_run_participant.grid(row=0, column=1, sticky=tk.W + tk.E, padx=(4, 0), pady=2)

        ttk.Label(
            tab_conf,
            text="💡 Сначала подтвердите заявки (кнопка 1), затем отметьте участие тех, кто был подтвержден (кнопка 2).",
            foreground="#64748b",
            font=("Segoe UI", 8, "italic")
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(3, 2))

        # --- ВКЛАДКА 3: ЗАЧИСЛЕНИЕ НА УЧЕБНУЮ ПРОГРАММУ (POST /api/rest/order) ---
        tab_program = ttk.Frame(tabs, padding=12)
        tabs.add(tab_program, text=" 🎓 3. Зачисление на программу ")

        initial_prog_name = str(self.saved_cfg.get("program_name") or DEFAULT_CONFIG.get("program_name", "Основы программирования и робототехники"))
        initial_prog_group = str(self.saved_cfg.get("program_group_name") or self.saved_cfg.get("program_group_id") or DEFAULT_CONFIG.get("program_group_name", "Группа 1"))
        initial_prog_year = str(self.saved_cfg.get("program_academic_year_id") or DEFAULT_CONFIG["program_academic_year_id"])
        initial_prog_use_cert = bool(self.saved_cfg.get("program_use_certificate", DEFAULT_CONFIG["program_use_certificate"]))
        initial_prog_create_cert = bool(self.saved_cfg.get("program_create_certificate", DEFAULT_CONFIG["program_create_certificate"]))

        # Название программы (автоматический поиск в Навигаторе по GET /api/rest/events)
        ttk.Label(tab_program, text="Название программы:").grid(row=0, column=0, sticky=tk.W, pady=(3, 0))
        
        prog_name_box = ttk.Frame(tab_program)
        prog_name_box.grid(row=0, column=1, sticky=tk.W + tk.E, pady=(3, 0))
        tab_program.columnconfigure(1, weight=1)
        prog_name_box.columnconfigure(0, weight=1)

        self.entry_prog_name = ttk.Combobox(prog_name_box, font=("Segoe UI", 10))
        self.entry_prog_name.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_prog = ttk.Button(
            prog_name_box,
            text="🔍 Проверить",
            command=self.check_program_info
        )
        self.btn_check_prog.grid(row=0, column=1, sticky=tk.E)

        ttk.Label(
            tab_program,
            text="💡 Нажмите на поле для выбора из истории. Поиск в Навигаторе (GET /api/rest/events).",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=1, column=1, sticky=tk.W, pady=(0, 4))

        # Группа (название или ID, поиск через GET /api/rest/eventGroups)
        ttk.Label(tab_program, text="Группа (название или ID):").grid(row=2, column=0, sticky=tk.W, pady=(3, 0))

        prog_group_box = ttk.Frame(tab_program)
        prog_group_box.grid(row=2, column=1, sticky=tk.W + tk.E, pady=(3, 0))
        prog_group_box.columnconfigure(0, weight=1)

        self.entry_prog_group = ttk.Combobox(prog_group_box, font=("Segoe UI", 10))
        self.entry_prog_group.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_group = ttk.Button(
            prog_group_box,
            text="🔍 Проверить",
            command=self.check_group_info
        )
        self.btn_check_group.grid(row=0, column=1, sticky=tk.E)

        ttk.Label(
            tab_program,
            text="💡 Нажмите на поле для истории. Поиск по названию (GET /api/rest/eventGroups) или точный ID.",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=3, column=1, sticky=tk.W, pady=(0, 4))

        # academic_year_id (поддержка обоих форматов: 2026/2027 и 2026)
        ttk.Label(tab_program, text="Учебный год (2026/2027 или 2026):").grid(row=4, column=0, sticky=tk.W, pady=(3, 0))
        self.entry_prog_year = ttk.Combobox(tab_program, font=("Segoe UI", 10))
        self.entry_prog_year.grid(row=4, column=1, sticky=tk.W + tk.E, pady=(3, 0))

        ttk.Label(
            tab_program,
            text="Поддерживает форматы: 2026/2027 или 2026 (в Навигатор передается academic_year_id=2026)",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=5, column=1, sticky=tk.W, pady=(0, 5))

        # Чекбоксы сертификатов
        opts_frame = ttk.Frame(tab_program)
        opts_frame.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(2, 6))

        self.var_prog_create_cert = tk.BooleanVar(value=initial_prog_create_cert)
        self.chk_prog_create_cert = ttk.Checkbutton(
            opts_frame,
            text="Создавать сертификат (create_certificate: true)",
            variable=self.var_prog_create_cert
        )
        self.chk_prog_create_cert.pack(side=tk.LEFT, padx=(0, 16))

        self.var_prog_use_cert = tk.BooleanVar(value=initial_prog_use_cert)
        self.chk_prog_use_cert = ttk.Checkbutton(
            opts_frame,
            text="Использовать сертификат (use_certificate)",
            variable=self.var_prog_use_cert
        )
        self.chk_prog_use_cert.pack(side=tk.LEFT)

        def save_prog_fields(event=None):
            try:
                p_name = self.entry_prog_name.get().strip() if hasattr(self, "entry_prog_name") else ""
                g_val = self.entry_prog_group.get().strip() if hasattr(self, "entry_prog_group") else ""
                y_id = self.entry_prog_year.get().strip() if hasattr(self, "entry_prog_year") else ""
                c_cert = self.var_prog_create_cert.get() if hasattr(self, "var_prog_create_cert") else True
                u_cert = self.var_prog_use_cert.get() if hasattr(self, "var_prog_use_cert") else False
                update_config(
                    program_name=p_name,
                    program_group_name=g_val,
                    program_group_id=g_val if g_val.isdigit() else str(self.saved_cfg.get("program_group_id") or ""),
                    program_academic_year_id=y_id,
                    program_create_certificate=c_cert,
                    program_use_certificate=u_cert
                )
                self.saved_cfg = load_config()
            except Exception:
                pass

        self.setup_history_combobox(self.entry_prog_name, "program_name", initial_prog_name, save_prog_fields)
        self.setup_history_combobox(self.entry_prog_group, "program_group_name", initial_prog_group, save_prog_fields)
        self.setup_history_combobox(self.entry_prog_year, "academic_year_id", initial_prog_year, save_prog_fields)
        self.chk_prog_create_cert.configure(command=save_prog_fields)
        self.chk_prog_use_cert.configure(command=save_prog_fields)

        ttk.Label(tab_program, text="Файл таблицы:", foreground="gray").grid(row=7, column=0, sticky=tk.W, pady=3)
        tab_prog_file_box = ttk.Frame(tab_program)
        tab_prog_file_box.grid(row=7, column=1, sticky=tk.W, pady=3)
        ttk.Label(
            tab_prog_file_box,
            text=f"Используется '{PROGRAM_EXCEL_FILE}' (ФИО, Дата рождения, Статус)",
            foreground="#2563eb",
            font=("Segoe UI", 9, "italic")
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(tab_prog_file_box, text="📂 Открыть programm_list.xlsx", command=lambda: open_file_in_os(PROGRAM_EXCEL_FILE)).pack(side=tk.LEFT)

        self.btn_run_program = ttk.Button(
            tab_program,
            text="▶ Запустить зачисление детей на программу из таблицы",
            command=self.run_batch_program_enroll
        )
        self.btn_run_program.grid(row=8, column=0, columnspan=2, sticky=tk.W + tk.E, pady=(6, 2))

        # --- ВКЛАДКА 4: ПОДТВЕРЖДЕНИЕ И ОБУЧЕНИЕ НА ПРОГРАММУ (POST /api/approveRequest, POST /api/studyRequest) ---
        tab_study = ttk.Frame(tabs, padding=12)
        tabs.add(tab_study, text=" 📖 4. Подтверждение и обучение (программа) ")

        initial_study_prog = str(self.saved_cfg.get("study_program_name") or self.saved_cfg.get("program_name") or DEFAULT_CONFIG["study_program_name"])
        initial_study_group = str(self.saved_cfg.get("study_program_group_name", ""))
        initial_study_year = str(self.saved_cfg.get("study_academic_year_id") or self.saved_cfg.get("program_academic_year_id") or DEFAULT_CONFIG["study_academic_year_id"])
        initial_decree = str(self.saved_cfg.get("study_decree_number", DEFAULT_CONFIG["study_decree_number"]))
        initial_signing = str(self.saved_cfg.get("study_date_signing", DEFAULT_CONFIG["study_date_signing"]))
        initial_start = str(self.saved_cfg.get("study_date_start", DEFAULT_CONFIG["study_date_start"]))
        initial_fin = str(self.saved_cfg.get("study_financing_source", DEFAULT_CONFIG["study_financing_source"]))

        # Название программы
        ttk.Label(tab_study, text="Название программы:").grid(row=0, column=0, sticky=tk.W, pady=(3, 0))
        study_prog_box = ttk.Frame(tab_study)
        study_prog_box.grid(row=0, column=1, sticky=tk.W + tk.E, pady=(3, 0))
        tab_study.columnconfigure(1, weight=1)
        study_prog_box.columnconfigure(0, weight=1)

        self.entry_study_prog = ttk.Combobox(study_prog_box, font=("Segoe UI", 10))
        self.entry_study_prog.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_study_prog = ttk.Button(
            study_prog_box,
            text="🔍 Проверить",
            command=self.check_study_program_info
        )
        self.btn_check_study_prog.grid(row=0, column=1, sticky=tk.E)

        ttk.Label(
            tab_study,
            text="💡 Нажмите на поле для выбора из истории. Поиск программы в Навигаторе (GET /api/rest/events).",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=1, column=1, sticky=tk.W, pady=(0, 4))

        # Название или ID группы
        ttk.Label(tab_study, text="Группа (название/ID):").grid(row=2, column=0, sticky=tk.W, pady=(3, 0))
        study_group_box = ttk.Frame(tab_study)
        study_group_box.grid(row=2, column=1, sticky=tk.W + tk.E, pady=(3, 0))
        study_group_box.columnconfigure(0, weight=1)

        self.entry_study_group = ttk.Combobox(study_group_box, font=("Segoe UI", 10))
        self.entry_study_group.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 6))

        self.btn_check_study_group = ttk.Button(
            study_group_box,
            text="🔍 Проверить",
            command=self.check_study_group_info
        )
        self.btn_check_study_group.grid(row=0, column=1, sticky=tk.E)

        ttk.Label(
            tab_study,
            text="💡 Нажмите на поле для истории. Оставьте пустым для зачисления ВСЕХ групп, либо укажите название/ID.",
            foreground="#2563eb",
            font=("Segoe UI", 8, "italic")
        ).grid(row=3, column=1, sticky=tk.W, pady=(0, 4))

        # Учебный год
        ttk.Label(tab_study, text="Учебный год:").grid(row=4, column=0, sticky=tk.W, pady=(3, 0))
        self.entry_study_year = ttk.Combobox(tab_study, font=("Segoe UI", 10))
        self.entry_study_year.grid(row=4, column=1, sticky=tk.W + tk.E, pady=(3, 0))

        ttk.Label(
            tab_study,
            text="Поддерживает форматы: 2026/2027 или 2026 (в фильтр передается academic_year_id=2026)",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=5, column=1, sticky=tk.W, pady=(0, 5))

        # Параметры приказа о зачислении (для POST /api/studyRequest)
        decree_frame = ttk.LabelFrame(tab_study, text="Параметры приказа о зачислении на обучение (для studyRequest)", padding=(8, 4))
        decree_frame.grid(row=6, column=0, columnspan=2, sticky=tk.W + tk.E, pady=(3, 6))
        decree_frame.columnconfigure(1, weight=1)
        decree_frame.columnconfigure(3, weight=1)
        decree_frame.columnconfigure(5, weight=1)

        ttk.Label(decree_frame, text="№ приказа:").grid(row=0, column=0, sticky=tk.W, padx=(0, 4))
        self.entry_study_decree = ttk.Combobox(decree_frame, font=("Segoe UI", 9))
        self.entry_study_decree.grid(row=0, column=1, sticky=tk.W + tk.E, padx=(0, 10))

        ttk.Label(decree_frame, text="Дата приказа:").grid(row=0, column=2, sticky=tk.W, padx=(0, 4))
        self.entry_study_signing = ttk.Combobox(decree_frame, font=("Segoe UI", 9))
        self.entry_study_signing.grid(row=0, column=3, sticky=tk.W + tk.E, padx=(0, 10))

        ttk.Label(decree_frame, text="Дата начала:").grid(row=0, column=4, sticky=tk.W, padx=(0, 4))
        self.entry_study_start = ttk.Combobox(decree_frame, font=("Segoe UI", 9))
        self.entry_study_start.grid(row=0, column=5, sticky=tk.W + tk.E, padx=(0, 10))

        def save_study_fields(event=None):
            try:
                p_name = self.entry_study_prog.get().strip() if hasattr(self, "entry_study_prog") else ""
                g_val = self.entry_study_group.get().strip() if hasattr(self, "entry_study_group") else ""
                y_id = self.entry_study_year.get().strip() if hasattr(self, "entry_study_year") else ""
                dec = self.entry_study_decree.get().strip() if hasattr(self, "entry_study_decree") else "183"
                
                raw_sign = self.entry_study_signing.get().strip() if hasattr(self, "entry_study_signing") else "2026-08-31"
                raw_start = self.entry_study_start.get().strip() if hasattr(self, "entry_study_start") else "2026-09-01"
                sign = normalize_date_input(raw_sign, "2026-08-31")
                st_d = normalize_date_input(raw_start, "2026-09-01")

                fin = self.saved_cfg.get("study_financing_source", "1")
                update_config(
                    study_program_name=p_name,
                    study_program_group_name=g_val,
                    study_academic_year_id=y_id,
                    study_decree_number=dec,
                    study_date_signing=sign,
                    study_date_start=st_d,
                    study_financing_source=fin
                )
                self.saved_cfg = load_config()
            except Exception:
                pass

        self.setup_history_combobox(self.entry_study_prog, "program_name", initial_study_prog, save_study_fields)
        self.setup_history_combobox(self.entry_study_group, "program_group_name", initial_study_group, save_study_fields)
        self.setup_history_combobox(self.entry_study_year, "academic_year_id", initial_study_year, save_study_fields)
        self.setup_history_combobox(self.entry_study_decree, "decree_number", initial_decree, save_study_fields)
        self.setup_history_combobox(self.entry_study_signing, "date_signing", initial_signing, save_study_fields)
        self.setup_history_combobox(self.entry_study_start, "date_start", initial_start, save_study_fields)

        # Блок сверки списка детей с отдельной таблицей Excel (2 столбца: ФИО и флаг зачисления)
        excel_verify_frame = ttk.LabelFrame(
            tab_study,
            text=" 📑 Сверка со списком детей из таблицы study_list.xlsx (2 столбца: ФИО и флаг) ",
            padding=(8, 4)
        )
        excel_verify_frame.grid(row=7, column=0, columnspan=2, sticky=tk.W + tk.E, pady=(3, 5))
        excel_verify_frame.columnconfigure(0, weight=1)

        ev_top = ttk.Frame(excel_verify_frame)
        ev_top.pack(fill=tk.X, expand=True)

        self.var_study_verify_excel = tk.BooleanVar(
            value=self.saved_cfg.get("study_verify_excel", True)
        )
        def on_toggle_verify():
            val = self.var_study_verify_excel.get()
            update_config(study_verify_excel=val)
            self.saved_cfg = load_config()
            st_text = "включена (пропуск лишних заявок)" if val else "отключена (обрабатывать все заявки с сервера)"
            self.log(f"⚙ Сверка с таблицей study_list.xlsx {st_text}", "cyan")

        chk_verify = ttk.Checkbutton(
            ev_top,
            text="Сверять Фамилию и Имя со study_list.xlsx (отчество игнорируется, инициалы отмечаются как ошибка)",
            variable=self.var_study_verify_excel,
            command=on_toggle_verify
        )
        chk_verify.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ev_btn_box = ttk.Frame(ev_top)
        ev_btn_box.pack(side=tk.RIGHT)

        ttk.Button(
            ev_btn_box,
            text="📑 Открыть study_list.xlsx",
            command=lambda: open_file_in_os(STUDY_EXCEL_FILE)
        ).pack(side=tk.LEFT, padx=(4, 2))

        ttk.Button(
            ev_btn_box,
            text="🔄 Обновить",
            command=self.reload_study_excel_info
        ).pack(side=tk.LEFT, padx=2)

        self.lbl_study_excel_status = ttk.Label(
            excel_verify_frame,
            text="Таблица study_list.xlsx: 2 столбца (ФИО и флаг). Сравнение по Фамилии и Имени активно.",
            foreground="#2563eb",
            font=("Segoe UI", 8, "italic")
        )
        self.lbl_study_excel_status.pack(anchor=tk.W, pady=(2, 0))

        self.btn_check_study_orders = ttk.Button(
            tab_study,
            text="🔍 Найти и проверить заявки программы (GET /api/rest/order)",
            command=self.check_study_program_orders
        )
        self.btn_check_study_orders.grid(row=8, column=0, columnspan=2, sticky=tk.W + tk.E, pady=(5, 4))

        # Блок с кнопками действий: 1. Подтвердить, 2. Отметить обучение, 3. Зачислить всё
        study_actions = ttk.Frame(tab_study)
        study_actions.grid(row=9, column=0, columnspan=2, sticky=tk.W + tk.E, pady=2)
        study_actions.columnconfigure(0, weight=1)
        study_actions.columnconfigure(1, weight=1)
        study_actions.columnconfigure(2, weight=1)

        self.btn_run_study_approve = ttk.Button(
            study_actions,
            text="✓ 1. Подтвердить заявки (initial ➜ approve)",
            command=self.run_study_batch_approve,
            state=tk.DISABLED
        )
        self.btn_run_study_approve.grid(row=0, column=0, sticky=tk.W + tk.E, padx=(0, 3), pady=2)

        self.btn_run_study_mark = ttk.Button(
            study_actions,
            text="🎓 2. Отметка об обучении (approve ➜ study)",
            command=self.run_study_batch_study,
            state=tk.DISABLED
        )
        self.btn_run_study_mark.grid(row=0, column=1, sticky=tk.W + tk.E, padx=3, pady=2)

        self.btn_run_study_all = ttk.Button(
            study_actions,
            text="⚡ 3. Выполнить оба шага (Подтвердить + Обучение)",
            command=self.run_study_batch_all,
            state=tk.DISABLED
        )
        self.btn_run_study_all.grid(row=0, column=2, sticky=tk.W + tk.E, padx=(3, 0), pady=2)

        ttk.Label(
            tab_study,
            text="💡 Для подтверждения заявка должна быть initial, для отметки об обучении — approve. Кнопка 3 выполняет оба шага подряд.",
            foreground="#64748b",
            font=("Segoe UI", 8, "italic")
        ).grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(3, 2))

        # ---------------------------------------------------------------------
        # БОЛЬШАЯ КРАСНАЯ КНОПКА ЭКСТРЕННОГО ПРЕРЫВАНИЯ ТЕКУЩЕЙ ЗАДАЧИ
        # ---------------------------------------------------------------------
        abort_frame = ttk.Frame(self.main_frame)
        abort_frame.pack(fill=tk.X, pady=(4, 8))

        self.btn_abort_task = tk.Button(
            abort_frame,
            text="🛑 ПРЕРВАТЬ ТЕКУЩУЮ ЗАДАЧУ И ВЕРНУТЬСЯ В ИСХОДНОЕ СОСТОЯНИЕ",
            font=("Segoe UI", 11, "bold"),
            bg="#dc2626",
            fg="#ffffff",
            activebackground="#991b1b",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=3,
            cursor="hand2",
            command=self.abort_current_task,
            pady=8
        )
        self.btn_abort_task.pack(fill=tk.X, expand=True)

        # Консоль журнала (Лог работы)
        log_frame = ttk.LabelFrame(self.main_frame, text="Журнал операций", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 10), bg="#0f172a", fg="#f8fafc", insertbackground="white")
        scroll = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)

        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Теги стилей для лога
        self.txt_log.tag_config("green", foreground="#4ade80")
        self.txt_log.tag_config("red", foreground="#f87171")
        self.txt_log.tag_config("yellow", foreground="#fbbf24")
        self.txt_log.tag_config("cyan", foreground="#38bdf8")

        self.log("Программа готова к работе.", "cyan")
        self.log("Заполните все поля, заполните Excel таблицу данными детей.")
        self.log("Если не получается найти мероприятия или программу - попробуйте перезапусить приложение.")
        self.log("Удачи!")

    def log(self, text: str, tag: Optional[str] = None):
        def _append():
            self.txt_log.insert(tk.END, text + "\n", tag)
            self.txt_log.see(tk.END)
        self.root.after(0, _append)

        # Сквозное автоматическое сохранение всех сообщений всех вкладок в файл текущей сессии
        try:
            write_to_log_file(text, with_timestamp=True)
        except Exception:
            pass

    def set_task_running(self, running: bool, task_name: str = ""):
        """Переключает статус выполнения задачи и обновляет вид большой красной кнопки."""
        prev_name = getattr(self, "current_task_name", "")
        self.is_task_running = running
        self.current_task_name = task_name if running else ""
        if running:
            self.stop_requested = False
            write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ▶️ СТАРТ ВЫПОЛНЕНИЯ ЗАДАЧИ: {task_name}")
            if hasattr(self, "btn_abort_task") and self.btn_abort_task:
                self.btn_abort_task.config(
                    text=f"⛔ ПРЕРВАТЬ ЗАДАЧУ: {task_name.upper()} (ВЫПОЛНЯЕТСЯ...) — НАЖМИТЕ ДЛЯ СТОПА",
                    bg="#b91c1c",
                    fg="#ffffff"
                )
        else:
            finished_title = task_name or prev_name
            if finished_title:
                write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ⏹️ ЗАВЕРШЕНИЕ ВЫПОЛНЕНИЯ ЗАДАЧИ: {finished_title}")
            self.stop_requested = False
            if hasattr(self, "btn_abort_task") and self.btn_abort_task:
                self.btn_abort_task.config(
                    text="🛑 ПРЕРВАТЬ ТЕКУЩУЮ ЗАДАЧУ И ВЕРНУТЬСЯ В ИСХОДНОЕ СОСТОЯНИЕ",
                    bg="#dc2626",
                    fg="#ffffff"
                )

    def abort_current_task(self):
        """
        Прерывает любую текущую задачу приложения (пакетную запись, подтверждение, зачисление)
        и немедленно возвращает приложение в исходное рабочее состояние.
        """
        self.stop_requested = True
        is_running = getattr(self, "is_task_running", False)
        task_desc = getattr(self, "current_task_name", "")

        self.log("\n" + "!" * 72, "red")
        if is_running and task_desc:
            self.log(f"🛑 [ПРЕРВАНО ПОЛЬЗОВАТЕЛЕМ] Нажата кнопка прерывания задачи '{task_desc}'!", "red")
            self.log("⏳ Немедленная остановка запросов и возврат интерфейса в исходное состояние...", "yellow")
        else:
            self.log("🛑 [СБРОС] Принудительный возврат интерфейса приложения в исходное состояние...", "yellow")
        self.log("!" * 72 + "\n", "red")

        write_to_log_file("\n" + "!" * 80)
        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🛑 ПОЛЬЗОВАТЕЛЬ НАЖАЛ БОЛЬШУЮ КРАСНУЮ КНОПКУ ПРЕРЫВАНИЯ ЗАДАЧИ!")
        write_to_log_file("!" * 80 + "\n")

        if hasattr(self, "btn_abort_task") and self.btn_abort_task:
            try:
                self.btn_abort_task.config(
                    text="⏳ ОСТАНОВКА... ВОЗВРАТ В ИСХОДНОЕ СОСТОЯНИЕ",
                    bg="#7f1d1d"
                )
            except Exception:
                pass

        self.root.after(100, self.reset_to_initial_state)

    def reset_to_initial_state(self):
        """Восстанавливает все кнопки, поля и переключатели приложения в исходное рабочее состояние."""
        self.is_task_running = False
        self.stop_requested = False
        self.current_task_name = ""

        # Восстановление кнопок вкладки 1 (Запись на мероприятие)
        if hasattr(self, "btn_run_add") and self.btn_run_add:
            try:
                self.btn_run_add.config(state=tk.NORMAL, text="▶ Запустить запись детей на мероприятие из таблицы")
            except Exception:
                pass
        if hasattr(self, "btn_check_add_event") and self.btn_check_add_event:
            try:
                self.btn_check_add_event.config(state=tk.NORMAL)
            except Exception:
                pass

        # Восстановление кнопок вкладки 2 (Подтверждение и отметка участия)
        if hasattr(self, "btn_check_orders") and self.btn_check_orders:
            try:
                self.btn_check_orders.config(state=tk.NORMAL, text="🔍 Проверить статус заявок мероприятия")
            except Exception:
                pass
        if hasattr(self, "btn_run_confirm") and self.btn_run_confirm:
            try:
                self.btn_run_confirm.config(state=tk.NORMAL, text="✓ 1. Подтвердить новые заявки")
            except Exception:
                pass
        if hasattr(self, "btn_run_participant") and self.btn_run_participant:
            try:
                self.btn_run_participant.config(state=tk.NORMAL, text="🎖️ 2. Отметить участие")
            except Exception:
                pass
        if hasattr(self, "btn_check_confirm_event") and self.btn_check_confirm_event:
            try:
                self.btn_check_confirm_event.config(state=tk.NORMAL)
            except Exception:
                pass

        # Восстановление кнопок вкладки 3 (Зачисление на программу)
        if hasattr(self, "btn_run_program") and self.btn_run_program:
            try:
                self.btn_run_program.config(state=tk.NORMAL, text="▶ Запустить зачисление детей на программу из таблицы")
            except Exception:
                pass
        if hasattr(self, "btn_check_prog") and self.btn_check_prog:
            try:
                self.btn_check_prog.config(state=tk.NORMAL)
            except Exception:
                pass
        if hasattr(self, "btn_check_group") and self.btn_check_group:
            try:
                self.btn_check_group.config(state=tk.NORMAL)
            except Exception:
                pass

        # Восстановление кнопок вкладки 4 (Подтверждение и обучение)
        if hasattr(self, "btn_check_study_orders") and self.btn_check_study_orders:
            try:
                self.btn_check_study_orders.config(state=tk.NORMAL, text="🔍 Найти и проверить заявки программы (GET /api/rest/order)")
            except Exception:
                pass
        if hasattr(self, "btn_run_study_approve") and self.btn_run_study_approve:
            try:
                self.btn_run_study_approve.config(state=tk.NORMAL, text="✓ 1. Подтвердить заявки (initial ➜ approve)")
            except Exception:
                pass
        if hasattr(self, "btn_run_study_mark") and self.btn_run_study_mark:
            try:
                self.btn_run_study_mark.config(state=tk.NORMAL, text="🎓 2. Отметка об обучении (approve ➜ study)")
            except Exception:
                pass
        if hasattr(self, "btn_run_study_all") and self.btn_run_study_all:
            try:
                self.btn_run_study_all.config(state=tk.NORMAL, text="⚡ 3. Выполнить оба шага (Подтвердить + Обучение)")
            except Exception:
                pass
        if hasattr(self, "btn_check_study_prog") and self.btn_check_study_prog:
            try:
                self.btn_check_study_prog.config(state=tk.NORMAL)
            except Exception:
                pass
        if hasattr(self, "btn_check_study_group") and self.btn_check_study_group:
            try:
                self.btn_check_study_group.config(state=tk.NORMAL)
            except Exception:
                pass

        # Восстановление монитора сервера
        if hasattr(self, "btn_ping_now") and self.btn_ping_now:
            try:
                self.btn_ping_now.config(state=tk.NORMAL, text="🔄 Проверить отклик")
            except Exception:
                pass

        # Восстановление большой красной кнопки
        if hasattr(self, "btn_abort_task") and self.btn_abort_task:
            try:
                self.btn_abort_task.config(
                    text="🛑 ПРЕРВАТЬ ТЕКУЩУЮ ЗАДАЧУ И ВЕРНУТЬСЯ В ИСХОДНОЕ СОСТОЯНИЕ",
                    bg="#dc2626",
                    fg="#ffffff"
                )
            except Exception:
                pass

        self.log("✓ Приложение находится в исходном состоянии, готово к выполнению команд.", "green")

    # ---------------------------------------------------------------------
    # ОБРАБОТЧИКИ ОПЕРАЦИЙ
    # ---------------------------------------------------------------------
    def run_batch_add(self):
        event_name = self.entry_add_event.get().strip()
        event_dt = self.entry_add_dt.get().strip()

        if not event_name or not event_dt:
            messagebox.showwarning("Внимание", "Заполните название мероприятия и дату/время.")
            return

        target_excel = EVENT_EXCEL_FILE if os.path.exists(EVENT_EXCEL_FILE) else (LEGACY_EXCEL_FILE if os.path.exists(LEGACY_EXCEL_FILE) else EVENT_EXCEL_FILE)
        if not os.path.exists(target_excel):
            ensure_default_excel(EVENT_EXCEL_FILE)
            target_excel = EVENT_EXCEL_FILE

        # Сохраняем в config.json при каждом запуске
        update_config(activity_name=event_name, activity_datetime=event_dt)
        self.saved_cfg = load_config()

        self.set_task_running(True, "Запись на мероприятие")
        self.btn_run_add.config(state=tk.DISABLED)
        self.log("\n" + "=" * 60, "cyan")
        self.log(f"[СТАРТ] Поиск мероприятия: '{event_name}'...", "cyan")

        def worker():
            try:
                if getattr(self, "stop_requested", False):
                    return

                act, msg = self.client.search_activity(event_name)
                if getattr(self, "stop_requested", False):
                    return
                if not act:
                    self.log(f"✗ Мероприятие не найдено: {msg}", "red")
                    self.root.after(0, lambda: self.btn_run_add.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.set_task_running(False))
                    return

                act_id = int(act["id"])
                self.log(f"✓ Найдено: '{act.get('name')}' (ID: {act_id})", "green")
                self.log(f"Дата и время заявки: {event_dt}")

                try:
                    records = self.client.read_excel(target_excel)
                except Exception as e:
                    self.log(f"✗ Ошибка чтения Excel: {e}", "red")
                    self.root.after(0, lambda: self.btn_run_add.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.set_task_running(False))
                    return

                self.log(f"Загружено записей из таблицы ({os.path.basename(target_excel)}): {len(records)}\n", "cyan")

                # Записываем начало сессии в файл results_log.txt
                start_time_str = time.strftime("%Y-%m-%d %H:%M:%S")
                write_to_log_file("=" * 80)
                write_to_log_file(f"[{start_time_str}] ПАКЕТНАЯ ЗАПИСЬ ДЕТЕЙ ИЗ ТАБЛИЦЫ")
                write_to_log_file(f"Мероприятие:         {act.get('name')} (ID: {act_id})")
                write_to_log_file(f"Дата и время заявки: {event_dt}")
                write_to_log_file(f"Файл таблицы:        {os.path.abspath(target_excel)} (строк: {len(records)})")
                write_to_log_file("-" * 80)

                stats = {"ok": 0, "skip_already": 0, "skip_unapproved": 0, "skip_dob": 0, "skip_notfound": 0, "err": 0}

                for item in records:
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Пакетная запись прервана пользователем!", "red")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🛑 ПАКЕТНАЯ ЗАПИСЬ ПРЕРВАНА ПОЛЬЗОВАТЕЛЕМ")
                        break

                    r_num = item["row"]
                    fio = item["fio"]
                    raw_dob = item["dob"]
                    norm_dob = self.client._normalize_dob(raw_dob)
                    curr_status = item.get("status", "")

                    # 1. Защита от дублирования заявок: если ребенок уже был успешно добавлен ранее
                    if item.get("is_processed"):
                        stats["skip_already"] += 1
                        self.log(f"[{r_num}] {fio} — ⏭️ ПРОПУЩЕНО: заявка уже обработана ранее ({curr_status})", "gray")
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ПРОПУЩЕНО: УЖЕ ОБРАБОТАН] Флаг в Excel: '{curr_status}'\n"
                        )
                        continue

                    # 2. Защита от инициалов / некорректного формата ФИО в Excel
                    if not item.get("is_valid_format", True):
                        fmt_err = item.get("format_error", "неправильный формат")
                        stats["err"] += 1
                        self.log(f"[{r_num}] {fio} — ⚠️ ПРОПУЩЕНО: {fmt_err} (укажите Фамилию и полное Имя)", "yellow")
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ОШИБКА ФОРМАТА] {fmt_err} | В Excel отмечено: '{curr_status}'\n"
                        )
                        continue

                    self.log(f"[{r_num}] {fio} (ДР: {raw_dob} -> {norm_dob or 'нет'})")

                    kid, search_msg = self.client.find_kid(fio, norm_dob)
                    if getattr(self, "stop_requested", False):
                        break

                    if not kid:
                        if "is_approved" in search_msg.lower():
                            reason = "НЕ ПОДТВЕРЖДЕН РОДИТЕЛЕМ ('is_approved': false)"
                            self.log(f"   ✗ ПРОПУСК: Аккаунт ребенка НЕ подтвержден ('is_approved': false)", "yellow")
                            stats["skip_unapproved"] += 1
                        elif "тезок" in search_msg.lower() or "не совпала" in search_msg.lower():
                            reason = "НЕ СОВПАЛА ДАТА РОЖДЕНИЯ (среди найденных тезок)"
                            self.log(f"   ✗ ПРОПУСК: Не совпала дата рождения среди тезок", "yellow")
                            stats["skip_dob"] += 1
                        else:
                            reason = f"НЕ НАЙДЕН В НАВИГАТОРЕ ({search_msg})"
                            self.log(f"   ✗ ПРОПУСК: {search_msg}", "red")
                            stats["skip_notfound"] += 1

                        # Фиксируем отказ в results_log.txt
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ПРОПУЩЕНО] Причина: {reason}\n"
                        )
                        continue

                    # Ребенок найден и подтвержден
                    ok, ord_msg, oid = self.client.create_order(act_id, event_dt, kid)
                    if getattr(self, "stop_requested", False):
                        break

                    if ok:
                        self.log(f"   ✓ УСПЕШНО: {ord_msg} (Kid ID: {kid.get('id')})", "green")
                        stats["ok"] += 1

                        # Автоматически переключаем флаг в 3-м столбце Excel, чтобы не допустить дублирования
                        order_tag = f"Заявка #{oid}" if oid else "Заявка создана"
                        excel_flag = f"Добавлен ({order_tag})"
                        saved, save_msg = self.client.update_child_status(target_excel, r_num, excel_flag, item.get("status_col", 3))
                        if saved:
                            self.log(f"   💾 Статус в Excel сохранен: '{excel_flag}' (строка {r_num})", "cyan")
                        else:
                            self.log(f"   ⚠️ Не удалось обновить статус в Excel: {save_msg}", "yellow")

                        # Фиксируем успех в results_log.txt
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [УСПЕШНО] Заявка #{oid or 'OK'} создана (Kid ID: {kid.get('id')}) | Excel: {excel_flag}\n"
                        )
                    else:
                        self.log(f"   ✗ ОШИБКА СОЗДАНИЯ ЗАЯВКИ: {ord_msg}", "red")
                        stats["err"] += 1
                        # Фиксируем ошибку отправки в results_log.txt
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ОШИБКА] {ord_msg} (Kid ID: {kid.get('id')})\n"
                        )

                # Записываем итоги в results_log.txt
                write_to_log_file("-" * 80)
                write_to_log_file("ИТОГИ ПАКЕТНОЙ ЗАПИСИ:")
                write_to_log_file(f"  Всего записей в таблице:    {len(records)}")
                write_to_log_file(f"  Успешно создано заявок:     {stats['ok']}")
                write_to_log_file(f"  Пропущено (ранее добавлены):{stats['skip_already']}")
                write_to_log_file(f"  Пропущено (не подтвержден): {stats['skip_unapproved']}")
                write_to_log_file(f"  Пропущено (ДР не совпала):  {stats['skip_dob']}")
                write_to_log_file(f"  Пропущено (не найден):      {stats['skip_notfound']}")
                write_to_log_file(f"  Ошибок отправки:            {stats['err']}")
                write_to_log_file("=" * 80 + "\n\n")

                self.log("\n" + "-" * 50, "cyan")
                self.log("ИТОГИ ПАКЕТНОЙ ЗАПИСИ:", "cyan")
                self.log(f"Всего в таблице:            {len(records)}")
                self.log(f"Успешно создано заявок:     {stats['ok']}", "green")
                if stats["skip_already"]:
                    self.log(f"Пропущено (ранее добавлены):{stats['skip_already']}", "cyan")
                else:
                    self.log(f"Пропущено (ранее добавлены):0")
                self.log(f"Пропущено (не подтвержден): {stats['skip_unapproved']}", "yellow")
                self.log(f"Пропущено (ДР не совпала):  {stats['skip_dob']}", "yellow")
                self.log(f"Пропущено (не найден):      {stats['skip_notfound']}", "red")
                self.log(f"Ошибок отправки:            {stats['err']}", "red" if stats["err"] else None)
                self.log(f"📄 Подробный отчет сохранен в: {LOG_FILE}", "cyan")
                self.log("-" * 50 + "\n", "cyan")

            except InterruptedError:
                self.log("\n🛑 [СТОП] Пакетная запись немедленно прервана!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    self.btn_run_add.config(state=tk.NORMAL)
                    if not getattr(self, "stop_requested", False):
                        messagebox.showinfo(
                            "Завершено",
                            f"Обработка завершена!\n\n"
                            f"✓ Успешно создано заявок: {stats['ok']}\n"
                            f"⏭️ Пропущено (ранее добавлены): {stats['skip_already']}\n"
                            f"✗ Пропущено других: {stats['skip_unapproved'] + stats['skip_dob'] + stats['skip_notfound']}\n"
                            f"❗ Ошибок: {stats['err']}"
                        )
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def check_add_event_info(self):
        """Проверяет наличие мероприятия для вкладки Записи на мероприятие."""
        event_name = self.entry_add_event.get().strip()
        if not event_name:
            messagebox.showwarning("Внимание", "Введите название мероприятия для проверки.")
            return

        add_to_field_history("activity_name", event_name)
        update_config(activity_name=event_name)
        self.saved_cfg = load_config()

        self.log(f"\n[ПОИСК] Поиск мероприятия (запись): '{event_name}'...", "cyan")

        def worker():
            act, msg = self.client.search_activity(event_name)
            if act:
                aid = act.get("id")
                aname = act.get("name")
                self.log(f"✓ Мероприятие найдено: '{aname}' (ID: {aid})", "green")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Мероприятие найдено",
                    f"✓ Мероприятие найдено в Навигаторе!\n\n"
                    f"Название: {aname}\n"
                    f"ID: {aid}"
                ))
            else:
                self.log(f"✗ {msg}", "red")
                self.root.after(0, lambda: messagebox.showerror(
                    "Не найдено",
                    f"Мероприятие '{event_name}' не найдено в Навигаторе.\n\n"
                    f"Проверьте правильность написания названия."
                ))

        threading.Thread(target=worker, daemon=True).start()

    def check_conf_event_info(self):
        """Проверяет наличие мероприятия для вкладки Подтверждения."""
        event_name = self.entry_conf_event.get().strip()
        if not event_name:
            messagebox.showwarning("Внимание", "Введите название мероприятия для проверки.")
            return

        add_to_field_history("activity_name", event_name)
        update_config(confirm_activity_name=event_name)
        self.saved_cfg = load_config()

        self.log(f"\n[ПОИСК] Поиск мероприятия (подтверждение): '{event_name}'...", "cyan")

        def worker():
            act, msg = self.client.search_activity(event_name)
            if act:
                aid = act.get("id")
                aname = act.get("name")
                self.log(f"✓ Мероприятие найдено: '{aname}' (ID: {aid})", "green")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Мероприятие найдено",
                    f"✓ Мероприятие найдено в Навигаторе!\n\n"
                    f"Название: {aname}\n"
                    f"ID: {aid}"
                ))
            else:
                self.log(f"✗ {msg}", "red")
                self.root.after(0, lambda: messagebox.showerror(
                    "Не найдено",
                    f"Мероприятие '{event_name}' не найдено в Навигаторе.\n\n"
                    f"Проверьте правильность написания названия."
                ))

        threading.Thread(target=worker, daemon=True).start()

    def check_pending_orders(self, silent: bool = False):
        event_name = self.entry_conf_event.get().strip()
        if not event_name:
            if not silent:
                messagebox.showwarning("Внимание", "Введите название мероприятия.")
            return

        # Сохраняем в config.json
        update_config(confirm_activity_name=event_name)
        self.saved_cfg = load_config()

        self.btn_check_orders.config(state=tk.DISABLED)
        if not silent:
            self.log(f"\n[ПОИСК] Мероприятие: '{event_name}'...", "cyan")

        def worker():
            act, msg = self.client.search_activity(event_name)
            if not act:
                self.log(f"✗ Мероприятие не найдено: {msg}", "red")
                self.root.after(0, lambda: self.btn_check_orders.config(state=tk.NORMAL))
                return

            act_id = int(act["id"])
            self.current_confirm_act_id = act_id
            orders_initial, total_initial = self.client.get_pending_orders(act_id, limit=100)
            orders_approved, total_approved = self.client.get_approved_orders(act_id, limit=100)
            orders_part, total_part = self.client.get_activity_orders_by_state(act_id, ["participant"], limit=1)

            self.log(f"✓ Мероприятие: '{act.get('name')}' (ID: {act_id})", "green")
            self.log(f"  • Неподтвержденных (initial): {total_initial}", "yellow" if total_initial > 0 else "gray")
            self.log(f"  • Подтвержденных без отметки участия (approve): {total_approved}", "cyan" if total_approved > 0 else "gray")
            self.log(f"  • С уже отмеченным участием (participant): {total_part}", "green" if total_part > 0 else "gray")

            def ui_update():
                self.btn_check_orders.config(state=tk.NORMAL)
                if total_initial > 0:
                    self.btn_run_confirm.config(state=tk.NORMAL, text=f"✓ 1. Подтвердить {total_initial} заявок")
                else:
                    self.btn_run_confirm.config(state=tk.DISABLED, text="✓ 1. Нет новых заявок")

                if total_approved > 0:
                    self.btn_run_participant.config(state=tk.NORMAL, text=f"🎖️ 2. Отметить участие ({total_approved} шт)")
                else:
                    self.btn_run_participant.config(state=tk.DISABLED, text="🎖️ 2. Нет заявок для отметки")

                if not silent and total_initial == 0 and total_approved == 0:
                    messagebox.showinfo(
                        "Статус заявок",
                        f"По мероприятию '{act.get('name')}':\n\n"
                        f"• Новых неподтвержденных: 0\n"
                        f"• Подтвержденных без участия: 0\n"
                        f"• Уже отмечено участие (participant): {total_part}"
                    )

            self.root.after(0, ui_update)

        threading.Thread(target=worker, daemon=True).start()

    def run_batch_confirm(self):
        event_name = self.entry_conf_event.get().strip()
        if not event_name:
            messagebox.showwarning("Внимание", "Введите название мероприятия.")
            return

        if not hasattr(self, "current_confirm_act_id") or not self.current_confirm_act_id:
            self.check_pending_orders()
            return

        act_id = self.current_confirm_act_id
        if not messagebox.askyesno(
            "Подтверждение заявок",
            "Вы действительно хотите подтвердить все новые заявки (state -> approve) по этому мероприятию?"
        ):
            return

        self.set_task_running(True, "Подтверждение заявок мероприятия")
        self.btn_run_confirm.config(state=tk.DISABLED)
        self.btn_run_participant.config(state=tk.DISABLED)
        self.log("\n[СТАРТ] Массовое подтверждение заявок (state -> approve)...", "cyan")

        def worker():
            total_approved = 0
            total_err = 0
            try:
                while True:
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Подтверждение заявок прервано пользователем!", "red")
                        break

                    batch, remaining = self.client.get_pending_orders(act_id, limit=100)
                    if getattr(self, "stop_requested", False):
                        break
                    if not batch:
                        break
                    for o in batch:
                        if getattr(self, "stop_requested", False):
                            self.log("\n🛑 [СТОП] Подтверждение заявок прервано пользователем!", "red")
                            break
                        oid = o["id"]
                        child_fio = o.get("child_fio") or o.get("child_name") or ""
                        name_str = f" ({child_fio})" if child_fio else ""
                        ok, ord_msg = self.client.approve_order(oid)
                        if getattr(self, "stop_requested", False):
                            break
                        if ok:
                            total_approved += 1
                            self.log(f"  ✓ Заявка #{oid}{name_str}: подтверждена (state -> approve)", "green")
                        else:
                            total_err += 1
                            self.log(f"  ✗ Заявка #{oid}{name_str}: ошибка подтверждения ({ord_msg})", "red")

                self.log(f"\n[ГОТОВО] Всего успешно подтверждено: {total_approved} заявок!", "green")
                if total_err > 0:
                    self.log(f"Ошибок подтверждения: {total_err}", "red")

                write_to_log_file(
                    "=" * 80 + "\n"
                    f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] МАССОВОЕ ПОДТВЕРЖДЕНИЕ ЗАЯВОК (APPROVE)\n"
                    f"Мероприятие ID: {act_id}\n"
                    f"Успешно подтверждено заявок: {total_approved}\n"
                    f"Ошибок: {total_err}\n"
                    + "=" * 80 + "\n\n"
                )
            except InterruptedError:
                self.log("\n🛑 [СТОП] Подтверждение заявок немедленно прервано!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    messagebox.showinfo(
                        "Подтверждение завершено",
                        f"Успешно подтверждено заявок: {total_approved}!\n\n"
                        "Теперь эти заявки можно отметить как участников (кнопка «🎖️ 2. Отметить участие»)."
                    )
                    self.check_pending_orders(silent=True)
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def run_batch_participant(self):
        event_name = self.entry_conf_event.get().strip()
        if not event_name:
            messagebox.showwarning("Внимание", "Введите название мероприятия.")
            return

        if not hasattr(self, "current_confirm_act_id") or not self.current_confirm_act_id:
            self.check_pending_orders()
            return

        act_id = self.current_confirm_act_id
        if not messagebox.askyesno(
            "Отметка участия",
            "Вы действительно хотите отметить участие (state -> participant) для всех подтвержденных заявок по этому мероприятию?"
        ):
            return

        self.set_task_running(True, "Отметка участия в мероприятии")
        self.btn_run_confirm.config(state=tk.DISABLED)
        self.btn_run_participant.config(state=tk.DISABLED)
        self.log("\n[СТАРТ] Массовая отметка участия детей (state -> participant)...", "cyan")

        def worker():
            total_participant = 0
            total_err = 0
            try:
                while True:
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Отметка участия прервана пользователем!", "red")
                        break

                    batch, remaining = self.client.get_approved_orders(act_id, limit=100)
                    if getattr(self, "stop_requested", False):
                        break
                    if not batch:
                        break
                    for o in batch:
                        if getattr(self, "stop_requested", False):
                            self.log("\n🛑 [СТОП] Отметка участия прервана пользователем!", "red")
                            break
                        oid = o["id"]
                        child_fio = o.get("child_fio") or o.get("child_name") or ""
                        name_str = f" ({child_fio})" if child_fio else ""
                        ok, ord_msg = self.client.mark_participant_order(oid)
                        if getattr(self, "stop_requested", False):
                            break
                        if ok:
                            total_participant += 1
                            self.log(f"  ✓ Заявка #{oid}{name_str}: участие успешно отмечено (state -> participant)", "green")
                        else:
                            total_err += 1
                            self.log(f"  ✗ Заявка #{oid}{name_str}: ошибка отметки участия ({ord_msg})", "red")

                self.log(f"\n[ГОТОВО] Всего успешно отмечено участие для {total_participant} заявок!", "green")
                if total_err > 0:
                    self.log(f"Ошибок: {total_err}", "red")

                write_to_log_file(
                    "=" * 80 + "\n"
                    f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] МАССОВАЯ ОТМЕТКА УЧАСТИЯ (PARTICIPANT)\n"
                    f"Мероприятие ID: {act_id}\n"
                    f"Успешно отмечено участие: {total_participant} заявок\n"
                    f"Ошибок: {total_err}\n"
                    + "=" * 80 + "\n\n"
                )
            except InterruptedError:
                self.log("\n🛑 [СТОП] Отметка участия немедленно прервана!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    messagebox.showinfo(
                        "Отметка участия завершена",
                        f"Успешно отмечено участие для {total_participant} заявок!\n"
                        "(state -> participant)"
                    )
                    self.check_pending_orders(silent=True)
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def check_program_info(self):
        """Проверяет существование программы в Навигаторе по названию и отображает ее данные."""
        prog_name = self.entry_prog_name.get().strip() if hasattr(self, "entry_prog_name") else ""
        if not prog_name:
            messagebox.showwarning("Внимание", "Введите название программы для проверки.")
            return

        add_to_field_history("program_name", prog_name)
        update_config(program_name=prog_name)
        self.saved_cfg = load_config()

        self.log(f"\n[ПОИСК] Поиск программы по названию: '{prog_name}'...", "cyan")

        def worker():
            prog, msg = self.client.search_program(prog_name)
            if prog:
                pid = prog.get("id")
                pname = prog.get("name")
                partner = prog.get("partner_name", "н/д")
                places = prog.get("count_group_places", "н/д")
                self.log(f"✓ Программа найдена: '{pname}' (ID: {pid})", "green")
                self.log(f"   • Организация: {partner} | Мест в группах: {places}", "cyan")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Программа найдена",
                    f"✓ Программа найдена в Навигаторе!\n\n"
                    f"Название: {pname}\n"
                    f"ID программы (event_id): {pid}\n"
                    f"Организация: {partner}\n"
                    f"Мест в группах: {places}"
                ))
            else:
                self.log(f"✗ {msg}", "red")
                self.root.after(0, lambda: messagebox.showerror(
                    "Не найдено",
                    f"Программа '{prog_name}' не найдена в Навигаторе.\n\n"
                    f"Проверьте правильность написания названия программы."
                ))

        threading.Thread(target=worker, daemon=True).start()

    def check_group_info(self):
        """Проверяет группу в Навигаторе и её принадлежность к выбранной программе через GET /api/rest/eventGroups."""
        group_val = self.entry_prog_group.get().strip() if hasattr(self, "entry_prog_group") else ""
        if not group_val:
            messagebox.showwarning("Внимание", "Введите название или ID группы для проверки.")
            return

        add_to_field_history("program_group_name", group_val)
        prog_name = self.entry_prog_name.get().strip() if hasattr(self, "entry_prog_name") else ""
        self.log(f"\n[ПРОВЕРКА] Проверка группы: '{group_val}'...", "cyan")

        def worker():
            target_event_id = None
            target_prog_name = prog_name
            if prog_name:
                if prog_name.isdigit():
                    target_event_id = prog_name
                else:
                    self.log(f"[ПОИСК] Поиск программы: '{prog_name}'...", "cyan")
                    prog, _ = self.client.search_program(prog_name)
                    if prog:
                        target_event_id = str(prog.get("id"))
                        target_prog_name = prog.get("name", prog_name)
                        self.log(f"✓ Найдена программа: '{target_prog_name}' (ID: {target_event_id})", "green")

            if target_event_id:
                self.log(f"[ПРОВЕРКА] Запрос групп программы '{target_prog_name}' (ID: {target_event_id})...", "cyan")
                group_obj, all_matches, msg = self.client.validate_group_for_program(group_val, target_event_id, program_name=target_prog_name)
                if group_obj:
                    gid = group_obj.get("id")
                    gname = group_obj.get("name", group_val)
                    gteach = group_obj.get("teacher", "н/д")
                    gevent = group_obj.get("event_id", target_event_id)
                    gsize = group_obj.get("size", "н/д")
                    self.log(f"✓ Группа ПОДТВЕРЖДЕНА: «{gname}» (ID: {gid}) принадлежит программе (ID: {gevent})", "green")
                    self.log(f"   • Педагог: {gteach} | Мест в группе: {gsize}", "cyan")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Группа подтверждена",
                        f"✓ Группа принадлежит выбранной программе!\n\n"
                        f"Программа: {target_prog_name} (ID: {gevent})\n"
                        f"Группа: {gname}\n"
                        f"ID группы: {gid}\n"
                        f"Педагог: {gteach}\n"
                        f"Мест в группе: {gsize}"
                    ))
                elif all_matches:
                    self.log(f"⚠️ {msg}", "yellow")
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Проверка группы программы",
                        msg
                    ))
                else:
                    self.log(f"✗ {msg}", "red")
                    self.root.after(0, lambda: messagebox.showerror(
                        "Группа не принадлежит программе",
                        msg
                    ))
            else:
                # Если программа вообще не указана в поле
                self.log(f"[ПОИСК] Программа не указана, выполняется общий поиск группы '{group_val}'...", "cyan")
                group_obj, all_matches, msg = self.client.search_group(group_val)
                if group_obj:
                    gid = group_obj.get("id")
                    gname = group_obj.get("name", group_val)
                    gteach = group_obj.get("teacher", "н/д")
                    gevent = group_obj.get("event_id", "н/д")
                    self.log(f"✓ Найдена группа: '{gname}' (ID: {gid}, программа: {gevent})", "green")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Группа найдена",
                        f"Название: {gname}\nID группы: {gid}\nПедагог: {gteach}\nID программы: {gevent}\n\n(Укажите программу в поле выше для точной проверки принадлежности группы к программе)"
                    ))
                else:
                    self.log(f"✗ {msg}", "red")
                    self.root.after(0, lambda: messagebox.showerror("Не найдено", msg))

        threading.Thread(target=worker, daemon=True).start()

    def run_batch_program_enroll(self):
        prog_name = self.entry_prog_name.get().strip() if hasattr(self, "entry_prog_name") else ""
        group_input = self.entry_prog_group.get().strip() if hasattr(self, "entry_prog_group") else ""
        raw_year = self.entry_prog_year.get().strip() if hasattr(self, "entry_prog_year") else ""
        create_cert = self.var_prog_create_cert.get() if hasattr(self, "var_prog_create_cert") else True
        use_cert = self.var_prog_use_cert.get() if hasattr(self, "var_prog_use_cert") else False

        if not prog_name or not group_input or not raw_year:
            messagebox.showwarning(
                "Внимание",
                "Пожалуйста, заполните Название программы, Группу (название или ID) и Учебный год (например: 2026/2027 или 2026)."
            )
            return

        academic_year_id = normalize_academic_year_id(raw_year)
        if not academic_year_id:
            messagebox.showwarning(
                "Некорректный учебный год",
                f"Не удалось распознать учебный год из '{raw_year}'.\nВведите в формате '2026/2027' или '2026'."
            )
            return

        target_excel = PROGRAM_EXCEL_FILE if os.path.exists(PROGRAM_EXCEL_FILE) else (LEGACY_EXCEL_FILE if os.path.exists(LEGACY_EXCEL_FILE) else PROGRAM_EXCEL_FILE)
        if not os.path.exists(target_excel):
            ensure_default_excel(PROGRAM_EXCEL_FILE)
            target_excel = PROGRAM_EXCEL_FILE

        # Сохраняем в config.json
        update_config(
            program_name=prog_name,
            program_group_name=group_input,
            program_group_id=group_input if group_input.isdigit() else str(self.saved_cfg.get("program_group_id") or ""),
            program_academic_year_id=raw_year,
            program_create_certificate=create_cert,
            program_use_certificate=use_cert
        )
        self.saved_cfg = load_config()

        self.set_task_running(True, "Зачисление на программу")
        self.btn_run_program.config(state=tk.DISABLED)
        self.log("\n" + "=" * 60, "cyan")
        self.log(f"[СТАРТ] Поиск программы: '{prog_name}'...", "cyan")

        def worker():
            try:
                if getattr(self, "stop_requested", False):
                    return

                # Шаг 1: Ищем программу по названию в Навигаторе (GET /api/rest/events)
                prog, search_msg = self.client.search_program(prog_name)
                if getattr(self, "stop_requested", False):
                    return
                if not prog:
                    self.log(f"✗ Программа не найдена: {search_msg}", "red")
                    self.root.after(0, lambda: self.btn_run_program.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.set_task_running(False))
                    return

                event_id = str(prog["id"])
                full_prog_name = prog.get("name", prog_name)
                self.log(f"✓ Найдена программа: '{full_prog_name}' (ID: {event_id})", "green")

                # Шаг 2: Проверяем, что группа принадлежит этой программе (GET /api/rest/eventGroups?extFilters=...)
                self.log(f"[ПРОВЕРКА] Проверка принадлежности группы '{group_input}' к программе '{full_prog_name}' (ID: {event_id})...", "cyan")
                group_obj, available_groups, group_msg = self.client.validate_group_for_program(group_input, event_id, program_name=full_prog_name)
                if getattr(self, "stop_requested", False):
                    return
                if not group_obj:
                    self.log(f"✗ ОШИБКА ПРИНАДЛЕЖНОСТИ: {group_msg}", "red")
                    self.root.after(0, lambda: messagebox.showerror(
                        "Группа не принадлежит программе",
                        group_msg
                    ))
                    self.root.after(0, lambda: self.btn_run_program.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.set_task_running(False))
                    return

                group_id = str(group_obj["id"])
                group_display_name = group_obj.get("name", group_input)
                teacher_name = group_obj.get("teacher", "н/д")
                self.log(f"✓ Группа подтверждена для программы: '{group_display_name}' (ID: {group_id}, педагог: {teacher_name})", "green")
                self.log(f"Параметры: ID Программы={event_id}, ID Группы={group_id}, Год={academic_year_id} (введено '{raw_year}'), Сертификат={create_cert}")

                try:
                    records = self.client.read_excel(target_excel)
                except Exception as e:
                    self.log(f"✗ Ошибка чтения Excel ({os.path.basename(target_excel)}): {e}", "red")
                    self.root.after(0, lambda: self.btn_run_program.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.set_task_running(False))
                    return

                self.log(f"Загружено записей из таблицы ({os.path.basename(target_excel)}): {len(records)}\n", "cyan")

                start_time_str = time.strftime("%Y-%m-%d %H:%M:%S")
                write_to_log_file("=" * 80)
                write_to_log_file(f"[{start_time_str}] ЗАЧИСЛЕНИЕ НА УЧЕБНУЮ ПРОГРАММУ (/api/rest/order)")
                write_to_log_file(f"Программа:                  {full_prog_name} (ID: {event_id})")
                write_to_log_file(f"Группа:                     {group_display_name} (ID: {group_id}, педагог: {teacher_name})")
                write_to_log_file(f"Учебный год (academic_year_id): {academic_year_id} (введено: '{raw_year}')")
                write_to_log_file(f"Создавать сертификат:       {create_cert}")
                write_to_log_file(f"Использовать сертификат:    {use_cert}")
                write_to_log_file(f"Файл таблицы:               {os.path.abspath(target_excel)} (строк: {len(records)})")
                write_to_log_file("-" * 80)

                stats = {"ok": 0, "skip_already_excel": 0, "skip_already_api": 0, "skip_unapproved": 0, "skip_dob": 0, "skip_notfound": 0, "err": 0}

                for item in records:
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Зачисление на программу прервано пользователем!", "red")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🛑 ЗАЧИСЛЕНИЕ ПРЕРВАНО ПОЛЬЗОВАТЕЛЕМ")
                        break

                    r_num = item["row"]
                    fio = item["fio"]
                    raw_dob = item["dob"]
                    norm_dob = self.client._normalize_dob(raw_dob)
                    curr_status = item.get("status", "")

                    # 1. Защита от повторного зачисления: если ребенок уже отмечен как обработанный в Excel
                    if item.get("is_processed"):
                        stats["skip_already_excel"] += 1
                        self.log(f"[{r_num}] {fio} — ⏭️ ПРОПУЩЕНО: заявка уже обработана в таблице ({curr_status})", "gray")
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ПРОПУЩЕНО: УЖЕ В ТАБЛИЦЕ] Флаг в Excel: '{curr_status}'\n"
                        )
                        continue

                    # 2. Защита от инициалов / некорректного формата ФИО в Excel
                    if not item.get("is_valid_format", True):
                        fmt_err = item.get("format_error", "неправильный формат")
                        stats["err"] += 1
                        self.log(f"[{r_num}] {fio} — ⚠️ ПРОПУЩЕНО: {fmt_err} (укажите Фамилию и полное Имя)", "yellow")
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ОШИБКА ФОРМАТА] {fmt_err} | В Excel отмечено: '{curr_status}'\n"
                        )
                        continue

                    self.log(f"[{r_num}] {fio} (ДР: {raw_dob} -> {norm_dob or 'нет'})")

                    kid, search_msg = self.client.find_kid(fio, norm_dob)
                    if getattr(self, "stop_requested", False):
                        break

                    if not kid:
                        if "is_approved" in search_msg.lower():
                            reason = "НЕ ПОДТВЕРЖДЕН РОДИТЕЛЕМ ('is_approved': false)"
                            self.log(f"   ✗ ПРОПУСК: Аккаунт ребенка НЕ подтвержден ('is_approved': false)", "yellow")
                            stats["skip_unapproved"] += 1
                        elif "тезок" in search_msg.lower() or "не совпала" in search_msg.lower():
                            reason = "НЕ СОВПАЛА ДАТА РОЖДЕНИЯ (среди найденных тезок)"
                            self.log(f"   ✗ ПРОПУСК: Не совпала дата рождения среди тезок", "yellow")
                            stats["skip_dob"] += 1
                        else:
                            reason = f"НЕ НАЙДЕН В НАВИГАТОРЕ ({search_msg})"
                            self.log(f"   ✗ ПРОПУСК: {search_msg}", "red")
                            stats["skip_notfound"] += 1

                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ПРОПУЩЕНО] Причина: {reason}\n"
                        )
                        continue

                    # Шаг 3: Защита от дубликатов: проверка через Навигатор (GET /api/rest/order)
                    # Проверяем, не подана ли уже заявка со статусом initial, approve или study
                    has_existing, exist_order, check_msg = self.client.check_existing_program_order(kid.get("id"), event_id)
                    if getattr(self, "stop_requested", False):
                        break

                    if has_existing:
                        oid = exist_order.get("id") if exist_order else "н/д"
                        st = exist_order.get("state", exist_order.get("state_grid", "активна")) if exist_order else "активна"
                        stats["skip_already_api"] += 1
                        self.log(f"   ⏭️ ПРОПУСК: Заявка уже зарегистрирована в Навигаторе (#{oid}, статус: '{st}')", "cyan")

                        # Фиксируем актуальный статус в programm_list.xlsx, чтобы исключить повторные запросы
                        excel_flag = f"Уже зачислен (Заявка #{oid})"
                        self.client.update_child_status(target_excel, r_num, excel_flag, item.get("status_col", 3))

                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ПРОПУЩЕНО: ДУБЛИКАТ В НАВИГАТОРЕ] Заявка #{oid} уже есть со статусом '{st}' (Kid ID: {kid.get('id')}) | Excel: {excel_flag}\n"
                        )
                        continue

                    # Ребенок найден, подтвержден и не имеет дублирующей заявки -> создание заявки на зачисление
                    ok, ord_msg, oid = self.client.create_program_order(
                        event_id=event_id,
                        group_id=group_id,
                        academic_year_id=academic_year_id,
                        kid=kid,
                        use_certificate=use_cert,
                        create_certificate=create_cert
                    )
                    if getattr(self, "stop_requested", False):
                        break

                    if ok:
                        self.log(f"   ✓ УСПЕШНО: {ord_msg} (Kid ID: {kid.get('id')}, Parent: {kid.get('site_user_id')})", "green")
                        stats["ok"] += 1

                        # Сохраняем отметку в programm_list.xlsx в 3-й столбец
                        order_tag = f"Заявка #{oid}" if oid else "Заявка создана"
                        excel_flag = f"Зачислен ({order_tag})"
                        saved, save_msg = self.client.update_child_status(target_excel, r_num, excel_flag, item.get("status_col", 3))
                        if saved:
                            self.log(f"   💾 Статус в Excel сохранен: '{excel_flag}' (строка {r_num})", "cyan")
                        else:
                            self.log(f"   ⚠️ Не удалось обновить статус в Excel: {save_msg}", "yellow")

                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [УСПЕШНО] Заявка #{oid or 'OK'} создана (Kid ID: {kid.get('id')}, Parent: {kid.get('site_user_id')}) | Excel: {excel_flag}\n"
                        )
                    else:
                        self.log(f"   ✗ ОШИБКА СОЗДАНИЯ ЗАЯВКИ: {ord_msg}", "red")
                        stats["err"] += 1
                        write_to_log_file(
                            f"Строка {r_num:02d} | {fio:<35} | ДР: {raw_dob} -> {norm_dob or 'н/д'}\n"
                            f"          -> СТАТУС: [ОШИБКА] {ord_msg} (Kid ID: {kid.get('id')})\n"
                        )

                # Записываем итоги в results_log.txt
                total_skipped_already = stats["skip_already_excel"] + stats["skip_already_api"]
                write_to_log_file("-" * 80)
                write_to_log_file("ИТОГИ ЗАЧИСЛЕНИЯ НА ПРОГРАММУ:")
                write_to_log_file(f"  Всего записей в таблице:         {len(records)}")
                write_to_log_file(f"  Успешно создано новых заявок:    {stats['ok']}")
                write_to_log_file(f"  Пропущено (были отмечены в Excel):{stats['skip_already_excel']}")
                write_to_log_file(f"  Пропущено (найдены в Навигаторе): {stats['skip_already_api']}")
                write_to_log_file(f"  Пропущено (не подтвержден):      {stats['skip_unapproved']}")
                write_to_log_file(f"  Пропущено (ДР не совпала):       {stats['skip_dob']}")
                write_to_log_file(f"  Пропущено (не найден):           {stats['skip_notfound']}")
                write_to_log_file(f"  Ошибок отправки:                 {stats['err']}")
                write_to_log_file("=" * 80 + "\n\n")

                self.log("\n" + "-" * 50, "cyan")
                self.log("ИТОГИ ЗАЧИСЛЕНИЯ НА ПРОГРАММУ:", "cyan")
                self.log(f"Всего в таблице:                 {len(records)}")
                self.log(f"Успешно создано новых заявок:    {stats['ok']}", "green")
                if stats["skip_already_excel"]:
                    self.log(f"Пропущено (отмечены в Excel):    {stats['skip_already_excel']}", "cyan")
                if stats["skip_already_api"]:
                    self.log(f"Пропущено (дубликат в Навигаторе): {stats['skip_already_api']}", "cyan")
                self.log(f"Пропущено (не подтвержден):      {stats['skip_unapproved']}", "yellow")
                self.log(f"Пропущено (ДР не совпала):       {stats['skip_dob']}", "yellow")
                self.log(f"Пропущено (не найден):           {stats['skip_notfound']}", "red")
                self.log(f"Ошибок отправки:                 {stats['err']}", "red" if stats["err"] else None)
                self.log(f"📄 Подробный отчет сохранен в: {LOG_FILE}", "cyan")
                self.log("-" * 50 + "\n", "cyan")

            except InterruptedError:
                self.log("\n🛑 [СТОП] Зачисление на программу немедленно прервано!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    self.btn_run_program.config(state=tk.NORMAL)
                    messagebox.showinfo(
                        "Зачисление завершено",
                        f"Обработка зачисления на программу завершена!\n\n"
                        f"✓ Успешно создано новых заявок: {stats['ok']}\n"
                        f"⏭️ Пропущено дубликатов: {total_skipped_already} (в Excel: {stats['skip_already_excel']}, в Навигаторе: {stats['skip_already_api']})\n"
                        f"✗ Пропущено других: {stats['skip_unapproved'] + stats['skip_dob'] + stats['skip_notfound']}\n"
                        f"❗ Ошибок: {stats['err']}"
                    )
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def check_study_program_info(self):
        """Проверяет программу в Навигаторе для вкладки подтверждения/обучения."""
        prog_name = self.entry_study_prog.get().strip() if hasattr(self, "entry_study_prog") else ""
        if not prog_name:
            messagebox.showwarning("Внимание", "Введите название программы для проверки.")
            return

        add_to_field_history("program_name", prog_name)
        update_config(study_program_name=prog_name)
        self.saved_cfg = load_config()

        self.log(f"\n[ПОИСК] Поиск программы: '{prog_name}'...", "cyan")

        def worker():
            prog, msg = self.client.search_program(prog_name)
            if prog:
                pid = prog.get("id")
                pname = prog.get("name")
                partner = prog.get("partner_name", "н/д")
                places = prog.get("count_group_places", "н/д")
                self.log(f"✓ Программа найдена: '{pname}' (ID: {pid})", "green")
                self.log(f"   • Организация: {partner} | Мест в группах: {places}", "cyan")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Программа найдена",
                    f"✓ Программа найдена в Навигаторе!\n\n"
                    f"Название: {pname}\n"
                    f"ID программы (event_id): {pid}\n"
                    f"Организация: {partner}\n"
                    f"Мест в группах: {places}"
                ))
            else:
                self.log(f"✗ {msg}", "red")
                self.root.after(0, lambda: messagebox.showerror(
                    "Не найдено",
                    f"Программа '{prog_name}' не найдена в Навигаторе.\n\n"
                    f"Проверьте правильность написания названия программы."
                ))

        threading.Thread(target=worker, daemon=True).start()

    def check_study_group_info(self):
        """Проверяет группу для вкладки подтверждения/обучения и её принадлежность к программе."""
        group_val = self.entry_study_group.get().strip() if hasattr(self, "entry_study_group") else ""
        prog_name = self.entry_study_prog.get().strip() if hasattr(self, "entry_study_prog") else ""

        if not group_val:
            messagebox.showinfo(
                "Все группы программы",
                "Поле группы оставлено ПУСТЫМ.\n\n"
                "При проверке и обработке заявок будут выбраны дети ИЗ ВСЕХ ГРУПП выбранной программы."
            )
            self.log("💡 Поле группы пустое: поиск заявок будет выполняться по ВСЕМ группам программы.", "cyan")
            return

        add_to_field_history("program_group_name", group_val)
        self.log(f"\n[ПРОВЕРКА] Проверка группы: '{group_val}'...", "cyan")

        def worker():
            target_event_id = None
            target_prog_name = prog_name
            if prog_name:
                if prog_name.isdigit():
                    target_event_id = prog_name
                else:
                    prog, _ = self.client.search_program(prog_name)
                    if prog:
                        target_event_id = str(prog.get("id"))
                        target_prog_name = prog.get("name", prog_name)

            if target_event_id:
                group_obj, all_matches, msg = self.client.validate_group_for_program(group_val, target_event_id, program_name=target_prog_name)
                if group_obj:
                    gid = group_obj.get("id")
                    gname = group_obj.get("name", group_val)
                    gteach = group_obj.get("teacher", "н/д")
                    gevent = group_obj.get("event_id", target_event_id)
                    gsize = group_obj.get("size", "н/д")
                    self.log(f"✓ Группа ПОДТВЕРЖДЕНА: «{gname}» (ID: {gid}) принадлежит программе (ID: {gevent})", "green")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Группа подтверждена",
                        f"✓ Группа принадлежит выбранной программе!\n\n"
                        f"Программа: {target_prog_name} (ID: {gevent})\n"
                        f"Группа: {gname}\n"
                        f"ID группы: {gid}\n"
                        f"Педагог: {gteach}\n"
                        f"Мест: {gsize}"
                    ))
                elif all_matches:
                    self.log(f"⚠️ {msg}", "yellow")
                    self.root.after(0, lambda: messagebox.showwarning("Проверка группы", msg))
                else:
                    self.log(f"✗ {msg}", "red")
                    self.root.after(0, lambda: messagebox.showerror("Группа не принадлежит программе", msg))
            else:
                group_obj, all_matches, msg = self.client.search_group(group_val)
                if group_obj:
                    gid = group_obj.get("id")
                    gname = group_obj.get("name", group_val)
                    self.log(f"✓ Найдена группа: '{gname}' (ID: {gid})", "green")
                    self.root.after(0, lambda: messagebox.showinfo("Группа найдена", f"Название: {gname}\nID: {gid}"))
                else:
                    self.log(f"✗ {msg}", "red")
                    self.root.after(0, lambda: messagebox.showerror("Не найдено", msg))

        threading.Thread(target=worker, daemon=True).start()

    def reload_study_excel_info(self):
        """Перечитывает отдельную таблицу study_list.xlsx и выводит информацию о детях, флагах и формате ФИО."""
        try:
            records = self.client.read_study_excel(STUDY_EXCEL_FILE)
            cnt = len(records)
            conf_cnt = sum(1 for r in records if r.get("is_confirmed"))
            enr_cnt = sum(1 for r in records if r.get("is_enrolled"))
            invalid_kids = [r for r in records if not r.get("is_valid_format", True)]

            msg = f"Таблица study_list.xlsx: {cnt} детей (подтверждено: {conf_cnt}, зачислено: {enr_cnt})."
            if invalid_kids:
                msg += f" ⚠️ Ошибок формата (инициалы): {len(invalid_kids)}!"

            if hasattr(self, "lbl_study_excel_status") and self.lbl_study_excel_status:
                self.lbl_study_excel_status.config(
                    text=msg,
                    foreground="#dc2626" if invalid_kids else ("#059669" if cnt else "#d97706")
                )
            self.log(f"📑 [ТАБЛИЦА] {msg}", "yellow" if invalid_kids else "cyan")

            if invalid_kids:
                self.log("⚠️ ВНИМАНИЕ: В таблице study_list.xlsx найдены записи с инициалами вместо полного имени:", "yellow")
                self.log("   (Сайт Навигатора выдает Фамилию и Имя — записи с инициалами отмечены как неправильный формат в файле):", "yellow")
                for r in invalid_kids:
                    self.log(f"   • Строка {r['row']}: '{r['fio']}' ➜ {r.get('format_error')} (отмечено в файле)", "yellow")
        except Exception as e:
            if hasattr(self, "lbl_study_excel_status") and self.lbl_study_excel_status:
                self.lbl_study_excel_status.config(text=f"Ошибка чтения study_list.xlsx: {e}", foreground="#dc2626")
            self.log(f"⚠️ Ошибка чтения study_list.xlsx: {e}", "red")

    def check_study_program_orders(self, silent: bool = False):
        """
        Запрашивает из Навигатора детей по названию программы и (опционально) группе:
        GET /api/rest/order?extFilters=[fact_academic_year_id, event_id, fact_group_id]
        Сверяет список детей по Фамилии и Имени с отдельной таблицей study_list.xlsx:
          - Сайт выдает Фамилию и Имя.
          - В Excel может быть Фамилия Имя Отчество (отчество игнорируется при сверке).
          - Если в Excel указаны Фамилия и инициалы, они отмечаются как неправильный формат.
        Защищает от ошибочных заявок родителей на другую программу.
        Подсчитывает и разделяет заявки:
          - initial -> ожидают подтверждения (/api/approveRequest)
          - approve -> ожидают отметки об обучении (/api/studyRequest)
          - study -> уже обучаются
        """
        prog_name = self.entry_study_prog.get().strip() if hasattr(self, "entry_study_prog") else ""
        group_input = self.entry_study_group.get().strip() if hasattr(self, "entry_study_group") else ""
        raw_year = self.entry_study_year.get().strip() if hasattr(self, "entry_study_year") else "2026"

        if not prog_name:
            if not silent:
                messagebox.showwarning("Внимание", "Укажите название программы.")
            return

        year_id = normalize_academic_year_id(raw_year)
        self.btn_check_study_orders.config(state=tk.DISABLED)

        if not silent:
            group_txt = f"группа: '{group_input}'" if group_input else "все группы"
            self.log(f"\n[ЗАПРОС ЗАЯВОК] Программа: '{prog_name}', {group_txt}, учебный год: {year_id}...", "cyan")

        def worker():
            # 1. Поиск программы
            prog_obj = None
            if prog_name.isdigit():
                prog_obj = {"id": prog_name, "name": f"Программа #{prog_name}"}
            else:
                prog_obj, p_msg = self.client.search_program(prog_name)
                if not prog_obj:
                    self.log(f"✗ Программа '{prog_name}' не найдена: {p_msg}", "red")
                    self.root.after(0, lambda: self.btn_check_study_orders.config(state=tk.NORMAL))
                    if not silent:
                        self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Программа '{prog_name}' не найдена."))
                    return

            event_id = str(prog_obj.get("id"))
            real_prog_name = prog_obj.get("name", prog_name)
            self.current_study_event_id = event_id
            self.current_study_prog_name = real_prog_name

            # 2. Проверка группы (если указана)
            target_group_id = None
            group_label = "Все группы программы"
            if group_input:
                g_obj, all_groups, v_msg = self.client.validate_group_for_program(group_input, event_id, program_name=real_prog_name)
                if not g_obj:
                    self.log(f"✗ Ошибка проверки группы:\n{v_msg}", "red")
                    self.root.after(0, lambda: self.btn_check_study_orders.config(state=tk.NORMAL))
                    if not silent:
                        self.root.after(0, lambda: messagebox.showerror("Группа не принадлежит программе", v_msg))
                    return
                target_group_id = str(g_obj.get("id"))
                group_label = f"Группа «{g_obj.get('name')}» (ID: {target_group_id})"
            
            self.current_study_group_id = target_group_id

            # 3. Запрос всех заявок через GET /api/rest/order
            orders, err = self.client.get_all_program_orders(
                event_id=event_id,
                academic_year_id=year_id,
                group_id=target_group_id
            )
            if err:
                self.log(f"✗ Ошибка получения заявок: {err}", "red")
                self.root.after(0, lambda: self.btn_check_study_orders.config(state=tk.NORMAL))
                return

            # 4. Сверка списка детей с таблицей study_list.xlsx по Фамилии и Имени
            verify_active = getattr(self, "var_study_verify_excel", None) and self.var_study_verify_excel.get()
            study_records = []
            invalid_kids = []
            if verify_active:
                try:
                    study_records = self.client.read_study_excel(STUDY_EXCEL_FILE)
                    invalid_kids = [r for r in study_records if not r.get("is_valid_format", True)]
                except Exception as e:
                    self.log(f"⚠️ Ошибка чтения study_list.xlsx: {e}", "yellow")

            # Сопоставляем каждую заявку с сервером по Фамилии и Имени
            for o in orders:
                kid_last = (o.get("kid_last_name") or "").strip()
                kid_first = (o.get("kid_first_name") or "").strip()
                kid_patr = (o.get("kid_patronymic") or o.get("kid_middle_name") or "").strip()
                server_fio = f"{kid_last} {kid_first} {kid_patr}".strip() or o.get("site_user_fio", "Ребенок")
                o["server_fio"] = server_fio

                matched_excel_item = None
                if verify_active and study_records:
                    # Ищем совпадение по Фамилии и Имени только среди записей без ошибок формата
                    matched_excel_item = next((r for r in study_records if r.get("is_valid_format", True) and match_fio(o, r["fio"])), None)

                o["excel_item"] = matched_excel_item
                o["is_in_excel"] = (matched_excel_item is not None) if verify_active else True

            # 5. Классификация заявок по статусам и сверке
            orders_initial = []
            orders_approve = []
            orders_study = []
            orders_mismatch = []
            orders_other = []

            for o in orders:
                st = str(o.get("state") or o.get("state_grid") or "").lower().strip()
                o["state_clean"] = st

                if verify_active and not o.get("is_in_excel"):
                    orders_mismatch.append(o)
                elif st == "initial":
                    orders_initial.append(o)
                elif st == "approve":
                    orders_approve.append(o)
                elif st == "study":
                    orders_study.append(o)
                else:
                    orders_other.append(o)

            # Дети из таблицы study_list.xlsx, по которым еще нет заявок на сайте
            matched_excel_rows = {o["excel_item"]["row"] for o in orders if o.get("excel_item")}
            unmatched_excel_kids = [r for r in study_records if r["row"] not in matched_excel_rows] if verify_active else []

            self.cached_study_orders = orders
            self.study_orders_initial = orders_initial
            self.study_orders_approve = orders_approve
            self.study_orders_study = orders_study
            self.study_orders_mismatch = orders_mismatch
            self.study_excel_records = study_records

            self.log(f"✓ Заявки программы '{real_prog_name}' (ID: {event_id}) успешно загружены:", "green")
            self.log(f"   • Фильтр по группе: {group_label}", "cyan")
            self.log(f"   • Учебный год: {year_id}", "cyan")
            self.log(f"   • Всего найдено заявок на сервере: {len(orders)}", "cyan")

            if verify_active:
                self.log(f"   • 📑 Сверка со study_list.xlsx ({len(study_records)} детей в файле):", "cyan")
                self.log(f"      ✓ Сравнение по Фамилии и Имени (отчество в Excel игнорируется): {len(orders) - len(orders_mismatch)} из {len(orders)} совпало", "green")
                if invalid_kids:
                    self.log(f"      ⚠️ ВНИМАНИЕ: {len(invalid_kids)} записей в Excel содержат инициалы вместо полного имени!", "yellow")
                    for ik in invalid_kids:
                        self.log(f"         • Строка {ik['row']}: '{ik['fio']}' — {ik.get('format_error')} (отмечено в файле study_list.xlsx)", "yellow")
                if orders_mismatch:
                    self.log(f"      ⚠️ ВНИМАНИЕ: Найдено {len(orders_mismatch)} заявок на сервере, которых НЕТ в таблице study_list.xlsx!", "yellow")
                    self.log(f"         (Родители могли ошибочно выбрать эту программу на сайте вместо указанной в заявлении)", "yellow")
                if unmatched_excel_kids:
                    self.log(f"      ⏳ ВНИМАНИЕ: {len(unmatched_excel_kids)} детей из таблицы study_list.xlsx еще НЕ подали заявку на сайте", "cyan")

            self.log(f"   • ⏳ Неподтвержденные (initial) — готовы к подтверждению: {len(orders_initial)}", "yellow" if orders_initial else "gray")
            self.log(f"   • 📋 Подтвержденные (approve) — готовы к зачислению на обучение: {len(orders_approve)}", "cyan" if orders_approve else "gray")
            self.log(f"   • 🎓 Обучаются (study) — уже зачислены приказом: {len(orders_study)}", "green" if orders_study else "gray")
            if orders_other:
                self.log(f"   • Другие статусы (отказ/отчисление): {len(orders_other)}", "gray")

            # Выводим подробный список детей в журнал
            if orders:
                self.log("\nСписок заявок по программе:", "cyan")
                for idx, o in enumerate(orders, 1):
                    oid = o.get("id")
                    st = o.get("state") or o.get("state_grid") or "?"
                    fio = o.get("server_fio", "Ребенок")
                    dob_raw = o.get("kid_birthday") or ""
                    dob_str = dob_raw.split(" ")[0] if " " in str(dob_raw) else str(dob_raw)
                    gid = o.get("fact_group_id") or o.get("group_id") or "-"
                    
                    badge = ""
                    if verify_active:
                        if o.get("is_in_excel"):
                            ex_row = o["excel_item"]["row"]
                            curr_flag = o["excel_item"]["flag"] or "флаг пуст"
                            ex_fio = o["excel_item"]["fio"]
                            badge = f" | ✓ Excel (стр.{ex_row}: '{ex_fio}')"
                        else:
                            badge = " | ⚠️ НЕТ В EXCEL (чужая заявка)"

                    color = "yellow" if st == "initial" else ("cyan" if st == "approve" else ("green" if st == "study" else None))
                    if verify_active and not o.get("is_in_excel"):
                        color = "yellow"

                    self.log(f"  [{idx:02d}] Заявка #{oid:<8} | {fio:<30} | ДР: {dob_str:<10} | Группа: {gid}{badge} | Статус: '{st}'", color)

            # Если есть дети из Excel, которых еще нет на сайте
            if verify_active and unmatched_excel_kids:
                self.log("\n📑 Дети из таблицы study_list.xlsx, по которым еще нет заявки на сервере:", "cyan")
                for k in unmatched_excel_kids:
                    fl_info = f", флаг: '{k['flag']}'" if k["flag"] else ""
                    if not k.get("is_valid_format", True):
                        self.log(f"  • {k['fio']} (строка {k['row']}) — ⚠️ НЕПРАВИЛЬНЫЙ ФОРМАТ (инициалы)! Заявка не сопоставится, пока не укажете полное имя", "yellow")
                    else:
                        self.log(f"  • {k['fio']} (строка {k['row']}{fl_info})", "cyan")

            def ui_update():
                self.btn_check_study_orders.config(state=tk.NORMAL)
                
                # Обновляем статусную строку таблицы в интерфейсе
                if hasattr(self, "lbl_study_excel_status") and self.lbl_study_excel_status and verify_active:
                    status_txt = f"Таблица study_list.xlsx: {len(study_records)} детей. На сайте: совпало {len(orders) - len(orders_mismatch)}, чужих: {len(orders_mismatch)}."
                    if invalid_kids:
                        status_txt += f" ⚠️ Ошибок формата (инициалы): {len(invalid_kids)}!"
                    self.lbl_study_excel_status.config(
                        text=status_txt,
                        foreground="#dc2626" if invalid_kids else ("#059669" if not orders_mismatch else "#d97706")
                    )

                # Кнопка подтверждения (initial -> approve)
                if len(orders_initial) > 0:
                    lbl = f"✓ 1. Подтвердить {len(orders_initial)} заявок"
                    if verify_active:
                        lbl += " из Excel"
                    lbl += " (initial)"
                    self.btn_run_study_approve.config(state=tk.NORMAL, text=lbl)
                else:
                    self.btn_run_study_approve.config(state=tk.DISABLED, text="✓ 1. Нет заявок initial")

                # Кнопка отметки об обучении (approve -> study)
                if len(orders_approve) > 0:
                    lbl = f"🎓 2. Зачислить на обучение {len(orders_approve)} заявок"
                    if verify_active:
                        lbl += " из Excel"
                    lbl += " (approve)"
                    self.btn_run_study_mark.config(state=tk.NORMAL, text=lbl)
                else:
                    self.btn_run_study_mark.config(state=tk.DISABLED, text="🎓 2. Нет заявок approve")

                # Кнопка оба шага
                total_to_process = len(orders_initial) + len(orders_approve)
                if total_to_process > 0:
                    lbl = f"⚡ 3. Зачислить всё ({total_to_process} заявок"
                    if verify_active:
                        lbl += " из Excel"
                    lbl += ")"
                    self.btn_run_study_all.config(state=tk.NORMAL, text=lbl)
                else:
                    self.btn_run_study_all.config(state=tk.DISABLED, text="⚡ 3. Все заявки обработаны")

                if not silent:
                    msg_text = (
                        f"Программа: {real_prog_name} (ID: {event_id})\n"
                        f"Группа: {group_label}\n"
                        f"Учебный год: {year_id}\n\n"
                        f"Всего заявок на сервере: {len(orders)}\n"
                    )
                    if verify_active:
                        msg_text += (
                            f"📑 Сверка со study_list.xlsx:\n"
                            f"  • Совпало со списком в Excel (Фамилия и Имя): {len(orders) - len(orders_mismatch)}\n"
                            f"  • Чужих заявок (нет в Excel): {len(orders_mismatch)}\n"
                            f"  • Еще не подали заявку из Excel: {len(unmatched_excel_kids)}\n"
                        )
                        if invalid_kids:
                            msg_text += f"  • ⚠️ Неправильный формат ФИО в Excel (инициалы): {len(invalid_kids)} (отмечено в файле)\n"
                        msg_text += "\n"
                    msg_text += (
                        f"• Готовы к подтверждению (initial): {len(orders_initial)}\n"
                        f"• Готовы к зачислению на обучение (approve): {len(orders_approve)}\n"
                        f"• Уже обучаются (study): {len(orders_study)}\n"
                    )
                    if orders_other:
                        msg_text += f"• Прочие статусы: {len(orders_other)}\n"
                    messagebox.showinfo("Статус заявок программы", msg_text)

            self.root.after(0, ui_update)

        threading.Thread(target=worker, daemon=True).start()

    def run_study_batch_approve(self):
        """
        Массовое подтверждение заявок на программу (state: initial -> approve)
        через POST /api/approveRequest с проверкой по отдельной таблице study_list.xlsx
        и автоматическим переключением флага зачисления.
        """
        if not hasattr(self, "study_orders_initial") or not self.study_orders_initial:
            messagebox.showinfo("Информация", "Нет заявок в статусе 'initial' для подтверждения.")
            return

        cnt = len(self.study_orders_initial)
        pname = getattr(self, "current_study_prog_name", "программе")
        verify_active = getattr(self, "var_study_verify_excel", None) and self.var_study_verify_excel.get()
        verify_note = "\n(Будет выполнена сверка со study_list.xlsx, чужие заявки будут пропущены)" if verify_active else ""

        if not messagebox.askyesno(
            "Подтверждение заявок",
            f"Вы действительно хотите подтвердить {cnt} заявок (state: initial ➜ approve)\n"
            f"по программе '{pname}' через POST /api/approveRequest?{verify_note}"
        ):
            return

        self.set_task_running(True, "Подтверждение заявок программы")
        self.btn_run_study_approve.config(state=tk.DISABLED)
        self.btn_run_study_mark.config(state=tk.DISABLED)
        self.btn_run_study_all.config(state=tk.DISABLED)

        self.log(f"\n[СТАРТ] Подтверждение {cnt} заявок (initial ➜ approve)...", "cyan")
        if verify_active:
            self.log("📑 Защита активна: подтверждаются только дети из таблицы study_list.xlsx", "cyan")

        def worker():
            total_ok = 0
            total_err = 0
            total_skipped = 0
            try:
                for idx, o in enumerate(self.study_orders_initial, 1):
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Подтверждение заявок программы прервано пользователем!", "red")
                        break
                    oid = o.get("id")
                    fio = o.get("server_fio") or f"{o.get('kid_last_name', '')} {o.get('kid_first_name', '')}".strip() or o.get("site_user_fio", "Ребенок")

                    # Проверка по Excel таблице
                    if verify_active and not o.get("is_in_excel"):
                        total_skipped += 1
                        self.log(f"  [{idx}/{cnt}] ⏭️ [ПРОПУСК: ЧУЖАЯ ЗАЯВКА] #{oid} ({fio}): нет в таблице study_list.xlsx (ошибочная заявка родителей)", "yellow")
                        continue

                    ok, msg = self.client.approve_program_order(oid)
                    if getattr(self, "stop_requested", False):
                        break
                    if ok:
                        total_ok += 1
                        self.log(f"  [{idx}/{cnt}] ✓ Заявка #{oid} ({fio}): подтверждена (approve)", "green")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ПОДТВЕРЖДЕНИЕ ПРОГРАММЫ | Заявка #{oid} ({fio}) -> УСПЕШНО (approve)")

                        # Переключаем флаг в отдельной таблице study_list.xlsx
                        if o.get("excel_item"):
                            ex_row = o["excel_item"]["row"]
                            ex_col = o["excel_item"].get("flag_col", 2)
                            flag_txt = f"Подтвержден (#{oid})"
                            ok_fl, fl_msg = self.client.update_study_child_flag(STUDY_EXCEL_FILE, ex_row, flag_txt, ex_col)
                            if ok_fl:
                                self.log(f"      💾 В study_list.xlsx переключен флаг: '{flag_txt}' (строка {ex_row})", "cyan")
                                o["excel_item"]["flag"] = flag_txt
                                o["excel_item"]["is_confirmed"] = True
                            else:
                                self.log(f"      ⚠️ Запись флага в study_list.xlsx: {fl_msg}", "yellow")
                    else:
                        total_err += 1
                        self.log(f"  [{idx}/{cnt}] ✗ Заявка #{oid} ({fio}): ошибка подтверждения ({msg})", "red")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ПОДТВЕРЖДЕНИЕ ПРОГРАММЫ | Заявка #{oid} ({fio}) -> ОШИБКА: {msg}")

                self.log(f"\n[ГОТОВО] Завершено подтверждение заявок:", "cyan")
                self.log(f"  • Успешно подтверждено: {total_ok}", "green")
                if total_skipped:
                    self.log(f"  • Пропущено чужих заявок (нет в Excel): {total_skipped}", "yellow")
                if total_err:
                    self.log(f"  • Ошибок: {total_err}", "red")

                write_to_log_file(
                    "=" * 80 + "\n"
                    f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ИТОГИ ПОДТВЕРЖДЕНИЯ ПРОГРАММЫ (APPROVE)\n"
                    f"Программа: {getattr(self, 'current_study_prog_name', '-')}\n"
                    f"Успешно подтверждено: {total_ok}\n"
                    f"Пропущено чужих заявок: {total_skipped}\n"
                    f"Ошибок: {total_err}\n"
                    + "=" * 80 + "\n\n"
                )
            except InterruptedError:
                self.log("\n🛑 [СТОП] Подтверждение заявок программы немедленно прервано!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    skip_info = f"\nПропущено чужих заявок (нет в Excel): {total_skipped}" if total_skipped else ""
                    messagebox.showinfo(
                        "Подтверждение завершено",
                        f"Успешно подтверждено: {total_ok} из {cnt} заявок!\n"
                        f"{skip_info}\n"
                        f"Флаги в таблице study_list.xlsx обновлены.\n"
                        f"Теперь эти заявки можно зачислить на обучение приказом (кнопка «🎓 2. Отметка об обучении»)."
                    )
                    self.check_study_program_orders(silent=True)
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def run_study_batch_study(self):
        """
        Массовая отметка об обучении (state: approve -> study)
        через POST /api/studyRequest с указанием номера приказа и дат.
        Сверяет детей со study_list.xlsx и обновляет флаг зачисления.
        """
        if not hasattr(self, "study_orders_approve") or not self.study_orders_approve:
            messagebox.showinfo("Информация", "Нет заявок в статусе 'approve' для зачисления на обучение.")
            return

        cnt = len(self.study_orders_approve)
        pname = getattr(self, "current_study_prog_name", "программе")

        decree_num = self.entry_study_decree.get().strip() if hasattr(self, "entry_study_decree") else "183"
        raw_signing = self.entry_study_signing.get().strip() if hasattr(self, "entry_study_signing") else "2026-08-31"
        raw_start = self.entry_study_start.get().strip() if hasattr(self, "entry_study_start") else "2026-09-01"
        date_signing = normalize_date_input(raw_signing, "2026-08-31")
        date_start = normalize_date_input(raw_start, "2026-09-01")
        fin_src = self.saved_cfg.get("study_financing_source", "1")
        verify_active = getattr(self, "var_study_verify_excel", None) and self.var_study_verify_excel.get()

        if not messagebox.askyesno(
            "Зачисление на обучение",
            f"Вы действительно хотите зачислить {cnt} заявок на обучение (state: approve ➜ study)\n\n"
            f"Программа: '{pname}'\n"
            f"№ приказа: {decree_num}\n"
            f"Дата приказа: {date_signing}\n"
            f"Дата начала обучения: {date_start}\n"
            f"Источник финансирования: {fin_src}\n"
            f"Сверка со study_list.xlsx: {'ВКЛЮЧЕНА' if verify_active else 'ВЫКЛЮЧЕНА'}\n\n"
            f"Запрос: POST /api/studyRequest"
        ):
            return

        self.set_task_running(True, "Зачисление на обучение (study)")
        self.btn_run_study_approve.config(state=tk.DISABLED)
        self.btn_run_study_mark.config(state=tk.DISABLED)
        self.btn_run_study_all.config(state=tk.DISABLED)

        self.log(f"\n[СТАРТ] Зачисление {cnt} заявок на обучение (approve ➜ study, приказ №{decree_num})...", "cyan")
        if verify_active:
            self.log("📑 Защита активна: зачисляются только дети из таблицы study_list.xlsx", "cyan")

        def worker():
            total_ok = 0
            total_err = 0
            total_skipped = 0
            try:
                for idx, o in enumerate(self.study_orders_approve, 1):
                    if getattr(self, "stop_requested", False):
                        self.log("\n🛑 [СТОП] Зачисление на обучение прервано пользователем!", "red")
                        break
                    oid = o.get("id")
                    fio = o.get("server_fio") or f"{o.get('kid_last_name', '')} {o.get('kid_first_name', '')}".strip() or o.get("site_user_fio", "Ребенок")

                    # Проверка по Excel таблице
                    if verify_active and not o.get("is_in_excel"):
                        total_skipped += 1
                        self.log(f"  [{idx}/{cnt}] ⏭️ [ПРОПУСК: ЧУЖАЯ ЗАЯВКА] #{oid} ({fio}): нет в таблице study_list.xlsx", "yellow")
                        continue

                    # Если у заявки есть свой доступный источник финансирования, берем его
                    avail_fin = o.get("available_financing_source")
                    current_fin = str(avail_fin[0]) if (isinstance(avail_fin, list) and avail_fin) else fin_src

                    ok, msg = self.client.study_program_order(
                        order_id=oid,
                        financing_source=current_fin,
                        date_start=date_start,
                        decree_number=decree_num,
                        date_signing=date_signing
                    )
                    if getattr(self, "stop_requested", False):
                        break
                    if ok:
                        total_ok += 1
                        self.log(f"  [{idx}/{cnt}] ✓ Заявка #{oid} ({fio}): зачислена на обучение (study)", "green")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ОБУЧЕНИЕ ПРОГРАММЫ | Заявка #{oid} ({fio}) -> УСПЕШНО (study, приказ №{decree_num})")

                        # Переключаем флаг зачисления в отдельной таблице study_list.xlsx
                        if o.get("excel_item"):
                            ex_row = o["excel_item"]["row"]
                            ex_col = o["excel_item"].get("flag_col", 2)
                            flag_txt = f"Зачислен (приказ №{decree_num} от {date_signing})"
                            ok_fl, fl_msg = self.client.update_study_child_flag(STUDY_EXCEL_FILE, ex_row, flag_txt, ex_col)
                            if ok_fl:
                                self.log(f"      💾 В study_list.xlsx переключен флаг: '{flag_txt}' (строка {ex_row})", "cyan")
                                o["excel_item"]["flag"] = flag_txt
                                o["excel_item"]["is_enrolled"] = True
                            else:
                                self.log(f"      ⚠️ Запись флага в study_list.xlsx: {fl_msg}", "yellow")
                    else:
                        total_err += 1
                        self.log(f"  [{idx}/{cnt}] ✗ Заявка #{oid} ({fio}): ошибка зачисления ({msg})", "red")
                        write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ОБУЧЕНИЕ ПРОГРАММЫ | Заявка #{oid} ({fio}) -> ОШИБКА: {msg}")

                self.log(f"\n[ГОТОВО] Завершено зачисление на обучение:", "cyan")
                self.log(f"  • Успешно зачислено (study): {total_ok}", "green")
                if total_skipped:
                    self.log(f"  • Пропущено чужих заявок (нет в Excel): {total_skipped}", "yellow")
                if total_err:
                    self.log(f"  • Ошибок: {total_err}", "red")

                write_to_log_file(
                    "=" * 80 + "\n"
                    f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ИТОГИ ЗАЧИСЛЕНИЯ НА ОБУЧЕНИЕ (STUDY)\n"
                    f"Программа: {getattr(self, 'current_study_prog_name', '-')}\n"
                    f"Приказ №: {decree_num} от {date_signing} (начало: {date_start})\n"
                    f"Успешно зачислено: {total_ok}\n"
                    f"Пропущено чужих заявок: {total_skipped}\n"
                    f"Ошибок: {total_err}\n"
                    + "=" * 80 + "\n\n"
                )
            except InterruptedError:
                self.log("\n🛑 [СТОП] Зачисление на обучение немедленно прервано!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    skip_info = f"\nПропущено чужих заявок: {total_skipped}" if total_skipped else ""
                    messagebox.showinfo(
                        "Зачисление на обучение завершено",
                        f"Успешно зачислено на обучение: {total_ok} из {cnt} заявок!\n"
                        f"{skip_info}\n"
                        f"Статус: study (обучение по приказу №{decree_num}).\n"
                        f"Флаги в study_list.xlsx обновлены."
                    )
                    self.check_study_program_orders(silent=True)
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()

    def run_study_batch_all(self):
        """
        Выполняет оба шага подряд:
        1. Подтверждает все initial заявки (initial -> approve)
        2. Зачисляет все approve заявки на обучение (approve -> study)
        Сверяет детей со study_list.xlsx и обновляет флаги зачисления на каждом этапе.
        """
        initial_list = getattr(self, "study_orders_initial", [])
        approve_list = getattr(self, "study_orders_approve", [])

        total = len(initial_list) + len(approve_list)
        if total == 0:
            messagebox.showinfo("Информация", "Нет заявок для подтверждения или зачисления.")
            return

        decree_num = self.entry_study_decree.get().strip() if hasattr(self, "entry_study_decree") else "183"
        raw_signing = self.entry_study_signing.get().strip() if hasattr(self, "entry_study_signing") else "2026-08-31"
        raw_start = self.entry_study_start.get().strip() if hasattr(self, "entry_study_start") else "2026-09-01"
        date_signing = normalize_date_input(raw_signing, "2026-08-31")
        date_start = normalize_date_input(raw_start, "2026-09-01")
        fin_src = self.saved_cfg.get("study_financing_source", "1")
        verify_active = getattr(self, "var_study_verify_excel", None) and self.var_study_verify_excel.get()

        if not messagebox.askyesno(
            "Полное зачисление (оба шага)",
            f"Будет выполнено:\n"
            f"1. Подтверждение заявок initial ➜ approve ({len(initial_list)} шт)\n"
            f"2. Зачисление на обучение approve ➜ study ({len(approve_list)} шт + вновь подтвержденные)\n\n"
            f"Сверка со study_list.xlsx: {'ВКЛЮЧЕНА' if verify_active else 'ВЫКЛЮЧЕНА'}\n"
            f"№ приказа: {decree_num}\n"
            f"Дата приказа: {date_signing}\n"
            f"Дата начала: {date_start}\n\n"
            f"Продолжить выполнение?"
        ):
            return

        self.set_task_running(True, "Полный цикл зачисления (approve + study)")
        self.btn_run_study_approve.config(state=tk.DISABLED)
        self.btn_run_study_mark.config(state=tk.DISABLED)
        self.btn_run_study_all.config(state=tk.DISABLED)

        self.log(f"\n[СТАРТ] Полный цикл зачисления: Шаг 1 (Подтверждение) + Шаг 2 (Обучение)...", "cyan")
        if verify_active:
            self.log("📑 Защита активна: зачисляются только дети из таблицы study_list.xlsx", "cyan")

        def worker():
            try:
                # ШАГ 1: Подтверждение initial
                approved_orders_ready = list(approve_list)
                if initial_list:
                    self.log(f"\n--- ШАГ 1: Подтверждение {len(initial_list)} заявок (initial -> approve) ---", "cyan")
                    for idx, o in enumerate(initial_list, 1):
                        if getattr(self, "stop_requested", False):
                            self.log("\n🛑 [СТОП] Полный цикл прерван пользователем!", "red")
                            break
                        oid = o.get("id")
                        fio = o.get("server_fio") or f"{o.get('kid_last_name', '')} {o.get('kid_first_name', '')}".strip() or o.get("site_user_fio", "Ребенок")

                        if verify_active and not o.get("is_in_excel"):
                            self.log(f"  [1.{idx}] ⏭️ [ПРОПУСК: ЧУЖАЯ ЗАЯВКА] #{oid} ({fio}): нет в таблице study_list.xlsx", "yellow")
                            continue

                        ok, msg = self.client.approve_program_order(oid)
                        if getattr(self, "stop_requested", False):
                            break
                        if ok:
                            self.log(f"  [1.{idx}] ✓ Заявка #{oid} ({fio}): подтверждена", "green")
                            approved_orders_ready.append(o)

                            # Переключаем флаг в study_list.xlsx
                            if o.get("excel_item"):
                                ex_row = o["excel_item"]["row"]
                                ex_col = o["excel_item"].get("flag_col", 2)
                                flag_txt = f"Подтвержден (#{oid})"
                                ok_fl, fl_msg = self.client.update_study_child_flag(STUDY_EXCEL_FILE, ex_row, flag_txt, ex_col)
                                if ok_fl:
                                    self.log(f"      💾 В study_list.xlsx переключен флаг: '{flag_txt}' (строка {ex_row})", "cyan")
                                    o["excel_item"]["flag"] = flag_txt
                                    o["excel_item"]["is_confirmed"] = True
                        else:
                            self.log(f"  [1.{idx}] ✗ Заявка #{oid} ({fio}): ошибка подтверждения ({msg})", "red")

                # ШАГ 2: Зачисление approve -> study
                if approved_orders_ready and not getattr(self, "stop_requested", False):
                    self.log(f"\n--- ШАГ 2: Зачисление {len(approved_orders_ready)} заявок на обучение (approve -> study) ---", "cyan")
                    total_study_ok = 0
                    total_study_err = 0
                    for idx, o in enumerate(approved_orders_ready, 1):
                        if getattr(self, "stop_requested", False):
                            self.log("\n🛑 [СТОП] Полный цикл прерван пользователем!", "red")
                            break
                        oid = o.get("id")
                        fio = o.get("server_fio") or f"{o.get('kid_last_name', '')} {o.get('kid_first_name', '')}".strip() or o.get("site_user_fio", "Ребенок")

                        if verify_active and not o.get("is_in_excel"):
                            self.log(f"  [2.{idx}] ⏭️ [ПРОПУСК: ЧУЖАЯ ЗАЯВКА] #{oid} ({fio}): нет в таблице study_list.xlsx", "yellow")
                            continue

                        avail_fin = o.get("available_financing_source")
                        current_fin = str(avail_fin[0]) if (isinstance(avail_fin, list) and avail_fin) else fin_src
                        ok, msg = self.client.study_program_order(
                            order_id=oid,
                            financing_source=current_fin,
                            date_start=date_start,
                            decree_number=decree_num,
                            date_signing=date_signing
                        )
                        if getattr(self, "stop_requested", False):
                            break
                        if ok:
                            total_study_ok += 1
                            self.log(f"  [2.{idx}] ✓ Заявка #{oid} ({fio}): успешно зачислена на обучение!", "green")

                            # Переключаем флаг в study_list.xlsx
                            if o.get("excel_item"):
                                ex_row = o["excel_item"]["row"]
                                ex_col = o["excel_item"].get("flag_col", 2)
                                flag_txt = f"Зачислен (приказ №{decree_num} от {date_signing})"
                                ok_fl, fl_msg = self.client.update_study_child_flag(STUDY_EXCEL_FILE, ex_row, flag_txt, ex_col)
                                if ok_fl:
                                    self.log(f"      💾 В study_list.xlsx переключен флаг: '{flag_txt}' (строка {ex_row})", "cyan")
                                    o["excel_item"]["flag"] = flag_txt
                                    o["excel_item"]["is_enrolled"] = True
                        else:
                            total_study_err += 1
                            self.log(f"  [2.{idx}] ✗ Заявка #{oid} ({fio}): ошибка ({msg})", "red")

                    write_to_log_file(
                        "=" * 80 + "\n"
                        f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ИТОГИ ПОЛНОГО ЦИКЛА (ШАГ 1: APPROVE + ШАГ 2: STUDY)\n"
                        f"Программа: {getattr(self, 'current_study_prog_name', '-')}\n"
                        f"Приказ №: {decree_num} от {date_signing} (начало: {date_start})\n"
                        f"Успешно зачислено на обучение (Шаг 2): {total_study_ok}\n"
                        f"Ошибок: {total_study_err}\n"
                        f"Флаги в study_list.xlsx обновлены.\n"
                        + "=" * 80 + "\n\n"
                    )
                    self.log(f"\n[ГОТОВО] Полный цикл завершен: успешно зачислено {total_study_ok} детей на обучение!", "green")
                    if not getattr(self, "stop_requested", False):
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Полный цикл завершен",
                            f"Успешно завершено!\n\n"
                            f"✓ Зачислено на обучение: {total_study_ok}\n"
                            f"Приказ №: {decree_num} от {date_signing}\n"
                            f"Флаги в study_list.xlsx обновлены."
                        ))
                else:
                    if not getattr(self, "stop_requested", False):
                        self.log("\n[ГОТОВО] Нет заявок для зачисления на обучение.", "yellow")

            except InterruptedError:
                self.log("\n🛑 [СТОП] Полный цикл зачисления немедленно прерван!", "red")
            finally:
                def _done():
                    if getattr(self, "stop_requested", False):
                        self.reset_to_initial_state()
                        return
                    self.set_task_running(False)
                    self.check_study_program_orders(silent=True)
                self.root.after(0, _done)

        threading.Thread(target=worker, daemon=True).start()


# =========================================================================
# ТОЧКА ВХОДА
# =========================================================================
if __name__ == "__main__":
    # Включаем аппаратный High-DPI (Windows Per-Monitor DPI Awareness v2)
    # до создания первого экземпляра tk.Tk()
    enable_high_dpi()

    root = tk.Tk()

    # Корректируем коэффициент масштабирования Tkinter под системный DPI
    if sys.platform.startswith("win"):
        try:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            if dpi > 0:
                # 96 DPI = 100% масштаб. При 120 DPI (125%) или 144 DPI (150%) передаем scaling
                root.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass

    # Установка иконки приложения (app.ico), если файл доступен
    ico_candidates = [
        getattr(sys, "_MEIPASS", None) and os.path.join(getattr(sys, "_MEIPASS", ""), "app.ico"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico"),
        "app.ico",
    ]
    for ico_path in ico_candidates:
        if ico_path and os.path.exists(ico_path):
            try:
                root.iconbitmap(ico_path)
                break
            except Exception:
                pass

    app = NavigatorApp(root)

    def on_closing():
        try:
            write_to_log_file(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🏁 ЗАВЕРШЕНИЕ РАБОТЫ ПРОГРАММЫ (ОКНО ЗАКРЫТО)")
            write_to_log_file("=" * 85 + "\n")
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
