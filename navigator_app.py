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
import subprocess
import threading
from typing import Optional, Tuple, Dict, Any, List
import requests

CONFIG_FILE = "config.json"
EXCEL_FILE = "list.xlsx"
LOG_FILE = "results_log.txt"


# =========================================================================
# 1. СЕРВИСНЫЕ ФУНКЦИИ (КОНФИГ, ЛОГИ И EXCEL)
# =========================================================================
def write_to_log_file(text: str) -> None:
    """Безопасно дописывает запись в файл results_log.txt в кодировке UTF-8."""
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception as e:
        print(f"Ошибка записи в {LOG_FILE}: {e}")


def ensure_log_file() -> None:
    """Создает файл results_log.txt с базовым заголовком, если он еще не существует."""
    if not os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.write("=" * 80 + "\n")
                f.write("ЖУРНАЛ ОПЕРАЦИЙ НАВИГАТОРА (results_log.txt)\n")
                f.write("Здесь сохраняются подробные результаты создания и подтверждения заявок по детям.\n")
                f.write(f"Создан: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
        except Exception as e:
            print(f"Ошибка инициализации {LOG_FILE}: {e}")
DEFAULT_CONFIG: Dict[str, Any] = {
    "email": "",
    "password": "",
    "saved_at": "",
    "activity_name": "Мастер-класс по анимации в ДОЛ Горный воздух",
    "activity_datetime": "2026-08-31 11:00:00",
    "confirm_activity_name": "Мастер-класс по анимации в ДОЛ Горный воздух",
}


def load_config() -> Dict[str, Any]:
    """Загружает конфиг, объединяя сохраненные параметры с дефолтными значениями-примерами."""
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    cfg.update(loaded)
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


def ensure_default_excel():
    """Создает шаблонный файл list.xlsx, если его нет, либо добавляет третий столбец статуса."""
    try:
        import openpyxl
        if not os.path.exists(EXCEL_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            # Заголовки (3 столбца)
            ws.cell(row=1, column=1, value="ФИО")
            ws.cell(row=1, column=2, value="Дата рождения")
            ws.cell(row=1, column=3, value="Статус заявки")
            # Пример строки
            ws.cell(row=2, column=1, value="Солодовникова Романа Александровна")
            ws.cell(row=2, column=2, value="23.06.2005")
            ws.cell(row=2, column=3, value="")  # Пустое: ребенок еще не обработан
            ws.column_dimensions["A"].width = 38
            ws.column_dimensions["B"].width = 20
            ws.column_dimensions["C"].width = 28
            wb.save(EXCEL_FILE)
            wb.close()
        else:
            # Если файл существует, проверяем наличие столбца со статусом
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb.active
                if ws.max_column < 3 or not ws.cell(1, 3).value:
                    ws.cell(row=1, column=3, value="Статус заявки")
                    ws.column_dimensions["C"].width = 28
                    wb.save(EXCEL_FILE)
                wb.close()
            except Exception:
                pass
    except Exception as e:
        print(f"Не удалось инициализировать list.xlsx: {e}")


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
        return False, "Файл list.xlsx открыт в Microsoft Excel или другой программе (файл заблокирован)"
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
    def __init__(self, min_interval: float = 1.2):
        # Гарантируем, что интервал не может быть меньше 1.0 с (не более 1 запроса в секунду)
        self.min_interval = max(1.0, float(min_interval))
        self.current_delay = self.min_interval
        self.last_request_time = 0.0
        self.last_latency_ms: Optional[float] = None
        self.last_status_code: int = 200

    def wait_before_request(self) -> float:
        """Блокирует поток на необходимое время, гарантируя соблюдение паузы между запросами."""
        now = time.time()
        elapsed = now - self.last_request_time
        needed_wait = self.current_delay - elapsed
        if needed_wait > 0:
            time.sleep(needed_wait)
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
        })
        self.access_token: Optional[str] = None
        self.user_info: Dict[str, Any] = {}

        # Адаптивный контроллер (строго <= 1 запрос в секунду, базовая задержка 1.2 с)
        self.limiter = AdaptiveRateLimiter(min_interval=1.2)
        self.on_ping_update: Optional[Callable[[float, str, str, float], None]] = None
        self.on_log_message: Optional[Callable[[str, Optional[str]], None]] = None

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Выполняет запрос с соблюдением адаптивного лимита скорости и перехватом задержки."""
        max_retries = 2
        for attempt in range(max_retries + 1):
            # Соблюдаем безопасную паузу: не более 1 запроса в секунду
            cur_delay = self.limiter.wait_before_request()

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
                    time.sleep(wait_time)
                    continue

                return resp
            except Exception as exc:
                lat_ms = (time.perf_counter() - start_t) * 1000.0
                label, new_delay, color = self.limiter.record_response(lat_ms, 504)
                if self.on_ping_update:
                    try:
                        self.on_ping_update(lat_ms, label, color, new_delay)
                    except Exception:
                        pass
                if attempt < max_retries:
                    time.sleep(3.0)
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

        # Маркеры успешной обработки заявки
        processed_keywords = ["добавлен", "обработан", "успешно", "заявка", "да", "ок", "ok", "+", "true", "1"]

        records = []
        for r in range(2, ws.max_row + 1):
            fio = ws.cell(r, fio_col).value
            dob = ws.cell(r, dob_col).value
            raw_status = ws.cell(r, status_col).value if ws.max_column >= status_col else None
            status_str = str(raw_status).strip() if raw_status is not None else ""

            # Если в статусе уже зафиксировано успешное добавление
            is_processed = bool(
                status_str and any(kw in status_str.lower() for kw in processed_keywords)
            )

            if fio and str(fio).strip():
                records.append({
                    "row": r,
                    "fio": str(fio).strip(),
                    "dob": dob,
                    "status": status_str,
                    "is_processed": is_processed,
                    "status_col": status_col,
                })
        wb.close()
        return records


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
        if isinstance(w, (tk.Entry, ttk.Entry, tk.Text)):
            return w
        if event and hasattr(event, "widget") and isinstance(event.widget, (tk.Entry, ttk.Entry, tk.Text)):
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
        if isinstance(widget, (tk.Entry, ttk.Entry, tk.Text)):
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
        if not isinstance(widget, (tk.Entry, ttk.Entry, tk.Text)):
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
        self.root.title("Навигатор: Автоматизация (NavAdd & NavConfirm)")
        self.root.geometry("860x700")
        self.root.minsize(800, 620)

        # Подключаем глобальную поддержку буфера обмена (Ctrl+C, Ctrl+V в RU/EN + ПКМ)
        setup_universal_clipboard(self.root)

        self.client = NavigatorClient()
        self.saved_cfg = load_config()

        # Настраиваем мониторинг отклика сервера в реальном времени
        self.is_monitoring_active = True
        self.client.on_ping_update = self._on_server_ping_update
        self.client.on_log_message = self.log

        # Создаем шаблонную таблицу и файл журнала при первом запуске
        ensure_default_excel()
        ensure_log_file()

        # Стили интерфейса
        self._setup_styles()

        # Запускаем фоновый цикл периодического замера отклика
        self._start_bg_ping_loop()

        # Стартовое состояние: окно авторизации
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.show_login_screen()

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
        self.entry_email = ttk.Entry(new_box, width=38, font=("Segoe UI", 10))
        self.entry_email.pack(fill=tk.X, pady=(2, 8))

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

    def on_new_login_click(self):
        email = self.entry_email.get().strip()
        pwd = self.entry_password.get().strip()
        if not email or not pwd:
            messagebox.showwarning("Внимание", "Пожалуйста, заполните email и пароль.")
            return
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
            self.show_main_dashboard()
        else:
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

        ttk.Button(btn_box, text="📂 Открыть list.xlsx", command=lambda: open_file_in_os(EXCEL_FILE)).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box, text="📋 Открыть results_log.txt", command=lambda: open_file_in_os(LOG_FILE)).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box, text="⚙ Открыть config.json", command=lambda: open_file_in_os(CONFIG_FILE)).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box, text="Выйти", command=self.show_login_screen).pack(side=tk.LEFT, padx=3)

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

        # Основной блок с двумя режимами
        tabs = ttk.Notebook(self.main_frame)
        tabs.pack(fill=tk.X, pady=(0, 10))

        # --- ВКЛАДКА 1: ПАКЕТНАЯ ЗАПИСЬ (NavAdd) ---
        tab_add = ttk.Frame(tabs, padding=12)
        tabs.add(tab_add, text=" 📝 1. Пакетная запись детей из Excel ")

        ttk.Label(tab_add, text="Название мероприятия:").grid(row=0, column=0, sticky=tk.W, pady=(4, 0))
        self.entry_add_event = ttk.Entry(tab_add, font=("Segoe UI", 10), width=50)
        self.entry_add_event.grid(row=0, column=1, columnspan=2, sticky=tk.W + tk.E, pady=(4, 0))
        self.entry_add_event.insert(0, initial_activity)

        # Подсказка-пример валидного ввода для мероприятия
        ttk.Label(
            tab_add,
            text="Пример: Мастер-класс по анимации в ДОЛ Горный воздух (сохраняется в config.json)",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(tab_add, text="Дата и время участия:").grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        self.entry_add_dt = ttk.Entry(tab_add, font=("Segoe UI", 10), width=28)
        self.entry_add_dt.grid(row=2, column=1, sticky=tk.W, pady=(4, 0))
        self.entry_add_dt.insert(0, initial_datetime)

        # Привязываем авто-сохранение в config.json при потере фокуса (FocusOut)
        def save_add_fields(event=None):
            try:
                name = self.entry_add_event.get().strip() if hasattr(self, "entry_add_event") else ""
                dt = self.entry_add_dt.get().strip() if hasattr(self, "entry_add_dt") else ""
                if name or dt:
                    update_config(activity_name=name, activity_datetime=dt)
                    self.saved_cfg = load_config()
            except Exception:
                pass

        self.entry_add_event.bind("<FocusOut>", save_add_fields)
        self.entry_add_dt.bind("<FocusOut>", save_add_fields)

        # Подсказка-пример валидного формата даты/времени
        ttk.Label(
            tab_add,
            text="Формат: ГГГГ-ММ-ДД ЧЧ:ММ:СС (пример: 2026-08-31 11:00:00, сохраняется в config.json)",
            foreground="#64748b",
            font=("Segoe UI", 8)
        ).grid(row=3, column=1, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(tab_add, text="Файл таблицы:", foreground="gray").grid(row=4, column=0, sticky=tk.W, pady=4)
        ttk.Label(tab_add, text=f"Используется '{EXCEL_FILE}' (кнопка сверху для редактирования)", foreground="#2563eb", font=("Segoe UI", 9, "italic")).grid(row=4, column=1, sticky=tk.W, pady=4)

        self.btn_run_add = ttk.Button(tab_add, text="▶ Запустить запись детей из таблицы", command=self.run_batch_add)
        self.btn_run_add.grid(row=5, column=0, columnspan=3, sticky=tk.W + tk.E, pady=(8, 4))

        # --- ВКЛАДКА 2: ПОДТВЕРЖДЕНИЕ ЗАЯВОК И ОТМЕТКА УЧАСТИЯ (NavConfirm) ---
        tab_conf = ttk.Frame(tabs, padding=12)
        tabs.add(tab_conf, text=" ✓ 2. Подтверждение и отметка участия ")

        ttk.Label(tab_conf, text="Название мероприятия:").grid(row=0, column=0, sticky=tk.W, pady=(4, 0))
        self.entry_conf_event = ttk.Entry(tab_conf, font=("Segoe UI", 10), width=50)
        self.entry_conf_event.grid(row=0, column=1, sticky=tk.W + tk.E, pady=(4, 0))
        self.entry_conf_event.insert(0, initial_confirm_act)

        def save_conf_fields(event=None):
            try:
                name = self.entry_conf_event.get().strip() if hasattr(self, "entry_conf_event") else ""
                if name:
                    update_config(confirm_activity_name=name)
                    self.saved_cfg = load_config()
            except Exception:
                pass

        self.entry_conf_event.bind("<FocusOut>", save_conf_fields)

        ttk.Label(
            tab_conf,
            text="Пример: Мастер-класс по анимации в ДОЛ Горный воздух (сохраняется в config.json)",
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
        self.log(f"Таблица данных: {os.path.abspath(EXCEL_FILE)}")
        self.log("При записи проверяются: 'is_approved': true и точное совпадение ДР.")

    def log(self, text: str, tag: Optional[str] = None):
        def _append():
            self.txt_log.insert(tk.END, text + "\n", tag)
            self.txt_log.see(tk.END)
        self.root.after(0, _append)

    # ---------------------------------------------------------------------
    # ОБРАБОТЧИКИ ОПЕРАЦИЙ
    # ---------------------------------------------------------------------
    def run_batch_add(self):
        event_name = self.entry_add_event.get().strip()
        event_dt = self.entry_add_dt.get().strip()

        if not event_name or not event_dt:
            messagebox.showwarning("Внимание", "Заполните название мероприятия и дату/время.")
            return

        if not os.path.exists(EXCEL_FILE):
            messagebox.showerror("Ошибка", f"Файл {EXCEL_FILE} не найден. Нажмите 'Открыть list.xlsx' для создания.")
            return

        # Сохраняем в config.json при каждом запуске
        update_config(activity_name=event_name, activity_datetime=event_dt)
        self.saved_cfg = load_config()

        self.btn_run_add.config(state=tk.DISABLED)
        self.log("\n" + "=" * 60, "cyan")
        self.log(f"[СТАРТ] Поиск мероприятия: '{event_name}'...", "cyan")

        def worker():
            act, msg = self.client.search_activity(event_name)
            if not act:
                self.log(f"✗ Мероприятие не найдено: {msg}", "red")
                self.root.after(0, lambda: self.btn_run_add.config(state=tk.NORMAL))
                return

            act_id = int(act["id"])
            self.log(f"✓ Найдено: '{act.get('name')}' (ID: {act_id})", "green")
            self.log(f"Дата и время заявки: {event_dt}")

            try:
                records = self.client.read_excel(EXCEL_FILE)
            except Exception as e:
                self.log(f"✗ Ошибка чтения Excel: {e}", "red")
                self.root.after(0, lambda: self.btn_run_add.config(state=tk.NORMAL))
                return

            self.log(f"Загружено записей из таблицы: {len(records)}\n", "cyan")

            # Записываем начало сессии в файл results_log.txt
            start_time_str = time.strftime("%Y-%m-%d %H:%M:%S")
            write_to_log_file("=" * 80)
            write_to_log_file(f"[{start_time_str}] ПАКЕТНАЯ ЗАПИСЬ ДЕТЕЙ ИЗ ТАБЛИЦЫ")
            write_to_log_file(f"Мероприятие:         {act.get('name')} (ID: {act_id})")
            write_to_log_file(f"Дата и время заявки: {event_dt}")
            write_to_log_file(f"Файл таблицы:        {os.path.abspath(EXCEL_FILE)} (строк: {len(records)})")
            write_to_log_file("-" * 80)

            stats = {"ok": 0, "skip_already": 0, "skip_unapproved": 0, "skip_dob": 0, "skip_notfound": 0, "err": 0}

            for item in records:
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

                self.log(f"[{r_num}] {fio} (ДР: {raw_dob} -> {norm_dob or 'нет'})")

                kid, search_msg = self.client.find_kid(fio, norm_dob)
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
                if ok:
                    self.log(f"   ✓ УСПЕШНО: {ord_msg} (Kid ID: {kid.get('id')})", "green")
                    stats["ok"] += 1

                    # Автоматически переключаем флаг в 3-м столбце Excel, чтобы не допустить дублирования
                    order_tag = f"Заявка #{oid}" if oid else "Заявка создана"
                    excel_flag = f"Добавлен ({order_tag})"
                    saved, save_msg = self.client.update_child_status(EXCEL_FILE, r_num, excel_flag, item.get("status_col", 3))
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

                # Частота запросов адаптивно регулируется AdaptiveRateLimiter (строго <= 1 запр/сек)

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

            self.root.after(0, lambda: self.btn_run_add.config(state=tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo(
                "Завершено",
                f"Обработка завершена!\n\n"
                f"✓ Успешно создано заявок: {stats['ok']}\n"
                f"⏭️ Пропущено (ранее добавлены): {stats['skip_already']}\n"
                f"✗ Пропущено других: {stats['skip_unapproved'] + stats['skip_dob'] + stats['skip_notfound']}\n"
                f"❗ Ошибок: {stats['err']}"
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

        self.btn_run_confirm.config(state=tk.DISABLED)
        self.btn_run_participant.config(state=tk.DISABLED)
        self.log("\n[СТАРТ] Массовое подтверждение заявок (state -> approve)...", "cyan")

        def worker():
            total_approved = 0
            total_err = 0
            while True:
                batch, remaining = self.client.get_pending_orders(act_id, limit=100)
                if not batch:
                    break
                for o in batch:
                    oid = o["id"]
                    child_fio = o.get("child_fio") or o.get("child_name") or ""
                    name_str = f" ({child_fio})" if child_fio else ""
                    ok, ord_msg = self.client.approve_order(oid)
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
            self.root.after(0, lambda: messagebox.showinfo(
                "Подтверждение завершено",
                f"Успешно подтверждено заявок: {total_approved}!\n\n"
                "Теперь эти заявки можно отметить как участников (кнопка «🎖️ 2. Отметить участие»)."
            ))
            # Автоматически обновляем статус и доступность кнопок
            self.check_pending_orders(silent=True)

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

        self.btn_run_confirm.config(state=tk.DISABLED)
        self.btn_run_participant.config(state=tk.DISABLED)
        self.log("\n[СТАРТ] Массовая отметка участия детей (state -> participant)...", "cyan")

        def worker():
            total_participant = 0
            total_err = 0
            while True:
                batch, remaining = self.client.get_approved_orders(act_id, limit=100)
                if not batch:
                    break
                for o in batch:
                    oid = o["id"]
                    child_fio = o.get("child_fio") or o.get("child_name") or ""
                    name_str = f" ({child_fio})" if child_fio else ""
                    ok, ord_msg = self.client.mark_participant_order(oid)
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
            self.root.after(0, lambda: messagebox.showinfo(
                "Отметка участия завершена",
                f"Успешно отмечено участие для {total_participant} заявок!\n"
                "(state -> participant)"
            ))
            # Автоматически обновляем статус и доступность кнопок
            self.check_pending_orders(silent=True)

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

    app = NavigatorApp(root)
    root.mainloop()
