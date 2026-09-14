"""
NavigatorClient: Универсальный легковесный клиент для автоматизации портала
"Навигатор дополнительного образования" через прямой REST API.

Полностью заменяет Selenium для:
1. Авторизации (POST /api/user/login) -> получение JWT access_token.
2. Поиска мероприятий (GET /api/rest/activity/) -> получение activity_id по названию.
3. Поиска детей (GET /api/activity/rest/kid) -> получение kid_id, site_user_id и сверка ДР.
4. Создания заявок (POST /api/rest/activityOrder) -> запись ребенка на мероприятие.
5. Выборки новых заявок пачками (GET /api/rest/activityOrder) -> фильтрация по activity_id и state=initial.
6. Подтверждения заявок и участия (POST /api/setActivityOrderState) -> перевод в state=approve.
"""

import json
import time
from typing import Optional, Tuple, Dict, Any, List, Callable
import requests


class NavigatorClient:
    """
    Клиент для взаимодействия с API регионального Навигатора.
    Все сетевые вызовы выполняются через единую requests.Session с keep-alive.
    """

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
        self.refresh_token: Optional[str] = None
        self.user_info: Dict[str, Any] = {}

    # =========================================================================
    # 1. АВТОРИЗАЦИЯ
    # =========================================================================
    def login(self, email: str, password: str) -> Tuple[bool, str]:
        """
        Вход в админку Навигатора.
        Получает JWT-токен и сохраняет его в заголовке Authorization для последующих вызовов.
        """
        url = f"{self.base_url}/api/user/login"
        payload = {
            "email": email.strip(),
            "password": password.strip()
        }

        try:
            resp = self.session.post(url, json=payload, timeout=15)
            data = resp.json()

            if data.get("success") and "access_token" in data.get("data", {}):
                self.access_token = data["data"]["access_token"]
                self.refresh_token = data["data"].get("refresh_token")
                self.user_info = data["data"].get("user", {})

                # Добавляем Bearer токен в постоянные заголовки сессии
                self.session.headers["Authorization"] = f"Bearer {self.access_token}"
                user_name = self.user_info.get("name") or self.user_info.get("email") or email
                return True, f"Успешный вход: {user_name} (ID: {self.user_info.get('id')})"
            else:
                err_msg = data.get("message") or data.get("errors") or "Неверный логин или пароль"
                return False, str(err_msg)
        except Exception as e:
            return False, f"Сетевой сбой при авторизации: {e}"

    # =========================================================================
    # 2. ПОИСК МЕРОПРИЯТИЯ
    # =========================================================================
    def search_activity(self, activity_name: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Ищет мероприятие по части или полному названию.
        Возвращает словарь мероприятия (включая 'id') или None.
        """
        url = f"{self.base_url}/api/rest/activity/"
        params = {
            "query": activity_name.strip(),
            "page": 1,
            "start": 0,
            "length": 25
        }

        try:
            resp = self.session.get(url, params=params, timeout=15)
            data = resp.json()

            if not data.get("success") or not data.get("data"):
                return None, f"Мероприятие '{activity_name}' не найдено на портале"

            activities: List[Dict[str, Any]] = data["data"]

            # 1. Попытка точного совпадения по имени
            for act in activities:
                if act.get("name", "").strip().lower() == activity_name.strip().lower():
                    return act, "Найдено точное совпадение"

            # 2. Иначе берем первый результат выдачи
            first_act = activities[0]
            return first_act, f"Найдено: '{first_act.get('name')}' (ID: {first_act.get('id')})"
        except Exception as e:
            return None, f"Ошибка при поиске мероприятия: {e}"

    # =========================================================================
    # 3. ПОИСК РЕБЕНКА ПО ФИО И СВЕРКА ДАТЫ РОЖДЕНИЯ + ПРОВЕРКА ПОДТВЕРЖДЕНИЯ
    # =========================================================================
    def find_kid(
        self, 
        fio: str, 
        birth_date: Optional[Any] = None,
        require_approved: bool = True
    ) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Ищет ребенка в реестре по ФИО.
        - Сверяет дату рождения 'birthday' (YYYY-MM-DD) для исключения тезок.
        - Проверяет 'is_approved' == True (подтвержденный аккаунт ребенка).
        """
        clean_fio = " ".join(str(fio).strip().split())
        url = f"{self.base_url}/api/activity/rest/kid"
        params = {
            "search[value]": clean_fio,
            "page": 1,
            "start": 0,
            "length": 25
        }

        try:
            resp = self.session.get(url, params=params, timeout=15)
            data = resp.json()

            if not data.get("success") or not data.get("data"):
                return None, f"Ребенок '{clean_fio}' не найден в реестре Навигатора"

            candidates: List[Dict[str, Any]] = data["data"]
            normalized_target_dob = self._normalize_dob(birth_date)

            # 1. Если задана дата рождения — ищем совпадения по дате
            dob_matches: List[Dict[str, Any]] = []
            if normalized_target_dob:
                for c in candidates:
                    c_dob = self._normalize_dob(c.get("birthday"))
                    if c_dob == normalized_target_dob:
                        dob_matches.append(c)

                if not dob_matches:
                    found_dobs = [str(c.get("birthday")) for c in candidates if c.get("birthday")]
                    return None, f"Найдено {len(candidates)} тезок, но ДР '{normalized_target_dob}' не совпала (в базе: {', '.join(found_dobs)})"
            else:
                dob_matches = candidates

            # 2. Проверка статуса 'is_approved' (подтвержденный аккаунт)
            if require_approved:
                approved_matches = [
                    c for c in dob_matches 
                    if c.get("is_approved") is True or str(c.get("is_approved")).lower() in ["true", "1"]
                ]

                if not approved_matches:
                    return None, f"Найден ребенок с совпавшей датой рождения ({normalized_target_dob}), но его аккаунт НЕ подтвержден ('is_approved': false)"

                # Берем подтвержденного
                selected = approved_matches[0]
                return selected, f"OK (аккаунт подтвержден, ДР {selected.get('birthday')} совпала)"

            selected = dob_matches[0]
            return selected, f"OK (найден кандидат, ДР: {selected.get('birthday')})"

        except Exception as e:
            return None, f"Сбой при поиске ребенка: {e}"

    # =========================================================================
    # 4. СОЗДАНИЕ ЗАЯВКИ (NavAdd)
    # =========================================================================
    def create_order(
        self, 
        activity_id: int, 
        event_datetime: str, 
        kid: Dict[str, Any]
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Создает заявку ребенка на мероприятие.
        event_datetime: 'ГГГГ-ММ-ДД ЧЧ:ММ:СС' или 'ДД.ММ.ГГГГ ЧЧ:ММ'
        """
        url = f"{self.base_url}/api/rest/activityOrder"
        norm_datetime = self._normalize_datetime(event_datetime)

        payload = {
            "data": {
                "activity_id": int(activity_id),
                "date": norm_datetime,
                "site_user_id": int(kid["site_user_id"]),
                "kid_id": kid["id"],
                "state": "initial"
            }
        }

        try:
            resp = self.session.post(url, json=payload, timeout=15)
            data = resp.json()

            if data.get("success"):
                order_id = data.get("data", {}).get("id")
                return True, f"Заявка создана (№ {order_id})", order_id
            else:
                err = data.get("message") or data.get("errors") or "Отказ сервера в создании заявки"
                return False, str(err), None
        except Exception as e:
            return False, f"Сетевая ошибка при создании заявки: {e}", None

    # =========================================================================
    # 5. ВЫБОРКА НОВЫХ ЗАЯВОК (NavConfirm)
    # =========================================================================
    def get_pending_orders(self, activity_id: int, limit: int = 100) -> Tuple[List[Dict[str, Any]], int]:
        """
        Получает пачку неподтвержденных заявок (state=initial) для выбранного мероприятия.
        Возвращает (список заявок, общее число оставшихся recordsFiltered).
        """
        url = f"{self.base_url}/api/rest/activityOrder"
        ext_filters = [
            {"property": "activity_id", "value": str(activity_id), "comparison": "eq"},
            {"property": "state", "value": ["initial"], "comparison": "in"}
        ]

        params = {
            "page": 1,
            "start": 0,
            "length": limit,
            "extFilters": json.dumps(ext_filters)
        }

        try:
            resp = self.session.get(url, params=params, timeout=20)
            data = resp.json()

            if data.get("success"):
                items = data.get("data", [])
                total_filtered = int(data.get("recordsFiltered", len(items)))
                return items, total_filtered
            return [], 0
        except Exception:
            return [], 0

    # =========================================================================
    # 6. ПОДТВЕРЖДЕНИЕ ЗАЯВКИ И УЧАСТИЯ (NavConfirm)
    # =========================================================================
    def approve_order(self, order_id: int, comment: str = "") -> Tuple[bool, str]:
        """
        Подтверждает участие ребенка по заявке (state -> approve).
        """
        url = f"{self.base_url}/api/setActivityOrderState"
        payload = {
            "data": {
                "approve_comment": comment,
                "id": int(order_id),
                "state": "approve"
            }
        }

        try:
            resp = self.session.post(url, json=payload, timeout=10)
            data = resp.json()
            if data.get("success"):
                return True, f"Заявка #{order_id} подтверждена"
            return False, data.get("message", "Отказ сервера в подтверждении")
        except Exception as e:
            return False, f"Ошибка при подтверждении #{order_id}: {e}"

    # =========================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ И НОРМАЛИЗАЦИЯ ДАТ
    # =========================================================================
    def _normalize_dob(self, raw_val: Optional[Any]) -> Optional[str]:
        """
        Приводит дату рождения к стандарту YYYY-MM-DD.
        Поддерживает:
        - datetime.date / datetime.datetime
        - '23.06.2005' (дд.мм.гггг)
        - '23.06.05' (дд.мм.гг, где 05 -> 2005)
        - '2005-06-23'
        - Разделители '.', '/', '-'
        """
        if raw_val is None:
            return None

        # 1. Если это уже объект datetime или date из Excel
        if hasattr(raw_val, "strftime"):
            return raw_val.strftime("%Y-%m-%d")

        s = str(raw_val).strip()
        if not s or s.lower() in ["none", "nan", "nat", "не указано", "null", ""]:
            return None

        # Убираем время, если оно попало в строку (например '2005-06-23 00:00:00')
        if " " in s:
            s = s.split(" ")[0]

        # Заменяем слэши и дефисы на точки для унификации
        s_clean = s.replace("/", ".").replace("-", ".")
        parts = [p.strip() for p in s_clean.split(".") if p.strip()]

        # Случай: ДД.ММ.ГГГГ или ДД.ММ.ГГ (3 части)
        if len(parts) == 3:
            # Проверим, идет ли первым год (ГГГГ.ММ.ДД) или день (ДД.ММ.ГГГГ)
            if len(parts[0]) == 4:  # ГГГГ.ММ.ДД
                year = parts[0]
                month = parts[1].zfill(2)
                day = parts[2].zfill(2)
                return f"{year}-{month}-{day}"
            else:  # ДД.ММ.ГГ или ДД.ММ.ГГГГ
                day = parts[0].zfill(2)
                month = parts[1].zfill(2)
                year_part = parts[2]

                # Если год указан 2 цифрами (дд.мм.гг, например 05 -> 2005)
                if len(year_part) == 2:
                    y_num = int(year_part)
                    # Если число <= 35, считаем 2000-е годы (2000..2035), иначе 1900-е
                    full_year = 2000 + y_num if y_num <= 35 else 1900 + y_num
                    year = str(full_year)
                elif len(year_part) == 4:
                    year = year_part
                else:
                    return None

                return f"{year}-{month}-{day}"

        return None

    def _normalize_datetime(self, dt_str: str) -> str:
        """Приводит дату-время к формату YYYY-MM-DD HH:MM:SS."""
        s = dt_str.strip()
        # Если формат "DD.MM.YYYY HH:MM"
        if "." in s and " " in s:
            date_part, time_part = s.split(" ", 1)
            parts = date_part.split(".")
            time_clean = time_part if len(time_part.split(":")) == 3 else f"{time_part}:00"
            return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)} {time_clean}"
        # Если уже "YYYY-MM-DD ..."
        if len(s.split(":")) == 2:
            return f"{s}:00"
        return s

    # =========================================================================
    # 7. ОБРАБОТКА EXCEL ТАБЛИЦЫ (NavAdd ПАКЕТНЫЙ РЕЖИМ)
    # =========================================================================
    @staticmethod
    def read_excel_records(file_path: str) -> List[Dict[str, Any]]:
        """
        Универсально считывает список детей из Excel (.xlsx, .xls).
        Автоматически определяет колонки по заголовкам (ФИО, Дата рождения)
        или берет первую (колонка A) и вторую (колонка B).
        """
        records: List[Dict[str, Any]] = []

        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active

            fio_col_idx = 1
            dob_col_idx = 2
            start_row = 2

            # Проверяем заголовки в первой строке
            header_row = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
            for col_idx, h in enumerate(header_row, start=1):
                if not h:
                    continue
                h_str = str(h).strip().lower()
                if "фио" in h_str or "ребенок" in h_str or "фамилия" in h_str:
                    fio_col_idx = col_idx
                elif "рожд" in h_str or "дат" in h_str or "др" in h_str:
                    dob_col_idx = col_idx

            for row_idx in range(start_row, sheet.max_row + 1):
                fio_val = sheet.cell(row_idx, fio_col_idx).value
                dob_val = sheet.cell(row_idx, dob_col_idx).value

                if fio_val and str(fio_val).strip():
                    records.append({
                        "row": row_idx,
                        "fio": str(fio_val).strip(),
                        "dob": dob_val
                    })
            wb.close()
            return records

        except ImportError:
            # Резервный вариант через pandas
            try:
                import pandas as pd
                df = pd.read_excel(file_path)
                fio_col = df.columns[0]
                dob_col = df.columns[1] if len(df.columns) > 1 else None

                for col in df.columns:
                    c_str = str(col).strip().lower()
                    if "фио" in c_str or "ребенок" in c_str or "фамилия" in c_str:
                        fio_col = col
                    elif "рожд" in c_str or "дат" in c_str or "др" in c_str:
                        dob_col = col

                for idx, row in df.iterrows():
                    fio_val = row[fio_col]
                    dob_val = row[dob_col] if dob_col else None
                    if pd.notna(fio_val) and str(fio_val).strip():
                        records.append({
                            "row": idx + 2,
                            "fio": str(fio_val).strip(),
                            "dob": dob_val if pd.notna(dob_val) else None
                        })
                return records
            except ImportError:
                raise ImportError(
                    "Для чтения Excel-файлов установите библиотеку openpyxl:\n"
                    "pip install openpyxl"
                )

    def process_excel(
        self,
        excel_path: str,
        activity_id: int,
        event_datetime: str,
        delay_seconds: float = 0.3,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Dict[str, Any]:
        """
        Пакетная запись детей из Excel таблицы на мероприятие:
        1. Читает строки (ФИО + Дата рождения).
        2. Нормализует ДР (дд.мм.гггг, дд.мм.гг).
        3. Ищет ребенка в реестре, проверяя 'is_approved' == True и совпадение ДР.
        4. Создает заявку на мероприятие (POST /api/rest/activityOrder).
        5. Возвращает подробный отчет.
        """
        records = self.read_excel_records(excel_path)
        log_msgs: List[str] = []

        def log(msg: str):
            print(msg)
            log_msgs.append(msg)
            if progress_callback:
                progress_callback(msg)

        log(f"\n[EXCEL] Загружено записей из файла: {len(records)}")
        log(f"[EXCEL] ID мероприятия: {activity_id}, Дата/время: {event_datetime}")
        log("-" * 65)

        stats = {
            "total": len(records),
            "success": 0,
            "skipped_unapproved": 0,
            "skipped_dob_mismatch": 0,
            "skipped_not_found": 0,
            "errors": 0,
            "details": []
        }

        for item in records:
            row_num = item["row"]
            fio = item["fio"]
            raw_dob = item["dob"]
            norm_dob = self._normalize_dob(raw_dob)

            log(f"\n[{row_num}] Обработка: {fio} (ДР в табл: {raw_dob} -> {norm_dob or 'нет'})")

            # Поиск ребенка с проверкой подтверждения и ДР
            kid, search_msg = self.find_kid(fio=fio, birth_date=norm_dob, require_approved=True)

            if not kid:
                log(f"    ✗ ПРОПУСК: {search_msg}")
                if "is_approved" in search_msg.lower():
                    stats["skipped_unapproved"] += 1
                elif "тезок" in search_msg.lower() or "не совпала" in search_msg.lower():
                    stats["skipped_dob_mismatch"] += 1
                else:
                    stats["skipped_not_found"] += 1

                stats["details"].append({
                    "row": row_num,
                    "fio": fio,
                    "dob": norm_dob,
                    "status": "skipped",
                    "reason": search_msg
                })
                continue

            # Ребенок успешно найден и подтвержден, создаем заявку
            ok, order_msg, order_id = self.create_order(
                activity_id=activity_id,
                event_datetime=event_datetime,
                kid=kid
            )

            if ok:
                stats["success"] += 1
                log(f"    ✓ УСПЕШНО: {order_msg} (Kid ID: {kid.get('id')})")
                stats["details"].append({
                    "row": row_num,
                    "fio": fio,
                    "dob": norm_dob,
                    "status": "success",
                    "order_id": order_id
                })
            else:
                stats["errors"] += 1
                log(f"    ✗ ОШИБКА СОЗДАНИЯ ЗАЯВКИ: {order_msg}")
                stats["details"].append({
                    "row": row_num,
                    "fio": fio,
                    "dob": norm_dob,
                    "status": "error",
                    "reason": order_msg
                })

            time.sleep(delay_seconds)

        log("\n" + "=" * 65)
        log("ИТОГИ ОБРАБОТКИ EXCEL:")
        log(f"Всего в таблице:            {stats['total']}")
        log(f"Успешно записано заявок:    {stats['success']}")
        log(f"Пропущено (не подтвержден): {stats['skipped_unapproved']}")
        log(f"Пропущено (ДР не совпала):  {stats['skipped_dob_mismatch']}")
        log(f"Пропущено (не найден):      {stats['skipped_not_found']}")
        log(f"Ошибок отправки:            {stats['errors']}")
        log("=" * 65)

        return stats


# =========================================================================
# БЛОК ДЛЯ ПРЯМОГО ЗАПУСКА ИЗ ТЕРМИНАЛА
# =========================================================================
if __name__ == "__main__":
    import os
    print("=" * 65)
    print(" КЛИЕНТ НАВИГАТОРА (REST API) - NavAdd & NavConfirm")
    print("=" * 65)

    client = NavigatorClient()

    # Запрашиваем логин и пароль
    email = input("\nВведите email/логин: ").strip()
    password = input("Введите пароль: ").strip()

    print("\n[1] Авторизация...")
    success, msg = client.login(email, password)
    print(f"Результат: {msg}")

    if not success:
        print("\n[ОШИБКА] Не удалось войти. Проверьте логин и пароль.")
        exit(1)

    print("\nВыберите режим работы:")
    print("  1 - Пакетная запись детей из Excel (NavAdd)")
    print("  2 - Массовое подтверждение заявок на мероприятие (NavConfirm)")
    print("  3 - Быстрый тест поиска ребенка по ФИО и ДР")

    choice = input("\nВаш выбор (1/2/3): ").strip()

    if choice == "1":
        excel_file = input("\nВведите путь к файлу Excel (по умолчанию 'list.xlsx'): ").strip()
        if not excel_file:
            excel_file = "list.xlsx"

        if not os.path.exists(excel_file):
            print(f"[ОШИБКА] Файл '{excel_file}' не найден в текущей папке!")
            exit(1)

        event_name = input("Введите название мероприятия: ").strip()
        act, act_msg = client.search_activity(event_name)
        if not act:
            print(f"[ОШИБКА] Мероприятие не найдено: {act_msg}")
            exit(1)

        act_id = int(act["id"])
        print(f"-> Найдено мероприятие: '{act.get('name')}' (ID: {act_id})")

        event_dt = input("Введите дату и время мероприятия (например '2026-08-31 11:00:00'): ").strip()
        if not event_dt:
            event_dt = "2026-08-31 11:00:00"

        print("\nЗапуск пакетной обработки Excel таблицы...")
        client.process_excel(excel_path=excel_file, activity_id=act_id, event_datetime=event_dt)

    elif choice == "2":
        event_name = input("\nВведите название мероприятия для подтверждения заявок: ").strip()
        act, act_msg = client.search_activity(event_name)
        if not act:
            print(f"[ОШИБКА] Мероприятие не найдено: {act_msg}")
            exit(1)

        act_id = int(act["id"])
        print(f"-> Мероприятие: '{act.get('name')}' (ID: {act_id})")

        orders, total = client.get_pending_orders(act_id, limit=100)
        print(f"-> Неподтвержденных заявок в очереди: {total}")

        if total == 0:
            print("Новых заявок для подтверждения нет!")
        else:
            confirm = input(f"Подтвердить все {total} заявок? (y/n): ").strip().lower()
            if confirm in ["y", "д", "yes", "да"]:
                done = 0
                while True:
                    batch, remaining = client.get_pending_orders(act_id, limit=100)
                    if not batch or remaining == 0:
                        break
                    for ord_item in batch:
                        oid = ord_item["id"]
                        ok_appr, _ = client.approve_order(oid)
                        if ok_appr:
                            done += 1
                            print(f"  ✓ Заявка #{oid} подтверждена ({done}/{total})")
                        time.sleep(0.05)
                print(f"\n[ГОТОВО] Успешно подтверждено {done} заявок!")

    elif choice == "3":
        kid_fio = input("\nВведите ФИО ребенка (например 'Солодовникова Романа Александровна'): ").strip()
        dob_input = input("Введите дату рождения (например '23.06.2005' или '23.06.05'): ").strip()
        print(f"\nПоиск ребенка '{kid_fio}'...")
        kid, msg = client.find_kid(kid_fio, birth_date=dob_input, require_approved=True)
        print(f"Результат: {msg}")
        if kid:
            print(f"-> ID (kid_id):       {kid.get('id')}")
            print(f"-> ID родителя:       {kid.get('site_user_id')}")
            print(f"-> Дата рождения:     {kid.get('birthday')}")
            print(f"-> Статус аккаунта:   {kid.get('is_approved')} (Подтвержден)")


