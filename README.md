# EML Watcher — автоматическая генерация Word документов

Демон мониторит директорию, и как только появляется новый `.json` файл —
парсит его, открывает связанный `.eml` файл, и генерирует заполненный `.docx`.

---

## Структура проекта

```
/opt/word/
├── daemon.py        # демон — точка входа, мониторинг + очередь
├── process.py       # логика обработки одного файла
├── template.docx    # шаблон Word документа
├── requirements.txt # зависимости Python
└── README.md
```

Файлы которые обрабатываются:
```
/opt/test/
└── 2026/02/25/08/
    ├── qwerty/          ← смотрим только сюда
    │   ├── abc.json
    │   ├── abc.eml
    │   └── abc.docx     ← результат
    └── other/           ← игнорируем
```

---

## Как это работает

### daemon.py — три уровня обнаружения файлов

```
┌─────────────────────────────────────────────────────┐
│ 1. При старте: scan_directory("startup")             │
│    Находит все .json без .docx — страховка от        │
│    файлов которые пришли пока сервис не работал      │
│                                                      │
│ 2. inotify поток                                     │
│    Слушает ядро Linux — мгновенная реакция           │
│    на новые файлы без нагрузки на CPU                │
│                                                      │
│ 3. scanner поток (каждые 5 мин)                      │
│    Периодическая проверка — финальная страховка      │
└────────────────────┬────────────────────────────────┘
                     │ json_path
                     ▼
              task_queue (RAM)
              хранит только пути к файлам (~40 байт каждый)
                     │
                     ▼
        ┌────────────────────────┐
        │  ThreadPoolExecutor    │
        │  max_workers = 10      │
        │                        │
        │  worker1 → file1.json  │
        │  worker2 → file2.json  │  ← обрабатываются параллельно
        │  ...                   │
        │  worker10→ file10.json │
        └────────────────────────┘
        file11...fileN ждут в очереди как строки
```

**in_flight** — множество файлов которые сейчас в работе.
Защищает от двойной обработки если inotify и periodic scan
увидят один файл одновременно.

**Файл считается обработанным** если рядом существует `.docx` с тем же именем.
Это позволяет корректно восстанавливаться после перезапуска.

### process.py — обработка одного файла

```
abc.json
   │
   ├─ parse_json()
   │   ├── start_time  → таблица1/Test1, таблица2/Test5
   │   ├── stop_time   → таблица1/Test2, таблица2/Test6
   │   ├── test4       → таблица1/Test3, таблица2/Test7
   │   ├── test3       → таблица1/Test4, таблица2/Test8
   │   └── filename    → имя .eml файла
   │
   ├─ parse_eml()  (если .eml найден)
   │   ├── текст письма (text/plain → text/html)
   │   └── вложения: имя + размер
   │
   └─ fill_document()
       ├── вставить текст и вложения над таблицами
       ├── заполнить таблицу 1
       ├── заполнить таблицу 2
       └── сохранить abc.docx
```

**Ожидание EML файла:**
Если `.eml` ещё не появился в директории — скрипт ждёт до 60 секунд
(проверяя каждые 2 секунды). Если не дождался — генерирует `.docx`
без текста письма и логирует предупреждение.

### Graceful shutdown

При `systemctl stop` демон получает `SIGTERM`:
1. Перестаёт принимать новые задачи
2. Дожидается завершения всех активных воркеров
3. Только после этого останавливается

Это гарантирует что `.docx` файлы не будут записаны частично.

---

## Настройки (daemon.py)

| Параметр | По умолчанию | Описание |
|---|---|---|
| `WATCH_ROOT` | `/opt/test` | Корень дерева директорий для мониторинга |
| `TARGET_DIR` | `qwerty` | Имя папки из которой брать файлы |
| `MAX_WORKERS` | `10` | Параллельных обработчиков |
| `SCAN_INTERVAL` | `300` | Интервал периодической проверки (сек) |
| `LOG_PATH` | `/var/log/eml-watcher.log` | Путь к лог файлу |

## Настройки (process.py)

| Параметр | По умолчанию | Описание |
|---|---|---|
| `EML_WAIT_TIMEOUT` | `60` | Максимум секунд ожидания EML файла |
| `EML_WAIT_INTERVAL` | `2` | Как часто проверять наличие EML (сек) |

---

## Установка и деплой

### 1. Зависимости

```bash
cd /opt/word
pip3 install -r requirements.txt
# или через venv:
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Лимит inotify watches (если дерево директорий большое)

```bash
echo fs.inotify.max_user_watches=524288 >> /etc/sysctl.conf
sysctl -p
```

### 4. systemd сервис

```bash
cat > /etc/systemd/system/eml-watcher.service << 'EOF'
[Unit]
Description=EML Watcher Daemon
After=network.target

[Service]
Type=simple
ExecStart=/opt/word/venv/bin/python3 /opt/word/daemon.py
Restart=always
RestartSec=5
TimeoutStopSec=120

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable eml-watcher
systemctl start eml-watcher
```

---

## Управление

```bash
# Запуск / остановка / статус
systemctl start eml-watcher
systemctl stop eml-watcher       # graceful — дождётся завершения задач
systemctl restart eml-watcher
systemctl status eml-watcher

# Логи в реальном времени
journalctl -u eml-watcher -f

# Лог файл
tail -f /var/log/eml-watcher.log
```

---

## Ручной запуск (для тестирования)

```bash
# Обработать один файл
python3 process.py /path/to/file.json

# Запустить демон в консоли (Ctrl+C для остановки)
python3 daemon.py
```

---

## Расход памяти

```
Очередь (10000 файлов-путей)  ≈ 400 KB
10 активных воркеров           ≈ 10 × 100 MB = ~1 GB

Итого: ~1 GB при MAX_WORKERS=10
```

Для уменьшения памяти — снизить `MAX_WORKERS`.
Для увеличения скорости — увеличить `MAX_WORKERS`.
