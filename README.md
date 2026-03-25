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
/mnt/test/
└── 2026/02/25/08/
    ├── mail/            ← смотрим только сюда
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
   │   ├── start_time   → Таблица 1: Бошланиш вақти
   │   ├── stop_time    → Таблица 1: Тугаш вақти
   │   ├── sender       → Таблица 1: Ким
   │   ├── receiver     → Таблица 1: Кимга
   │   ├── order_code   → Таблица 0: префикс перед каждым вложением
   │   ├── user_id      → Таблица 3: Бажарувчи
   │   └── filename     → имя .eml файла
   │
   ├─ parse_eml()  (если .eml найден)
   │   ├── текст письма (text/plain → text/html)
   │   └── вложения: имя + размер в байтах
   │
   └─ fill_document()
       ├── Таблица 0 (шапка)
       │   └── вложения: "({order_code}) {имя} qwerty {размер} байт"
       │       нет order_code → "()" перед именем файла
       │       нет вложений → ячейки очищаются полностью
       ├── Таблица 1 (данные)
       │   └── время начала/окончания, отправитель, получатель
       ├── Таблица 3 (подпись)
       │   └── Бажарувчи: {user_id}
       ├── текст письма — под таблицей 1
       └── сохранить abc.docx (атомарно через .tmp)
```

**Ожидание EML файла:**
Если `.eml` ещё не появился — скрипт ждёт до 60 секунд
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
| `WATCH_ROOT` | `/mnt/test` | Корень дерева директорий для мониторинга |
| `TARGET_DIR` | `mail` | Имя папки из которой брать файлы |
| `MAX_WORKERS` | `10` | Параллельных обработчиков |
| `SCAN_INTERVAL` | `300` | Интервал периодической проверки (сек) |
| `LOG_PATH` | `/var/log/eml-watcher.log` | Путь к лог файлу |
| `LOG_MAX_BYTES` | `10 MB` | Максимальный размер одного лог файла |
| `LOG_BACKUP_COUNT` | `5` | Количество хранимых лог файлов (итого до 50 MB) |

## Настройки (process.py)

| Параметр | По умолчанию | Описание |
|---|---|---|
| `TEMPLATE_PATH` | `template.docx` рядом со скриптом | Путь к шаблону Word |
| `EML_WAIT_TIMEOUT` | `60` | Максимум секунд ожидания EML файла |
| `EML_WAIT_INTERVAL` | `2` | Как часто проверять наличие EML (сек) |

## Форматирование текста в документе

Весь текст вставляемый скриптом: **Times New Roman, 14pt**.
Применяется к: вложениям (Таблица 0), данным (Таблица 1), тексту письма, Бажарувчи (Таблица 3).

---

## Установка и деплой

### 1. Зависимости

```bash
cd /opt/word
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Если сервер без интернета — скачать пакеты на другой машине:
```bash
pip download -r requirements.txt -d ./packages \
  --platform manylinux2014_x86_64 \
  --python-version 3.11 \
  --only-binary=:all:
```
Перенести папку `packages` на сервер и установить:
```bash
pip install --no-index --find-links ./packages -r requirements.txt
```

### 2. Лимит inotify watches (если дерево директорий большое)

```bash
echo fs.inotify.max_user_watches=524288 >> /etc/sysctl.conf
sysctl -p
```

### 3. systemd сервис

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

# Сколько файлов уже обработано
find /mnt/test -name "*.docx" | wc -l

# Сколько ещё не обработано
find /mnt/test -path "*/mail/*.json" | while read f; do
  [ ! -f "${f%.json}.docx" ] && echo "$f"
done | wc -l
```

---

## Ручной запуск (для тестирования)

```bash
# Обработать один файл
/opt/word/venv/bin/python3 /opt/word/process.py /path/to/file.json

# Запустить демон в консоли (Ctrl+C для остановки)
/opt/word/venv/bin/python3 /opt/word/daemon.py
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
