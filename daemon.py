import os
import signal
import time
import logging
import threading
import queue
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from inotify_simple import INotify, flags

# --- Настройки ---
WATCH_ROOT    = "/opt/test"  # корень мониторинга
TARGET_DIR    = "qwerty"     # имя папки которая нас интересует
MAX_WORKERS   = 10           # параллельных воркеров
SCAN_INTERVAL = 300          # периодическая проверка каждые N секунд (5 мин)
LOG_PATH      = "/var/log/eml-watcher.log"

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)

# Импортируем после настройки логгера
from process import process_json

# --- Состояние ---
task_queue     = queue.Queue()
in_flight      = set()
in_flight_lock = threading.Lock()
shutdown_event = threading.Event()


def is_processed(json_path: Path) -> bool:
    """Файл считается обработанным если рядом уже есть .docx"""
    return json_path.with_suffix(".docx").exists()


def try_enqueue(json_path: Path):
    """Добавить в очередь если файл ещё не обработан и не в работе."""
    key = str(json_path)
    with in_flight_lock:
        if is_processed(json_path) or key in in_flight:
            return
        in_flight.add(key)
    task_queue.put(json_path)
    logging.info(f"Queued: {json_path} | queue size: {task_queue.qsize()}")


def scan_directory(label: str):
    """Найти все необработанные .json в папках TARGET_DIR и поставить в очередь."""
    found = 0
    for json_file in Path(WATCH_ROOT).rglob(f"{TARGET_DIR}/*.json"):
        if not is_processed(json_file):
            try_enqueue(json_file)
            found += 1
    if found:
        logging.info(f"[{label}] Found {found} unprocessed files")
    else:
        logging.info(f"[{label}] No unprocessed files")


def watcher_thread():
    """
    Следим за всем деревом WATCH_ROOT через inotify.
    При появлении новой директории — добавляем на неё watch.
    Обрабатываем только .json файлы из папок с именем TARGET_DIR.
    """
    inotify   = INotify()
    watch_map = {}  # wd → путь к директории

    def add_watch(path: str):
        try:
            wd = inotify.add_watch(path, flags.CLOSE_WRITE | flags.CREATE)
            watch_map[wd] = path
        except Exception as e:
            logging.warning(f"Cannot add watch for {path}: {e}")

    # Добавить watch на все уже существующие директории
    for dirpath, _, _ in os.walk(WATCH_ROOT):
        add_watch(dirpath)
    logging.info(f"inotify: watching {len(watch_map)} dirs under {WATCH_ROOT}")

    while not shutdown_event.is_set():
        for event in inotify.read(timeout=1000):  # timeout 1s чтобы проверять shutdown
            parent_path = watch_map.get(event.wd, "")
            if not parent_path:
                continue

            filepath    = Path(parent_path) / event.name
            event_flags = flags.from_mask(event.mask)

            # Новая директория — добавить watch чтобы видеть файлы внутри
            if flags.CREATE in event_flags and flags.ISDIR in event_flags:
                add_watch(str(filepath))
                logging.info(f"New dir, added watch: {filepath}")

            # Новый .json файл — обработать только если лежит в TARGET_DIR
            elif (flags.CLOSE_WRITE in event_flags
                  and event.name.endswith(".json")
                  and Path(parent_path).name == TARGET_DIR):
                logging.info(f"inotify event: {filepath}")
                try_enqueue(filepath)


def scanner_thread():
    """Периодически сканировать директорию как страховка от пропущенных файлов."""
    while not shutdown_event.is_set():
        # Ждём SCAN_INTERVAL секунд мелкими шагами, чтобы реагировать на shutdown
        for _ in range(SCAN_INTERVAL):
            if shutdown_event.is_set():
                return
            time.sleep(1)
        scan_directory("periodic")


def worker(json_path: Path):
    """Обработать один файл. Снять с in_flight после завершения."""
    try:
        process_json(json_path)
    except Exception as e:
        logging.error(f"Worker error {json_path}: {e}")
    finally:
        with in_flight_lock:
            in_flight.discard(str(json_path))


def main():
    logging.info("=" * 50)
    logging.info(
        f"Daemon started | root={WATCH_ROOT} | target={TARGET_DIR} | workers={MAX_WORKERS}"
    )

    # --- Graceful shutdown ---
    def on_shutdown(signum, frame):
        logging.info(f"Signal {signum} received — graceful shutdown initiated")
        shutdown_event.set()
        # Кладём None в очередь чтобы разблокировать главный цикл
        task_queue.put(None)

    signal.signal(signal.SIGTERM, on_shutdown)
    signal.signal(signal.SIGINT,  on_shutdown)

    # 1. Обработать файлы которые появились пока сервис не работал
    scan_directory("startup")

    # 2. Запустить inotify поток
    threading.Thread(target=watcher_thread, daemon=True, name="inotify").start()

    # 3. Запустить поток периодической проверки
    threading.Thread(target=scanner_thread, daemon=True, name="scanner").start()

    # 4. Пул воркеров — executor.shutdown(wait=True) вызывается автоматически
    #    при выходе из with-блока, дожидаясь завершения всех активных задач
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        while not shutdown_event.is_set():
            try:
                json_path = task_queue.get(timeout=1)
                if json_path is None:  # сигнал остановки
                    break
                executor.submit(worker, json_path)
            except queue.Empty:
                continue

    logging.info("All workers finished. Daemon stopped.")


if __name__ == "__main__":
    main()
