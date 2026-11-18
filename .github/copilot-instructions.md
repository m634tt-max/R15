# Инструкции для GitHub Copilot

## Стек и окружение
- Python 3.10+ (Windows, пути через pathlib.Path).
- GUI: tkinter (+ttkbootstrap, tkinterdnd2) для утилит; PySide6 (Qt) для модульных приложений.
- Библиотеки уже установлены: openpyxl, pandas, numpy, pillow, pdfplumber, pypdf, PyPDF2,
  python-docx, reportlab, PySide6, requests, keyring, lxml, pydantic, python-dotenv,
  pdf2image, win32com, pyinstaller.
- Кодировка файлов: UTF-8. Комментарии и тексты сообщений — на русском.

## Архитектура
- Скрипты (tkinter): бизнес-логика отдельно от GUI. Долгие операции — в
  threading.Thread(daemon=True), отмена через threading.Event (проверять в циклах).
  Обновление GUI из потока — только через root.after(0, ...).
- Приложения (PySide6) — модульная архитектура:
  main.py; core/ (event_bus.py — синглтон EventBus; logger.py — AppLogger в БД/файл;
  data_store.py — единая точка доступа к SQLite; task_queue.py — фоновые задачи с отменой);
  ui/ (main_window.py, styles.py — тема QSS, widgets/);
  modules/ — каждый модуль: класс-наследник BaseModule, QWidget со своей логикой,
  регистрируется через register_module(instance).
- Конфигурации сохранять (JSON / QSettings); при передаче изменяемых словарей — copy.deepcopy.

## Обязательный стиль кода
- Аннотации типов на всех функциях: аргументы и возвращаемое значение
  (from typing import Optional, List, Dict, Tuple, Callable, Any, Set).
- Обработка ошибок: все исключения ловить, логировать с traceback.format_exc(),
  возвращать (success: bool, message: str). В GUI — messagebox.showerror.
- Функции не длиннее 150 строк. Не использовать глобальные переменные (кроме констант).
- Не зашивать пути жёстко (параметры, конфиги). Магические числа не допускать.
- Комментарии на русском, блоки разделять `# ────` / `# ════════`.
- Для данных-классов массового создания — __slots__.
- Пути — только pathlib.Path. tempfile.TemporaryDirectory() для временных файлов.

## GUI-паттерны (обязательные)
- tkinter: hover-кнопки (HBtn), поля с placeholder (MEntry/ModernEntry), Toast-уведомления
  (Toplevel, overrideredirect, правый нижний угол родителя, авто-скрытие через after),
  ttk.Treeview с tag_configure по статусам (ОК / ПРЕДУПР. / ОШИБКА), лог в
  scrolledtext.ScrolledText с кнопкой копирования, ttk.Progressbar с процентом и стилем
  через ttk.Style.
- tkinterdnd2 — опционально: try/except ImportError, флаг DND_AVAILABLE,
  DND_BASE = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk; события <<DragEnter>>/<<DragLeave>>;
  путь из данных — data.strip("{}").strip().
- На время долгой операции блокировать кнопки действий, активировать «Отмена»;
  в finally возвращать исходное состояние. По завершении — диалог «Открыть файл?»
  (messagebox.askyesno; os.startfile(path)).
- Цвета: пастельная светлая, Dracula или тёмная мягкая темы (константы в коде).
  Для openpyxl hex БЕЗ "#", для GUI — С "#": отдельные константы
  (XL_OK_H = "D4EDDA" / GUI_OK = "#D4EDDA" и т. п.).

## Excel (openpyxl)
- Запись/стилизация — openpyxl; чтение данных — pandas; чтение листов — openpyxl
  read_only=True + wb.close().
- Отчёты: заголовки, чередование строк, автоширина колонок (макс. 50),
  freeze_panes = "A4", легенда цветов, merge_cells при необходимости,
  высота строк ws.row_dimensions[rx].height = 15.
- Сводные отчёты: лист «Сводный отчёт» + лист «Только ошибки» (с форматированием).

## Работа с документами
- PDF текст: pdfplumber.open() → page.extract_text(). Штампы/слияние/разделение: pypdf
  (PdfWriter, add_outline_item) + reportlab; диапазоны страниц: all/even/odd/everyN/"1,3-5,10".
- DOCX: python-docx (doc.paragraphs). XLSX: openpyxl.
- Поиск фраз: режимы prefix/suffix/exact/phrase_end/regex, флаг re.IGNORECASE,
  контекст ±50 символов, выгрузка в Excel.

## Сборка
pyinstaller --onefile --windowed --icon=icon.ico --name="Название" main.py
(--hidden-import win32timezone для win32com, --hidden-import ttkbootstrap;
 ресурсы — через sys._MEIPASS).

## Запрещено
- time.sleep() в основном потоке.
- Долгие операции без возможности отмены.
- Операции ввода-вывода без обработки ошибок.
- Удаление существующих тестов (только добавление новых).

## Проверка перед сдачей
- Аннотации типов везде; прогресс-бар и лог в GUI; поток daemon=True и отмена;
  блокировка кнопок; диалог «Открыть файл?»; комментарии на русском; импорты сгруппированы.

## Что делать при неоднозначности
- Если задача допускает разные трактовки или есть риск сломать существующий код —
  не угадывать: задать уточняющий вопрос в комментарии к issue/PR.
- Не менять публичный API и конфиги (.env, pyproject) без явного запроса.
