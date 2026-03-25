# rustypaste-windows-tool

**VBS script for fast file uploads to RustyPaste server**


**VBS скрипт для быстрой загрузки файлов\текста\изображения\URL-ссылок на RustyPaste сервер**


## English | [Русский](#русский)

### Settings

**Edit via any text editor (Windows-1251 codepage recommended).**

- Tokens for adding and deleting files
- Custom server URL
- Date/time prefix option
- Default file lifetime
- Language support (English and Russian pre-configured; easy to add more)

### File Upload Methods

1. **Drag & drop** files directly onto the VBS script
2. **Copy file to clipboard** (Ctrl+C or right-click → Copy) and run the script
3. **Copy file path** to clipboard (e.g., `C:\Users\User\Desktop\somefile.txt`) and run
4. **Copy any text** to clipboard → script creates and uploads a `.txt` file
5. **PrintScreen** (screenshot to clipboard) → script saves as `.png` and uploads
6. **Copy URL** to clipboard → create short link (YES) or download/upload file (NO)

### File Lifetime Formats

- `60d` — 60 days
- `1h` — 1 hour
- `2h` — 2 hours
- `12min` — 12 minutes
- `120` — 120 minutes (2 hours)

### Result

After upload, the **file URL** is automatically copied to your clipboard.

---

## [English](#english) | Русский  {#русский}

### Настройки

**Редактируйте в любом текстовом редакторе (рекомендуется кодировка Windows-1251).**

- Токены для добавления и удаления файлов
- URL вашего сервера
- Опция префикса дата\время
- Срок жизни файлов по умолчанию
- Поддержка языков (английский и русский уже есть; легко добавить другие)

### Способы загрузки файлов

1. **Drag & drop** — перетащите файлы на VBS-скрипт
2. **Копировать файл** в буфер обмена (Ctrl+C или ПКМ → Копировать) и запустите скрипт
3. **Скопировать путь к файлу** в буфер (например, `C:\Users\User\Desktop\somefile.txt`) и запустите скрипт
4. **Скопировать любой текст** в буфер и запустите скрипт → создастся `.txt` файл и загрузится
5. **PrintScreen** (скриншот в буфер) и запустите скрипт → сохранится как `.png` и загрузится
6. **Скопировать URL** в буфер и запустите скрипт → короткая ссылка (ДА) или загрузка файла по ссылке (НЕТ)

### Форматы срока жизни

- `60d` — 60 дней
- `1h` — 1 час
- `2h` — 2 часа
- `12min` — 12 минут
- `120` — 120 минут (2 часа)

### Результат

После загрузки **ссылка на файл** автоматически копируется в буфер обмена.

---

### PS

- Закрепите ярлык на этот VBS-скрипт в меню «Пуск» или на панели задач
- Поместите ярлык в меню ПКМ «Отправить →» (папка открывается через `shell:sendto`)
- Можно перетаскивать файлы прямо на ярлык скрипта или сам скрипт
