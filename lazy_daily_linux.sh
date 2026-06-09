#!/bin/bash

# В .bashrc добавить alias lazyd="$HOME/путь_до_папки"

# Папка загрузок
DOWNLOADS="$HOME/Загрузки"

# Папка, где будут создаваться задачи
TARGET_BASE="$HOME/Daily"

# Папка логов
LOG_FILE="$HOME/.lazy_daily_processed.log"
touch "$LOG_FILE"

LAST=""
LAST_SIZE=0
SKIP_NOTIFIED=""  # запоминаем, для каких архивов уже вывели предупреждение

# Функция для очистки при выходе
cleanup() {
    printf "Скрипт остановлен \n"

    if [ -n "$LAST" ]; then
        echo "Последний обработанный архив: $LAST"
    fi
    exit 0
}
trap cleanup SIGINT SIGTERM

while true; do
    # Поиск самого нового zip-архива
    ARCHIVE=$(find "$DOWNLOADS" -maxdepth 1 -name "*.zip" -printf "%T@ %p\n" \
        | sort -nr | head -n 1 | cut -d' ' -f2-)

    if [ -n "$ARCHIVE" ] && [ "$ARCHIVE" != "$LAST" ]; then

        # Проверка по логу
        if grep -qxF "$ARCHIVE" "$LOG_FILE"; then
            # Выводим сообщение только один раз для этого архива
            if [[ "$SKIP_NOTIFIED" != "$ARCHIVE" ]]; then
                echo ""
                echo "Архив $(basename "$ARCHIVE") уже обработан, пропускаем"
                echo ""
                SKIP_NOTIFIED="$ARCHIVE"
            fi
            sleep 3
            continue
        fi

        # Новый архив → сбрасываем флаг уведомления
        SKIP_NOTIFIED=""

        SIZE=$(stat -c%s "$ARCHIVE")

        # Проверяем что файл перестал скачиваться
        if [ "$SIZE" = "$LAST_SIZE" ]; then

            # Определяем номер папки ЗадачаN
            i=1
            while [ -d "$TARGET_BASE/Задача_$i" ]; do
                i=$((i+1))
            done

            TASK_DIR="$TARGET_BASE/Задача_$i"
            mkdir -p "$TASK_DIR"
            echo "$(basename "$TASK_DIR") создана"

            # Распаковка архива
            unzip -q "$ARCHIVE" -d "$TASK_DIR"

            # Удаляем архив после успешной распаковки
            rm "$ARCHIVE"
            echo "$(basename "$ARCHIVE") удалён"
            echo "$ARCHIVE" >> "$LOG_FILE"
            echo "Архив добавлен в историю"

            # Ищем geojson файлы
            mapfile -t GEOJSON_FILES < <(find "$TASK_DIR" -type f -name "*.geojson")
            COUNT=${#GEOJSON_FILES[@]}
            printf "Найдено файлов формата geojson : %d" "$COUNT"
            echo ""

            # Открываем папку
            dolphin "$TASK_DIR" &

            # Ищем docx файл
            DOCX=$(find "$TASK_DIR" -name "*.docx" | head -n 1)

            if [ -n "$DOCX" ]; then
                mostechoffice --nologo "$DOCX" &
            fi

            # Запоминаем обработанный архив
            LAST="$ARCHIVE"
            LAST_SIZE=0

        else
            LAST_SIZE="$SIZE"
        fi
    else
        # Если нет нового архива — сбрасываем флаг уведомления
        SKIP_NOTIFIED=""
    fi

    sleep 3
done
