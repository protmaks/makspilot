# MaxPilot Version Management System

Эта система позволяет централизованно управлять версией приложения MaxPilot на всех страницах сайта.

## Файлы системы

1. **`version.json`** - Центральный файл конфигурации версии
2. **`update_version.py`** - Python скрипт для автоматического обновления версии
3. **`javascript/version.js`** - JavaScript для динамического отображения версии
4. **`VERSION_README.md`** - Этот файл с инструкциями

## Структура version.json

```json
{
  "version": "v 4.0 Beta",
  "releaseDate": "2025-08-08",
  "description": "Application version configuration",
  "changelog": {
    "v 4.0 Beta": [
      "Enhanced Excel export functionality",
      "Improved file comparison algorithms",
      "Better UI/UX for comparison results",
      "Multi-language support"
    ]
  }
}
```

## Как обновить версию

### Метод 1: Использование Python скрипта (Рекомендуется)

```bash
# Перейти в корневую папку проекта
cd /path/to/maxpilot

# Запустить скрипт обновления
python3 update_version.py
```

Скрипт выполнит:
1. Покажет текущую версию
2. Запросит новую версию
3. Обновит все HTML файлы
4. Обновит `version.json`

### Метод 2: Ручное обновление

1. Отредактировать `version.json`
2. Вручную найти и заменить версию во всех HTML файлах
3. Убедиться, что все языковые версии обновлены

## Использование в HTML

### Статический способ (текущий)
```html
<h1>Compare Data Between Excel and CSV Files - v 4.0 Beta</h1>
```

### Динамический способ (опционально)
```html
<!-- Подключить версионный JavaScript -->
<script src="javascript/version.js"></script>

<!-- Использовать data-атрибуты -->
<h1>Compare Data Between Excel and CSV Files - <span data-version="full"></span></h1>
<p>Version: <span data-version="number"></span></p>
<p>Release Date: <span data-version="date"></span></p>
```

## Интеграция с GitHub Releases

### Автоматизация через GitHub Actions

Создайте файл `.github/workflows/release.yml`:

```yaml
name: Update Version and Create Release

on:
  workflow_dispatch:
    inputs:
      version:
        description: 'New version (e.g., v 4.1 Beta)'
        required: true
        type: string

jobs:
  update-version:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.x'
      
      - name: Update version
        run: |
          echo "${{ github.event.inputs.version }}" | python3 update_version.py
      
      - name: Commit changes
        run: |
          git config --local user.email "action@github.com"
          git config --local user.name "GitHub Action"
          git add .
          git commit -m "Update version to ${{ github.event.inputs.version }}"
          git push
      
      - name: Create Release
        uses: actions/create-release@v1
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        with:
          tag_name: ${{ github.event.inputs.version }}
          release_name: Release ${{ github.event.inputs.version }}
          draft: false
          prerelease: true
```

### Ручное создание релиза

1. Обновить версию через `update_version.py`
2. Закоммитить изменения
3. Создать тег: `git tag v4.0-beta`
4. Отправить тег: `git push origin v4.0-beta`
5. Создать релиз на GitHub с этим тегом

## Файлы, которые обновляются автоматически

- `index.html`
- `compare/index.html`
- `ar/compare/index.html`
- `de/compare/index.html`
- `es/compare/index.html`
- `ja/compare/index.html`
- `pl/compare/index.html`
- `pt/compare/index.html`
- `ru/compare/index.html`
- `zh/compare/index.html`

## Проверка обновления

После обновления версии проверьте:

1. Все языковые версии страниц сравнения
2. Главную страницу
3. Файл `version.json`
4. Отсутствие ошибок в консоли браузера

## Пример использования

```bash
# Обновление с v 4.0 Beta на v 4.1 Release
python3 update_version.py
# Введите: v 4.1 Release
# Подтвердите: y

# Результат: все HTML файлы и version.json обновлены
```

## Troubleshooting

**Проблема**: Скрипт не находит файлы
**Решение**: Убедитесь, что запускаете скрипт из корневой папки проекта

**Проблема**: Версия не обновляется на сайте
**Решение**: Очистите кеш браузера или проверьте правильность путей к файлам

**Проблема**: JavaScript не загружает version.json
**Решение**: Убедитесь, что файл доступен по HTTP (не только локально)
