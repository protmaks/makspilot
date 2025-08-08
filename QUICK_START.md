# 🚀 Быстрое руководство по управлению версиями

## 🎉 Текущая версия: v 5.0 Beta

### ✨ Что нового в v 5.0 Beta:
- 📊 **Excel экспорт** с продвинутым форматированием
- 🎨 **Цветовое кодирование** различий в экспортированных файлах
- 🔧 **Улучшенные уведомления** об успешном экспорте
- 🌐 **Поддержка всех языков** для функции экспорта

## Что создано:

1. **`version.json`** - центральный файл с версией
2. **`update_version.py`** - автоматическое обновление всех файлов
3. **`.github/workflows/release.yml`** - автоматизация GitHub релизов
4. **`javascript/version.js`** - динамическое отображение версии

## Как обновить версию:

### Простой способ:
```bash
cd /путь/к/проекту
python3 update_version.py
```

### Через GitHub:
1. Идите в Actions → "Create Release and Update Version"
2. Нажмите "Run workflow"  
3. Введите новую версию (например: "v 5.1 Release")
4. Выберите тип релиза (beta/release/hotfix)
5. Нажмите "Run workflow"

## Что происходит автоматически:

✅ Обновляется `version.json`  
✅ Обновляются все HTML файлы со страницами сравнения  
✅ Создается Git тег  
✅ Создается GitHub Release  
✅ Генерируется changelog  

## Текущие файлы с версией:

- `/compare/index.html`
- `/ar/compare/index.html`  
- `/de/compare/index.html`
- `/es/compare/index.html`
- `/ja/compare/index.html`
- `/pl/compare/index.html`
- `/pt/compare/index.html`
- `/ru/compare/index.html`
- `/zh/compare/index.html`

Все готово к использованию! 🎉
