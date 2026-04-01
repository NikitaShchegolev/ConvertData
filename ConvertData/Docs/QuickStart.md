# Быстрый старт ConvertData

> **5-минутное руководство для быстрого начала работы**

## 🚀 Запуск за 3 шага

### Шаг 1: Подготовьте данные

Поместите Excel файлы с данными о соединениях в папку `EXCEL/`:

```
ConvertData/
└── EXCEL/
    ├── M_P.xls      ← ваш файл 1
    ├── M_BK.xlsx    ← ваш файл 2
    └── ...
```

**Требования к Excel файлам:**
- Формат: `.xls` или `.xlsx`
- Обязательные колонки: `Name`, `CONNECTION_CODE`, `ProfileBeam`
- Опционально: листы `geometry`, `bolts`, `weld`

---

### Шаг 2: Запустите конвертацию

```bash
ConvertData.exe
```

Вот и всё! Программа автоматически:
1. ✅ Прочитает все Excel файлы из `EXCEL/`
2. ✅ Применит справочник профилей из `EXCEL_Profile/ProfileBeam.xls`
3. ✅ Создаст JSON файлы
4. ✅ Объединит и обогатит данные
5. ✅ Устранит дубликаты
6. ✅ Экспортирует справочники

---

### Шаг 3: Получите результаты

После завершения вы найдёте:

```
ConvertData/
├── JSON_OUT/              ← Индивидуальные JSON файлы
│   ├── M_P.json
│   ├── M_BK.json
│   └── ...
├── JSON_All/              ← Объединённые результаты
│   ├── all.json                      ← все записи
│   ├── all_NotDuplicate.json         ← без дубликатов ✨
│   ├── CONNECTION_CODE_duplicates.txt
│   ├── CONNECTION_CODE_replacements.txt
│   └── ...
└── EXCEL_Profile_OUT/     ← Экспортированные справочники
    ├── Profile.json
    ├── CONNECTION_CODE_new.json
    └── NameConnections.json
```

**Основной результат:** `JSON_All/all_NotDuplicate.json`

---

## 📋 Частые сценарии

### Сценарий 1: Конвертировать один файл

```bash
ConvertData.exe path\to\your\file.xlsx
```

### Сценарий 2: Только создать JSON (без профилей)

```bash
ConvertData.exe 1
```

### Сценарий 3: Только применить профили к существующим JSON

```bash
ConvertData.exe 2
```

### Сценарий 4: Использовать свою колонку профиля

```bash
ConvertData.exe --profile-column=MyCustomProfileColumn
```

---

## 📄 Формат входных данных

### Минимальный Excel файл

**Main лист:**

| Name | CONNECTION_CODE | ProfileBeam | variable | Nt | My |
|------|----------------|-------------|----------|----|----|
| Балка-Колонна M | M-3 | 20Б1 | 1 | 100 | 50 |
| Балка-Колонна M | M-4 | 30Б1 | 1 | 200 | 80 |

**Опционально — лист geometry:**

| CODE | H | B | tp |
|------|---|---|----|
| M-3 | 300 | 150 | 10 |
| M-4 | 350 | 180 | 12 |

**Опционально — лист bolts:**

| CODE | F | Nb | e1 | d1 | d2 | Марка опорного столика |
|------|---|----|----|----|----|------------------------|
| M-3 | 20 | 4 | 50 | 60 | 120 | Т1 |
| M-4 | 24 | 6 | 60 | 70 | 140 | Т2 |

---

## 🔍 Проверка результатов

### Откройте `all_NotDuplicate.json`

```json
[
  {
    "Name": "Балка-Колонна M",
    "CONNECTION_CODE": "M-3",
    "variable": 1,
    "TableBrand": "Т1",
    
    "Geometry": {
      "Beam": {
        "ProfileBeam": "20Б1",
        "Beam_H": 200,
        "Beam_B": 100,
        "Beam_s": 5.6,
        "Beam_t": 8.5,
        ...
      },
      "Plate": {
        "Plate_H": 300,
        "Plate_B": 150,
        "Plate_t": 10
      }
    },
    
    "Bolts": {
      "DiameterBolt": { "F": 20 },
      "CountBolt": { "Bolts_Nb": 4 },
      "CoordinatesBolts": {
        "Y": { "Bolt1_e1": 50, ... },
        "X": { "d1": 60, "d2": 120 }
      }
    },
    
    "InternalForces": {
      "Nt": 100,
      "My": 50,
      ...
    }
  }
]
```

✅ **Успешно!** Данные сконвертированы.

---

## ⚠️ Частые ошибки

### Ошибка: "Cannot find CODE column"

**Причина:** В Excel нет колонки `CONNECTION_CODE`, `Code` или `Код`

**Решение:** 
1. Откройте Excel файл
2. Убедитесь, что есть колонка с одним из этих имён
3. Проверьте регистр (не важен, но название должно быть точным)

---

### Ошибка: "File not accessible"

**Причина:** Excel файл открыт в другой программе

**Решение:**
1. Закройте Excel
2. Запустите ConvertData снова

---

### Ошибка: COM Interop (.xls конвертация)

**Причина:** Microsoft Excel не установлен

**Решение:**
1. Установите Microsoft Excel
2. Или конвертируйте .xls → .xlsx вручную
3. Запустите ConvertData с .xlsx файлами

---

## 💡 Полезные советы

### 1. Проверьте дубликаты

После запуска откройте `CONNECTION_CODE_duplicates.txt`:

```
Найдено дубликатов: 3
M-3: 3 вхождения
M-5: 2 вхождения
```

Затем посмотрите `CONNECTION_CODE_replacements.txt`:

```
M-3 => M-6
M-3 => M-7
M-5 => M-8
```

### 2. Используйте справочник профилей

Поместите `ProfileBeam.xls` в `EXCEL_Profile/`:

**ProfileI лист:**

| Profile | H | B | s | t | A | P | Iz | Iy | ... |
|---------|---|---|---|---|---|---|----|----|-----|
| 20Б1 | 200 | 100 | 5.6 | 8.5 | 30.6 | 24.0 | 1840 | 155 | ... |
| 30Б1 | 300 | 135 | 6.5 | 9.5 | 42.7 | 33.5 | 5010 | 247 | ... |

ConvertData автоматически применит эти данные!

### 3. Обогащение неполных записей

Если у вас несколько записей с одинаковым `CONNECTION_CODE`:

```json
[
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": { ... },  // Полностью заполнено
    "Bolts": { ... },     // Полностью заполнено
    "InternalForces": { "Nt": 100 }
  },
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": {},       // Пусто
    "Bolts": {},          // Пусто
    "InternalForces": { "Nt": 200 }  // Другие силы
  }
]
```

ConvertData автоматически скопирует `Geometry` и `Bolts` из первой записи во вторую!

Результат:

```json
[
  { ... },  // Первая без изменений
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": { ... },  // ← СКОПИРОВАНО!
    "Bolts": { ... },     // ← СКОПИРОВАНО!
    "InternalForces": { "Nt": 200 }  // Осталось своё
  }
]
```

---

## 🎓 Что дальше?

### Изучите документацию

1. **[README.md](README.md)** — полное описание проекта
2. **[Architecture.md](Architecture.md)** — архитектура системы
3. **[DataFlow.md](DataFlow.md)** — детали обработки данных
4. **[API.md](API.md)** — справочник по API
5. **[UML.md](UML.md) / [Mermaid.md](Mermaid.md)** — диаграммы

### Попробуйте расширенные возможности

```bash
# Обработать конкретные файлы
ConvertData.exe file1.xlsx file2.xls

# Переопределить колонку профиля
ConvertData.exe --profile-column=ProfileColumn

# Только создать JSON
ConvertData.exe 1 file1.xlsx file2.xlsx

# Только применить профили
ConvertData.exe 2
```

---

## 📞 Нужна помощь?

- **GitHub Issues:** [Сообщить о проблеме](https://github.com/NikitaShchegolev/ConvertData/issues)
- **Документация:** См. [INDEX.md](INDEX.md)
- **Email:** Свяжитесь через GitHub

---

## ✅ Чек-лист быстрого старта

- [ ] Скачал и распаковал ConvertData
- [ ] Поместил Excel файлы в `EXCEL/`
- [ ] (Опционально) Добавил `ProfileBeam.xls` в `EXCEL_Profile/`
- [ ] Запустил `ConvertData.exe`
- [ ] Проверил результаты в `JSON_All/all_NotDuplicate.json`
- [ ] Изучил отчёты о дубликатах
- [ ] (Опционально) Экспортировал справочники из `EXCEL_Profile_OUT/`

**Готово!** 🎉

---

**Время прочтения:** ~5 минут  
**Время настройки:** ~2 минуты  
**Время первой конвертации:** ~1 минута (зависит от объёма данных)

---

[← Вернуться к INDEX](INDEX.md) | [Полная документация →](README.md)