# Бар BI Lite WebApp

Мобильный Telegram WebApp без backend. Работает как локальный обработчик Excel-файлов прямо в браузере.

## Что умеет

- Загружать ОСВ `.xlsx/.xls`
- Загружать прайсы поставщиков `.xlsx/.xls`
- Строить дашборд
- Считать стопы, риск стопа, оверстоки
- Делать заявку по поставщику
- Автоматически ставить поставщика `Овощник` для фруктов, овощей, ягод и зелени, если поставщик не найден в прайсах
- Делать ABC-анализ
- Экспортировать заявку в CSV

## Важно

Файлы не отправляются на сервер. Всё считается в браузере пользователя.

## Как запустить локально

Просто открой `index.html` в браузере.

## Как залить на GitHub Pages

1. Создай репозиторий, например `bar-bi-lite`
2. Загрузи туда файлы:
   - `index.html`
   - `style.css`
   - `app.js`
3. В GitHub открой Settings → Pages
4. Source: Deploy from branch
5. Branch: `main`, folder `/root`
6. Получишь ссылку вида:

```text
https://username.github.io/bar-bi-lite/
```

## Как открыть через Telegram WebApp

В BotFather:

```text
/mybots → твой бот → Bot Settings → Menu Button → Configure menu button
```

Вставь URL GitHub Pages.

Также можно сделать кнопку в боте с WebApp URL.
