# REST API Service

Учебный проект REST API сервиса на Python (FastAPI, SQLAlchemy, Alembic, PostgreSQL).
Разработано для Занятия 02 (Вариант 5: REST API-сервис) колледжа Хекслет Казахстан.

## Структура проекта
- `app/` — ядро приложения, роутеры API, схемы валидации Pydantic, ORM-модели.
- `migrations/` — версионирование схемы БД через Alembic.
- `tests/` — модульные и интеграционные тесты (pytest).
- `docs/` — спецификация OpenAPI и архитектурное описание.
- `data/` — фикстуры и начальные данные для наполнения БД.

## Запуск в едином окружении (Docker)
```bash
cp .env.example .env
docker compose up --build -d
```
