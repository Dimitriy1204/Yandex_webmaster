#!/usr/bin/env python3
"""
Yandex Wordstat API Agent — сбор и анализ поискового спроса.

Скрипт собирает помесячную частотность поисковых запросов через
Yandex Search API v2 (Wordstat Dynamics) и формирует Excel-отчёт.

API endpoint: POST https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics
Документация: https://aistudio.yandex.ru/docs/ru/search-api/api-ref/Wordstat/getDynamics

Использование:
    1. Получите API-ключ Яндекс.Облака:
       https://console.cloud.yandex.ru/ → Сервисные аккаунты → API-ключи
    2. Скопируйте .env.example → .env и заполните YANDEX_API_KEY и FOLDER_ID
    3. Установите зависимости: pip install -r requirements.txt
    4. Запустите: python analyzer.py

Выходной файл: yandex_analysis.xlsx (6 листов: О файле, Данные, Сводка, График,
    Голубые океаны (3 блока: momentum, mean, slope),
    Растущие рынки (3 блока: momentum, mean, slope))
"""

from __future__ import annotations

import os
import sys
import time
import logging
import traceback
from typing import Any
from datetime import datetime

import numpy as np
import requests
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

matplotlib.use("Agg")
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False

from dotenv import load_dotenv

# ─────────────────────────────────────────────────────────────────
# Конфигурация — загружается из .env рядом со скриптом
# ─────────────────────────────────────────────────────────────────
_SCRIPT_DIR: str = os.path.dirname(os.path.abspath(__file__))
_ENV_PATH: str = os.path.join(_SCRIPT_DIR, ".env")
load_dotenv(_ENV_PATH)

YANDEX_API_KEY: str = os.getenv("YANDEX_API_KEY", "")
FOLDER_ID: str = os.getenv("FOLDER_ID", "")
_ENV_QUERIES: str = os.getenv("SEARCH_QUERIES", "")
_ENV_DATE_FROM: str = os.getenv("DATE_FROM", "")
_ENV_DATE_TO: str = os.getenv("DATE_TO", "")

# ─────────────────────────────────────────────────────────────────
# Настройки
# ─────────────────────────────────────────────────────────────────
QUERIES: list[str] = [q.strip() for q in _ENV_QUERIES.split(",") if q.strip()]

FALLBACK_QUERIES: list[str] = [
    "купить оборудование для",
    "купить станок для",
    "купить линию для",
    "купить производство",
    "купить комплекс",
    "купить вибропресс",
    "купить бетонный завод",
    "купить ДСК",
    "купить дробилку",
    "купить горно добывающее оборудование",
]

if not QUERIES:
    QUERIES = FALLBACK_QUERIES

API_URL: str = "https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics"
REQUEST_DELAY: float = 0.4
OUTPUT_FILE: str = "yandex_analysis.xlsx"

# ─────────────────────────────────────────────────────────────────
# Логирование (UTF-8 для Windows)
# ─────────────────────────────────────────────────────────────────
class ForceUtf8Handler(logging.StreamHandler):
    def __init__(self, stream=None):
        super().__init__(stream)
        try:
            import io
            self.stream = io.TextIOWrapper(
                self.stream.buffer, encoding="utf-8", errors="replace"
            )
        except Exception:
            pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
    handlers=[ForceUtf8Handler()],
)
logger = logging.getLogger("wordstat")


# ─────────────────────────────────────────────────────────────────
# Валидация .env
# ─────────────────────────────────────────────────────────────────
def validate_env() -> None:
    if not os.path.exists(_ENV_PATH):
        logger.error(f".env not found at: {_ENV_PATH}")
        sys.exit(1)
    if not YANDEX_API_KEY or YANDEX_API_KEY == "ваш_api_key_яндекс_облака":
        logger.error("YANDEX_API_KEY is missing or unchanged in .env")
        sys.exit(1)
    if not FOLDER_ID or FOLDER_ID == "ваш_folder_id":
        logger.error("FOLDER_ID is missing in .env")
        sys.exit(1)
    logger.info(".env OK")


def _last_day_of_month(year: int, month: int) -> int:
    """Возвращает последний день месяца (1-31)."""
    if month == 12:
        next_month = datetime(year + 1, 1, 1)
    else:
        next_month = datetime(year, month + 1, 1)
    return (next_month - datetime(year, month, 1)).days


def _parse_date_range() -> tuple[str, str]:
    """
    Парсит DATE_FROM / DATE_TO из .env.
    Ожидаемый формат: YYYY-MM (например, 2022-01).
    Если не указаны — используются дефолтные значения.

    Returns:
        (from_date_iso, to_date_iso) — ISO-строки для API.
        from_date всегда = 1-е число месяца.
        to_date всегда = последний день месяца (требование API).
    """
    today = datetime.now()

    # DATE_FROM: если не указан → 2022-01
    if _ENV_DATE_FROM.strip():
        try:
            parts = _ENV_DATE_FROM.strip().split("-")
            year = int(parts[0])
            month = int(parts[1])
            from_date = f"{year:04d}-{month:02d}-01T00:00:00Z"
            logger.info(f"DATE_FROM from .env: {_ENV_DATE_FROM} → {from_date}")
        except (ValueError, IndexError):
            logger.warning(f"Invalid DATE_FROM='{_ENV_DATE_FROM}', using default 2022-01")
            from_date = "2022-01-01T00:00:00Z"
    else:
        from_date = "2022-01-01T00:00:00Z"

    # DATE_TO: если не указан → последний ПОЛНЫЙ месяц (текущего года, а не прошлого)
    if _ENV_DATE_TO.strip():
        try:
            parts = _ENV_DATE_TO.strip().split("-")
            year = int(parts[0])
            month = int(parts[1])
            last_day = _last_day_of_month(year, month)
            to_date = f"{year:04d}-{month:02d}-{last_day:02d}T00:00:00Z"
            logger.info(f"DATE_TO from .env: {_ENV_DATE_TO} → {to_date}")
        except (ValueError, IndexError):
            logger.warning(f"Invalid DATE_TO='{_ENV_DATE_TO}', using default last full month")
            # Последний полный месяц
            if today.month == 1:
                y, m = today.year - 1, 12
            else:
                y, m = today.year, today.month - 1
            last_day = _last_day_of_month(y, m)
            to_date = f"{y:04d}-{m:02d}-{last_day:02d}T00:00:00Z"
    else:
        # Если DATE_TO не указан — берём последний полный месяц
        if today.month == 1:
            y, m = today.year - 1, 12
        else:
            y, m = today.year, today.month - 1
        last_day = _last_day_of_month(y, m)
        to_date = f"{y:04d}-{m:02d}-{last_day:02d}T00:00:00Z"
        logger.info(f"DATE_TO not set, using last full month: {y:04d}-{m:02d}")

    return from_date, to_date


# ─────────────────────────────────────────────────────────────────
# Запрос к Wordstat API v2
# ─────────────────────────────────────────────────────────────────
def fetch_wordstat_dynamics(query: str, date_from: str, date_to: str) -> list[dict[str, Any]]:
    """
    Запрашивает динамику частотности для одного запроса.
    Использует period=PERIOD_MONTHLY как требует API v2.

    Args:
        query: Поисковый запрос
        date_from: Начальная дата в ISO (напр. 2022-01-01T00:00:00Z)
        date_to: Конечная дата в ISO (должна быть последним днём месяца)

    Returns:
        Список словарей с данными динамики
    """
    headers: dict[str, str] = {
        "Authorization": f"Api-Key {YANDEX_API_KEY}",
        "Content-Type": "application/json",
    }

    payload: dict[str, Any] = {
        "folderId": FOLDER_ID,
        "phrase": query,
        "period": "PERIOD_MONTHLY",
        "from_date": date_from,
        "to_date": date_to,
        "geo_ids": [225],  # Россия
        "group_by": "TIME",
    }

    max_retries: int = 3
    last_error: Exception | None = None

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(
                API_URL, headers=headers, json=payload, timeout=60
            )

            if response.status_code in (401, 403):
                logger.error(
                    f"Auth error {response.status_code}: invalid API key.\n"
                    f"Response: {response.text[:300]}"
                )
                sys.exit(1)

            if response.status_code == 429:
                wait = 2 ** attempt
                logger.warning(f"Rate limit (429). Pausing {wait}s (attempt {attempt}/{max_retries})")
                time.sleep(wait)
                continue

            if response.status_code != 200:
                err_body = response.text[:400] if response.text else "(empty)"
                logger.warning(f"HTTP {response.status_code} for '{query}': {err_body}")
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                    continue
                return []

            data: dict[str, Any] = response.json()

            # API v2 возвращает { "results": [ { "date": "...", "count": "12345", "share": 0.xxx }, ... ] }
            results: list[dict[str, Any]] = data.get("results", [])

            if not results:
                logger.warning(f"Empty data for '{query}'")
                return []

            first_date = results[0].get("date", "?")
            last_date = results[-1].get("date", "?")
            logger.info(f"  -> {len(results)} datapoints ({first_date}..{last_date})")
            return results

        except (requests.ConnectionError, requests.Timeout) as e:
            last_error = e
            logger.warning(f"Network error (attempt {attempt}/{max_retries}): {e}")
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
        except requests.RequestException as e:
            last_error = e
            logger.error(f"Request error for '{query}': {e}")
            return []

    logger.error(f"All {max_retries} attempts failed for '{query}'. Last error: {last_error}")
    return []


# ─────────────────────────────────────────────────────────────────
# Парсинг ответа API v2
# ─────────────────────────────────────────────────────────────────
def parse_dynamics(query: str, results: list[dict[str, Any]]) -> pd.DataFrame:
    """
    Преобразует массив results из API v2 в DataFrame.

    Структура элемента results:
      {"date": "2025-01-01T00:00:00Z", "count": "73897", "share": 0.0006227}
    """
    rows: list[dict[str, Any]] = []

    for entry in results:
        date_str: str | None = entry.get("date")
        if not date_str:
            continue

        # count — строка с числом, может быть "0" или пустой
        count_raw = entry.get("count")
        if count_raw is None:
            continue

        try:
            frequency = int(str(count_raw))
        except (ValueError, TypeError):
            logger.warning(f"  -> Non-numeric count '{count_raw}' for '{query}' at {date_str}")
            continue

        rows.append({
            "date": date_str,
            "query": query,
            "frequency": frequency,
        })

    if not rows:
        logger.warning(f"  -> No parsable data for '{query}'")

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────
# Сбор данных по всем запросам
# ─────────────────────────────────────────────────────────────────
def collect_all_data(date_from: str, date_to: str) -> pd.DataFrame:
    logger.info("=" * 60)
    logger.info(f"Collecting data for {len(QUERIES)} queries")
    logger.info(f"Date range: {date_from} .. {date_to}")
    logger.info("=" * 60)

    all_dfs: list[pd.DataFrame] = []

    for i, query in enumerate(QUERIES, 1):
        logger.info(f"[{i}/{len(QUERIES)}] '{query}'")

        results = fetch_wordstat_dynamics(query, date_from, date_to)
        df = parse_dynamics(query, results)

        if not df.empty:
            all_dfs.append(df)
            logger.info(f"  -> {len(df)} records")
        else:
            logger.warning(f"  -> No data for '{query}'")

        if i < len(QUERIES):
            time.sleep(REQUEST_DELAY)

    if not all_dfs:
        logger.error("No data collected for any query.")
        sys.exit(1)

    result: pd.DataFrame = pd.concat(all_dfs, ignore_index=True)
    result["date"] = pd.to_datetime(result["date"])
    result = result.sort_values(["query", "date"]).reset_index(drop=True)

    logger.info(f"Total: {len(result)} records")
    return result


# ─────────────────────────────────────────────────────────────────
# Расчёт метрик (включая growth_momentum и trend_slope)
# ─────────────────────────────────────────────────────────────────
def _calc_trend_slope(series: pd.Series) -> float:
    """Наклон линейной регрессии (полином 1-й степени) по frequency."""
    x = np.arange(len(series))
    y = series.values.astype(float)
    if len(x) < 2:
        return 0.0
    coeffs = np.polyfit(x, y, 1)
    return float(coeffs[0])  # slope = прирост показов в месяц


def calculate_metrics(df: pd.DataFrame) -> pd.DataFrame:
    logger.info("Calculating metrics...")

    df = df.copy().sort_values(["query", "date"]).reset_index(drop=True)

    # freq_ma3
    df["freq_ma3"] = df.groupby("query")["frequency"].transform(
        lambda x: x.rolling(window=3, min_periods=1).mean()
    )

    # growth_pct
    df["growth_pct"] = df.groupby("query")["frequency"].transform(
        lambda x: x.pct_change(periods=1) * 100
    ).fillna(0.0)

    # growth_momentum — средний рост за последние 6 месяцев
    momentum_map: dict[str, float] = {}
    for q, grp in df.groupby("query"):
        grp = grp.sort_values("date")
        # БЕРЁМ последние 7 строк (6 переходов), growth_pct первой = 0
        tail = grp.tail(7)
        vals = tail["growth_pct"].iloc[1:] if len(tail) > 1 else tail["growth_pct"]
        momentum_map[q] = vals.mean() if len(vals) > 0 else 0.0
    df["growth_momentum"] = df["query"].map(momentum_map)

    # trend_slope — наклон линейной регрессии по frequency
    slope_map: dict[str, float] = df.groupby("query")["frequency"].apply(_calc_trend_slope).to_dict()
    df["trend_slope"] = df["query"].map(slope_map)

    # volatility
    vol_map: dict[str, float] = df.groupby("query")["frequency"].std().to_dict()
    df["volatility"] = df["query"].map(vol_map)

    # seasonality_idx
    mean_map: dict[str, float] = df.groupby("query")["frequency"].mean().to_dict()
    df["seasonality_idx"] = df.apply(
        lambda r: r["frequency"] / mean_map[r["query"]] if mean_map.get(r["query"], 0) > 0 else 1.0,
        axis=1,
    )

    # competition_score
    agg: pd.DataFrame = df.groupby("query").agg(
        mean_freq=("frequency", "mean"),
        std_freq=("frequency", "std"),
    ).reset_index()

    max_freq = agg["mean_freq"].max()
    min_freq = agg["mean_freq"].min()
    max_vol = agg["std_freq"].max()
    min_vol = agg["std_freq"].min()

    def norm(v: float, mn: float, mx: float) -> float:
        return 0.5 if mx == mn else (v - mn) / (mx - mn)

    agg["freq_norm"] = agg["mean_freq"].apply(lambda x: norm(x, min_freq, max_freq))
    agg["vol_norm"] = agg["std_freq"].apply(lambda x: 1 - norm(x, min_vol, max_vol))
    agg["competition_score"] = 0.5 * agg["freq_norm"] + 0.5 * agg["vol_norm"]

    score_map = dict(zip(agg["query"], agg["competition_score"]))
    df["competition_score"] = df["query"].map(score_map)

    logger.info("Metrics calculated")
    return df


# ─────────────────────────────────────────────────────────────────
# Вспомогательные функции для Excel
# ─────────────────────────────────────────────────────────────────

def _fmt_thousands(val: float) -> str:
    """Форматирует число с пробелом как разделителем тысяч."""
    s = f"{val:_.0f}"
    return s.replace("_", " ")


def _fmt_thousands_dec(val: float, dec: int = 1) -> str:
    """Форматирует число с пробелом как разделителем тысяч и dec знаков."""
    s = f"{val:_.{dec}f}"
    return s.replace("_", " ")


# ── Универсальные функции вердиктов ────────────────────────────

def blue_verdict(comp: float, growth: float) -> str:
    if comp < 0.4 and growth > 5:
        return "🟢 Голубой океан - низкая конкуренция + высокий рост"
    elif comp < 0.4 and growth > 0:
        return "🟢 Низкая конкуренция, есть рост"
    elif comp < 0.5 and growth > 3:
        return "🟡 Умеренная конкуренция, хороший рост"
    elif comp < 0.5 and growth > 0:
        return "🟡 Стабильная ниша с ростом"
    elif growth > 0:
        return "🟠 Рост есть, но высокая конкуренция"
    else:
        return "🔴 Спрос падает - осторожно"


def growth_verdict(growth: float) -> str:
    if growth > 5:
        return "🟢 Быстрый рост"
    elif growth > 2:
        return "🟢 Уверенный рост"
    elif growth > 0.5:
        return "🟡 Слабый, но положительный рост"
    elif growth > 0:
        return "🟡 Незначительный рост"
    else:
        return "🔴 Падение спроса"


# ── Генерация графика ──────────────────────────────────────────

def _plot_dynamics(raw_data: pd.DataFrame, queries: list[str],
                   title: str, filename: str, date_range_str: str) -> None:
    """Сохраняет линейный график для списка запросов."""
    plot_df = raw_data[raw_data["query"].isin(queries)]

    plt.figure(figsize=(12, 6))
    for q in queries:
        qd = plot_df[plot_df["query"] == q].sort_values("date")
        if not qd.empty:
            plt.plot(qd["date"], qd["frequency"], marker="o", label=q[:40], linewidth=2)

    plt.title(f"{title}\n{date_range_str}", fontsize=16, fontweight="bold")
    plt.xlabel("Дата", fontsize=14)
    plt.ylabel("Показов в месяц", fontsize=14)
    plt.legend(bbox_to_anchor=(1.05, 1), loc="upper left", fontsize=10)
    plt.xticks(rotation=45, fontsize=12)
    plt.yticks(fontsize=12)
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(filename, dpi=150, bbox_inches="tight")
    plt.close()


# ── Генерация блока таблицы на листе Excel ─────────────────────
# Каждая функция пишет блок начиная с заданной строки (start_row)

def _write_ocean_block(ws, df: pd.DataFrame, start_row: int,
                       block_title: str, date_range_str: str,
                       hdr_fmt, title_fmt, txt_fmt,
                       int_fmt, num_fmt, comp_fmt,
                       v_fmt_good, v_fmt_bad, v_fmt_neutral,
                       growth_col: str) -> int:
    """Пишет блок 'Голубые океаны' начиная с start_row. Возвращает следующую строку."""
    r = start_row
    ws.merge_range(r, 0, r, 5,
        f"{block_title}\n{date_range_str}", title_fmt)
    r += 1
    cols = [
        "N",
        "Поисковый запрос\nquery",
        "Индекс конкуренции\nCompetition Index\n🟢<0.40 🟡0.40-0.60 🔴>0.60",
        "Метрика роста\nGrowth (%)\n(значение метрики)",
        "Суммарная частотность\nSum Frequency\nЕмкость рынка",
        "Вердикт",
    ]
    ws.write_row(r, 0, cols, hdr_fmt)
    r += 1
    for ri, (_, row_data) in enumerate(df.iterrows()):
        ws.write(r, 0, ri + 1, int_fmt)
        ws.write(r, 1, row_data["query"], txt_fmt)
        score = round(row_data["competition_score"], 3)
        ws.write(r, 2, score, comp_fmt(score))
        growth_val = row_data[growth_col]
        ws.write(r, 3, round(growth_val, 2), num_fmt)
        ws.write(r, 4, int(row_data["sum_frequency"]), int_fmt)
        bv = blue_verdict(row_data["competition_score"], growth_val)
        is_good = "🟢" in bv
        is_bad = "🔴" in bv or "🟠" in bv
        vf = v_fmt_good if is_good else (v_fmt_bad if is_bad else v_fmt_neutral)
        ws.write(r, 5, bv, vf)
        r += 1
    return r + 2  # отступ между блоками


def _write_growing_block(ws, df: pd.DataFrame, start_row: int,
                         block_title: str, date_range_str: str,
                         hdr_fmt, title_fmt, txt_fmt,
                         int_fmt, num_fmt, comp_fmt, vol_fmt_h,
                         v_fmt_good, v_fmt_bad, v_fmt_neutral,
                         growth_col: str) -> int:
    """Пишет блок 'Растущие рынки' начиная с start_row. Возвращает следующую строку."""
    r = start_row
    ws.merge_range(r, 0, r, 6,
        f"{block_title}\n{date_range_str}", title_fmt)
    r += 1
    cols = [
        "N",
        "Поисковый запрос\nquery",
        "Метрика роста\nGrowth (%)\n(значение метрики)",
        "Суммарная частотность\nSum Frequency\nЕмкость рынка",
        "Индекс конкуренции\nCompetition Index\n🟢<0.40 🟡0.40-0.60 🔴>0.60",
        "Волатильность\nVolatility\n🟢низкая 🟡средняя 🔴высокая",
        "Вердикт",
    ]
    ws.write_row(r, 0, cols, hdr_fmt)
    r += 1
    for ri, (_, row_data) in enumerate(df.iterrows()):
        ws.write(r, 0, ri + 1, int_fmt)
        ws.write(r, 1, row_data["query"], txt_fmt)
        growth_val = row_data[growth_col]
        ws.write(r, 2, round(growth_val, 2), num_fmt)
        ws.write(r, 3, int(row_data["sum_frequency"]), int_fmt)
        score = round(row_data["competition_score"], 3)
        ws.write(r, 4, score, comp_fmt(score))
        vol = round(row_data["mean_volatility"], 1)
        label = row_data["vol_label"]
        ws.write(r, 5, vol, vol_fmt_h(vol, label))
        gv = growth_verdict(growth_val)
        is_growing = "🟢" in gv
        is_bad = "🔴" in gv
        vf = v_fmt_good if is_growing else (v_fmt_bad if is_bad else v_fmt_neutral)
        ws.write(r, 6, gv, vf)
        r += 1
    return r + 2  # отступ между блоками


# ─────────────────────────────────────────────────────────────────
# Создание форматированного Excel-отчёта (6 листов)
# ─────────────────────────────────────────────────────────────────
def create_excel_report(raw_data: pd.DataFrame, date_from: str, date_to: str,
                        output_path: str = OUTPUT_FILE) -> str:
    logger.info(f"Creating Excel report: {output_path}")

    import xlsxwriter

    # ── Подготовка данных ──────────────────────────────────────
    cols: list[str] = [
        "date", "query", "frequency", "freq_ma3", "growth_pct",
        "volatility", "seasonality_idx", "competition_score",
    ]
    sheet1_df = raw_data[cols].copy()
    sheet1_df["date"] = sheet1_df["date"].dt.strftime("%Y-%m-%d")

    # Сводка по запросам (все 3 метрики роста)
    summary = raw_data.groupby("query").agg(
        mean_frequency=("frequency", "mean"),
        sum_frequency=("frequency", "sum"),
        std_frequency=("frequency", "std"),
        mean_growth=("growth_pct", "mean"),
        growth_momentum=("growth_momentum", "first"),
        trend_slope=("trend_slope", "first"),
        mean_volatility=("volatility", "mean"),
        competition_score=("competition_score", "first"),
    ).reset_index().sort_values("sum_frequency", ascending=False).reset_index(drop=True)

    # Для метрик волатильности — категория на основе перцентилей
    vol_values = summary["mean_volatility"]
    vol_lo = vol_values.quantile(0.33)
    vol_hi = vol_values.quantile(0.66)

    def vol_label(v: float) -> str:
        if v <= vol_lo:
            return "Низкая"
        elif v <= vol_hi:
            return "Средняя"
        else:
            return "Высокая"

    summary["vol_label"] = summary["mean_volatility"].apply(vol_label)

    # Границы для competition_score
    comp_values = summary["competition_score"]
    comp_lo = comp_values.quantile(0.33)
    comp_hi = comp_values.quantile(0.66)

    # ── Форматированная строка периода для заголовков ──────────
    try:
        df_dt = datetime.strptime(date_from[:10], "%Y-%m-%d")
        dt_dt = datetime.strptime(date_to[:10], "%Y-%m-%d")
        date_range_str = f"{df_dt.year}-{df_dt.month:02d} .. {dt_dt.year}-{dt_dt.month:02d}"
    except Exception:
        date_range_str = f"{date_from[:7]} .. {date_to[:7]}"

    # ── Таблицы "Голубые океаны" (3 варианта) ────────────────────
    def make_blue(growth_col: str) -> pd.DataFrame:
        df = summary.copy()
        g = df[growth_col]
        df["blue_score"] = (
            (1 - df["competition_score"]) * 50 +
            (g - g.min()) / (g.max() - g.min() + 1e-9) * 50
        )
        return df.sort_values("blue_score", ascending=False).reset_index(drop=True)

    blue_momentum = make_blue("growth_momentum")
    blue_mean = make_blue("mean_growth")
    blue_slope = make_blue("trend_slope")

    # ── Таблицы "Растущие рынки" (3 варианта) ────────────────────
    growing_momentum = summary.sort_values("growth_momentum", ascending=False).reset_index(drop=True)
    growing_mean = summary.sort_values("mean_growth", ascending=False).reset_index(drop=True)
    growing_slope = summary.sort_values("trend_slope", ascending=False).reset_index(drop=True)

    # ── Графики ─────────────────────────────────────────────────
    img_demand: str = "_temp_plot_demand.png"
    img_ocean_mom: str = "_temp_plot_ocean_mom.png"
    img_growing_mom: str = "_temp_plot_growing_mom.png"

    # График 1: Топ-10 по спросу
    top10_demand = summary["query"].head(10).tolist()
    _plot_dynamics(raw_data, top10_demand,
                   "Динамика частотности - Топ-10 запросов",
                   img_demand, date_range_str)

    # График 2: Топ-10 Голубые океаны (momentum)
    ocean_top10 = blue_momentum["query"].head(10).tolist()
    _plot_dynamics(raw_data, ocean_top10,
                   "Топ-10: Низкая конкуренция + рост (momentum)",
                   img_ocean_mom, date_range_str)

    # График 3: Топ-10 Растущие рынки (momentum)
    growing_top10 = growing_momentum["query"].head(10).tolist()
    _plot_dynamics(raw_data, growing_top10,
                   "Топ-10: Растущие рынки (momentum)",
                   img_growing_mom, date_range_str)

    # ── Создание книги Excel ──────────────────────────────────
    wb = xlsxwriter.Workbook(output_path)

    # Форматы данных
    hdr_fmt = wb.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "white",
                              "border": 2, "align": "center", "text_wrap": True, "font_size": 14,
                              "valign": "vcenter"})
    title_fmt = wb.add_format({"bold": True, "font_size": 16, "font_color": "#2F5496",
                                "bottom": 2, "bottom_color": "#4472C4", "text_wrap": True})
    conclusion_fmt = wb.add_format({"bold": True, "font_size": 16, "font_color": "#2F5496",
                                     "bottom": 2, "bottom_color": "#4472C4", "text_wrap": True})
    guide_fmt = wb.add_format({"text_wrap": True, "font_size": 14, "valign": "top"})
    guide_bold = wb.add_format({"text_wrap": True, "font_size": 14, "bold": True, "valign": "top"})
    num_fmt = wb.add_format({"num_format": "# ##0.00", "border": 1, "font_size": 14})
    int_fmt = wb.add_format({"num_format": "# ##0", "border": 1, "font_size": 14})
    pct_fmt = wb.add_format({"num_format": "+0.00%;-0.00%", "border": 1, "font_size": 14})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1, "font_size": 14})
    txt_fmt = wb.add_format({"border": 1, "text_wrap": True, "font_size": 14, "valign": "vcenter"})
    green_fmt = wb.add_format({"num_format": "0.00", "border": 1, "bg_color": "#A9D18E", "font_size": 14})
    yellow_fmt = wb.add_format({"num_format": "0.00", "border": 1, "bg_color": "#FFD966", "font_size": 14})
    red_fmt = wb.add_format({"num_format": "0.00", "border": 1, "bg_color": "#F4B183", "font_size": 14})
    green_vol_fmt = wb.add_format({"num_format": "# ##0.0", "border": 1, "bg_color": "#A9D18E", "font_size": 14})
    yellow_vol_fmt = wb.add_format({"num_format": "# ##0.0", "border": 1, "bg_color": "#FFD966", "font_size": 14})
    red_vol_fmt = wb.add_format({"num_format": "# ##0.0", "border": 1, "bg_color": "#F4B183", "font_size": 14})
    v_fmt_good = wb.add_format({"border": 1, "text_wrap": True, "bold": True,
                                 "font_size": 14, "font_color": "#006400"})
    v_fmt_bad = wb.add_format({"border": 1, "text_wrap": True, "bold": True,
                                "font_size": 14, "font_color": "#8B0000"})
    v_fmt_neutral = wb.add_format({"border": 1, "text_wrap": True, "bold": True,
                                    "font_size": 14, "font_color": "#404040"})

    def comp_fmt(score: float):
        if score < 0.4:
            return green_fmt
        elif score > 0.6:
            return red_fmt
        else:
            return yellow_fmt

    def vol_fmt_h(v: float, label: str):
        if label == "Низкая":
            return green_vol_fmt
        elif label == "Высокая":
            return red_vol_fmt
        else:
            return yellow_vol_fmt

    # ==========================================================
    # ЛИСТ 1: О файле
    # ==========================================================
    ws_info = wb.add_worksheet("О файле")
    ws_info.set_column(0, 0, 8)
    ws_info.set_column(1, 3, 65)

    row = 0
    ws_info.merge_range(row, 0, row, 3,
        f"Анализ поискового спроса - Yandex Wordstat API\n{date_range_str}", title_fmt)
    row += 2

    # Выводы по данным
    ws_info.merge_range(row, 0, row, 3,
        "ПРАКТИЧЕСКИЙ ВЫВОД ПО ВАШИМ ДАННЫМ", conclusion_fmt)
    row += 2

    s = summary
    top_demand = s.nlargest(3, "sum_frequency")
    top_growth_mom = s.nlargest(3, "growth_momentum")
    low_comp_growth_mom = s[(s["competition_score"] < 0.4) & (s["growth_momentum"] > 0)] \
        .sort_values("growth_momentum", ascending=False)
    low_comp = s[s["competition_score"] < 0.4].sort_values("competition_score")

    conclusions = []
    conclusions.append(("Самая востребованная ниша (емкость рынка):",
        f"{top_demand.iloc[0]['query']} - {_fmt_thousands(top_demand.iloc[0]['sum_frequency'])} показов\n"
        f"Также в топе: {top_demand.iloc[1]['query']} ({_fmt_thousands(top_demand.iloc[1]['sum_frequency'])}) и "
        f"{top_demand.iloc[2]['query']} ({_fmt_thousands(top_demand.iloc[2]['sum_frequency'])})"))
    conclusions.append(("Самая перспективная растущая ниша (momentum):",
        f"{top_growth_mom.iloc[0]['query']} - рост {top_growth_mom.iloc[0]['growth_momentum']:+.2f}% в месяц\n"
        f"Также растут: {top_growth_mom.iloc[1]['query']} ({top_growth_mom.iloc[1]['growth_momentum']:+.2f}%) и "
        f"{top_growth_mom.iloc[2]['query']} ({top_growth_mom.iloc[2]['growth_momentum']:+.2f}%)"))
    if not low_comp_growth_mom.empty:
        best = low_comp_growth_mom.iloc[0]
        conclusions.append(("Лучшее для старта (низкая конкуренция + рост):",
            f"{best['query']} - конкуренция {best['competition_score']:.2f}, "
            f"рост (momentum) {best['growth_momentum']:+.2f}%/мес"))
    else:
        conclusions.append(("Лучшее для старта (низкая конкуренция + рост):",
            "Не найдено запросов, одновременно удовлетворяющих условиям."))
    if not low_comp.empty:
        lowest = low_comp.iloc[0]
        conclusions.append(("Самая низкая конкуренция (легкий вход):",
            f"{lowest['query']} - индекс конкуренции {lowest['competition_score']:.2f}"))

    for label, text in conclusions:
        ws_info.write(row, 0, ">", guide_bold)
        ws_info.merge_range(row, 1, row, 3, label, guide_bold)
        row += 1
        ws_info.write(row, 1, text, guide_fmt)
        row += 2

    # Описание листов
    ws_info.merge_range(row, 0, row, 3,
        "ОПИСАНИЕ ЛИСТОВ ФАЙЛА", conclusion_fmt)
    row += 2

    sheets_desc = [
        ("Лист Данные", "Сырые данные по каждому запросу помесячно."),
        ("Лист Сводка по запросам",
         "Итоговые метрики по каждому запросу, включая все 3 метрики роста: "
         "mean_growth (средний за весь период), growth_momentum (средний за последние 6 мес.), "
         "trend_slope (наклон тренда, показов/мес)."),
        ("Лист График", "Визуальная динамика топ-10 запросов."),
        ("Листы Голубые океаны (3 шт)",
         "3 варианта оценки низкоконкурентных ниш с ростом: "
         "по momentum (свежий тренд), по mean (весь период), по slope (линейный тренд)."),
        ("Листы Растущие рынки (3 шт)",
         "3 варианта оценки растущих рынков: "
         "по momentum, по mean, по slope."),
    ]
    for title, desc in sheets_desc:
        ws_info.write(row, 1, title, guide_bold)
        ws_info.merge_range(row + 1, 1, row + 1, 3, desc, guide_fmt)
        row += 3

    # Методология
    ws_info.merge_range(row, 0, row, 3,
        "МЕТОДОЛОГИЯ", conclusion_fmt)
    row += 2

    vol_lo_str = _fmt_thousands_dec(vol_lo)
    vol_hi_str = _fmt_thousands_dec(vol_hi)
    methodology = [
        ("Источник данных",
         f"Yandex Search API v2 (Wordstat GetDynamics).\n"
         f"Период: {date_range_str} ({date_from[:10]} .. {date_to[:10]}).\n"
         f"Регион: вся Россия."),
        ("growth_momentum (Моментум роста)",
         "Средний помесячный темп роста за последние 6 месяцев.\n"
         "Расчет: среднее growth_pct по последним 7 записям (6 переходов).\n"
         "Показывает свежий тренд, не подвержен давним скачкам."),
        ("trend_slope (Наклон тренда)",
         "Абсолютный прирост частотности в месяц (линейная регрессия).\n"
         "Расчет: numpy polyfit(x, frequency, 1), где x = 0,1,2...\n"
         "Показывает, на сколько показов в месяц в среднем меняется спрос."),
        ("mean_growth (Средний рост)",
         "Средний помесячный темп роста за весь период наблюдений.\n"
         "Подвержен влиянию давних скачков — менее точен для текущего тренда."),
        ("volatility (Волатильность)",
         "Стандартное отклонение частотности за весь период.\n"
         f"Границы данного отчета:\n"
         f"  🟢 Низкая (≤ {vol_lo_str})\n"
         f"  🟡 Средняя ({vol_lo_str}–{vol_hi_str})\n"
         f"  🔴 Высокая (≥ {vol_hi_str})"),
        ("competition_score (Индекс конкуренции, 0..1)",
         "0.5 × норм.частотность + 0.5 × (1 − норм.волатильность).\n"
         f"Границы: 🟢 < 0.40, 🟡 0.40–0.60, 🔴 > 0.60"),
        ("Голубые океаны",
         "Составной индекс: 50% × (1 − competition_score) + 50% × норм.growth.\n"
         "Чем выше — тем перспективнее ниша для входа."),
    ]
    for title, desc in methodology:
        ws_info.write(row, 1, title, guide_bold)
        ws_info.merge_range(row + 1, 1, row + 1, 3, desc, guide_fmt)
        row += 4

    # ==========================================================
    # ЛИСТ 2: Данные
    # ==========================================================
    ws1 = wb.add_worksheet("Данные")
    ws1.merge_range(0, 0, 0, 7,
        f"Сырые данные: каждая строка - один месяц x один запрос\n{date_range_str}", title_fmt)
    data_headers = [
        "Дата\nМесяц",
        "Поисковый запрос\nquery",
        "Показы в месяц\nfrequency",
        "Скользящее среднее\nза 3 мес.\nfreq_ma3",
        "Темп роста, %\ngrowth_pct",
        "Волатильность\nvolatility\n🟢низкая 🟡средняя 🔴высокая",
        "Индекс сезонности\nseasonality_idx",
        "Индекс конкуренции\ncompetition_score\n🟢<0.40 🟡0.40-0.60 🔴>0.60",
    ]
    ws1.write_row(1, 0, data_headers, hdr_fmt)
    ws1.set_column(0, 0, 14)
    ws1.set_column(1, 1, 42)
    ws1.set_column(2, 7, 26)

    for ri, row_data in sheet1_df.iterrows():
        r = ri + 2
        ws1.write(r, 0, row_data["date"], date_fmt)
        ws1.write(r, 1, row_data["query"], txt_fmt)
        ws1.write(r, 2, row_data["frequency"], int_fmt)
        ws1.write(r, 3, round(row_data["freq_ma3"], 2), num_fmt)
        ws1.write(r, 4, round(row_data["growth_pct"], 2), pct_fmt)
        vol = round(row_data["volatility"], 2)
        q = row_data["query"]
        q_label = summary.loc[summary["query"] == q, "vol_label"].values
        vlabel = q_label[0] if len(q_label) > 0 else "Средняя"
        ws1.write(r, 5, vol, vol_fmt_h(vol, vlabel))
        seas = round(row_data["seasonality_idx"], 3)
        if seas > 1.2:
            sf = wb.add_format({"num_format": "0.000", "border": 1, "bg_color": "#F4B183", "font_size": 14})
        elif seas < 0.8:
            sf = wb.add_format({"num_format": "0.000", "border": 1, "bg_color": "#A9D18E", "font_size": 14})
        else:
            sf = num_fmt
        ws1.write(r, 6, seas, sf)
        score = round(row_data["competition_score"], 3)
        ws1.write(r, 7, score, comp_fmt(score))

    # ==========================================================
    # ЛИСТ 3: Сводка по запросам
    # ==========================================================
    ws2 = wb.add_worksheet("Сводка по запросам")
    ws2.merge_range(0, 0, 0, 9,
        f"Итоговые метрики, от популярных к нишевым\n{date_range_str}", title_fmt)
    sum_headers = [
        "Поисковый запрос\nquery",
        "Средняя частотность\nMean Frequency",
        "Суммарная частотность\nSum Frequency\n(емкость рынка)",
        "Стандартное отклонение\nStd Frequency",
        "Mean Growth (%)\nсредний за весь\nпериод",
        "Growth Momentum (%)\nсредний за\nпоследние 6 мес.",
        "Trend Slope\n(показов/мес)\nнаклон тренда",
        "Волатильность\nVolatility",
        "Индекс конкуренции\nCompetition Index",
    ]
    ws2.write_row(1, 0, sum_headers, hdr_fmt)
    ws2.set_column(0, 0, 42)
    ws2.set_column(1, 8, 22)

    for ri, row_data in summary.iterrows():
        r = ri + 2
        ws2.write(r, 0, row_data["query"], txt_fmt)
        ws2.write(r, 1, round(row_data["mean_frequency"], 1), int_fmt)
        ws2.write(r, 2, int(row_data["sum_frequency"]), int_fmt)
        ws2.write(r, 3, round(row_data["std_frequency"], 1), num_fmt)
        ws2.write(r, 4, round(row_data["mean_growth"], 2), num_fmt)
        ws2.write(r, 5, round(row_data["growth_momentum"], 2), num_fmt)
        ws2.write(r, 6, round(row_data["trend_slope"], 2), num_fmt)
        vol = round(row_data["mean_volatility"], 1)
        label = row_data["vol_label"]
        ws2.write(r, 7, vol, vol_fmt_h(vol, label))
        score = round(row_data["competition_score"], 3)
        ws2.write(r, 8, score, comp_fmt(score))

    # ==========================================================
    # ЛИСТ 4: График
    # ==========================================================
    ws3 = wb.add_worksheet("График")
    ws3.merge_range(0, 0, 0, 3,
        f"Динамика запросов\n{date_range_str}", title_fmt)
    rg = 1
    for img_path in [img_demand, img_ocean_mom, img_growing_mom]:
        if os.path.exists(img_path):
            ws3.insert_image(rg, 0, img_path, {"x_scale": 0.70, "y_scale": 0.70})
            rg += 20
    ws3.set_column(0, 0, 100)

    # ==========================================================
    # ЛИСТ 5: Голубые океаны (3 блока на одном листе)
    # ==========================================================
    ws4 = wb.add_worksheet("Голубые океаны")
    ws4.set_column(0, 0, 6)
    ws4.set_column(1, 1, 42)
    ws4.set_column(2, 4, 28)
    ws4.set_column(5, 5, 54)

    next_row = _write_ocean_block(ws4, blue_momentum, 0,
        "Блок 1: Низкая конкуренция + рост (momentum)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "growth_momentum")

    next_row = _write_ocean_block(ws4, blue_mean, next_row,
        "Блок 2: Низкая конкуренция + рост (mean)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "mean_growth")

    _write_ocean_block(ws4, blue_slope, next_row,
        "Блок 3: Низкая конкуренция + рост (slope)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "trend_slope")

    # ==========================================================
    # ЛИСТ 6: Растущие рынки (3 блока на одном листе)
    # ==========================================================
    ws5 = wb.add_worksheet("Растущие рынки")
    ws5.set_column(0, 0, 6)
    ws5.set_column(1, 1, 42)
    ws5.set_column(2, 5, 28)
    ws5.set_column(6, 6, 44)

    next_row = _write_growing_block(ws5, growing_momentum, 0,
        "Блок 1: Растущие рынки (momentum)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "growth_momentum")

    next_row = _write_growing_block(ws5, growing_mean, next_row,
        "Блок 2: Растущие рынки (mean)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "mean_growth")

    _write_growing_block(ws5, growing_slope, next_row,
        "Блок 3: Растущие рынки (slope)", date_range_str,
        hdr_fmt, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "trend_slope")

    # ── Закрытие книги ─────────────────────────────────────────
    wb.close()

    for tmp_img in [img_demand, img_ocean_mom, img_growing_mom]:
        if os.path.exists(tmp_img):
            os.remove(tmp_img)

    logger.info(f"Report saved: {output_path}")
    return output_path


# ─────────────────────────────────────────────────────────────────
# Инсайты в консоль
# ─────────────────────────────────────────────────────────────────
def print_insights(raw_data: pd.DataFrame) -> None:
    logger.info("=" * 60)
    logger.info("INSIGHTS")
    logger.info("=" * 60)

    s = raw_data.groupby("query").agg(
        sum_freq=("frequency", "sum"),
        growth_momentum=("growth_momentum", "first"),
        competition=("competition_score", "first"),
    ).reset_index()

    top5 = s.nlargest(5, "sum_freq")
    logger.info("TOP-5 DEMAND:")
    for _, r in top5.iterrows():
        logger.info(f"  {r['query']:<35s} {int(r['sum_freq']):>10,} shows")

    top3 = s.nlargest(3, "growth_momentum")
    logger.info("\nTOP-3 GROWTH (momentum):")
    for _, r in top3.iterrows():
        logger.info(f"  {r['query']:<35s} {r['growth_momentum']:>+8.2f}%")

    low = s[(s["competition"] < 0.4) & (s["growth_momentum"] > 0)] \
        .sort_values("growth_momentum", ascending=False)
    if not low.empty:
        logger.info("\nLOW COMPETITION + GROWTH (momentum):")
        for _, r in low.head(5).iterrows():
            logger.info(f"  {r['query']:<35s} momentum {r['growth_momentum']:>+6.2f}% | comp {r['competition']:.2f}")
    else:
        logger.info("\nNo queries with low competition + growth (momentum)")

    logger.info("=" * 60)


# ─────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────
def main() -> None:
    start = time.time()
    validate_env()

    date_from, date_to = _parse_date_range()

    raw = collect_all_data(date_from, date_to)
    analyzed = calculate_metrics(raw)
    create_excel_report(analyzed, date_from, date_to)
    print_insights(analyzed)

    elapsed = time.time() - start
    logger.info(f"Time: {elapsed:.1f}s")
    logger.info(f"Output: {os.path.abspath(OUTPUT_FILE)}")
    logger.info("Done!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("Interrupted")
        sys.exit(1)
    except Exception as exc:
        logger.error(f"Fatal error: {exc}")
        traceback.print_exc()
        sys.exit(1)