#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook

logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

SEARCH_URL = "https://old.bankrot.fedresurs.ru/TradeList.aspx"
REGION_NAME = "Ленинградская область"
KEYWORD = "кадастровый номер 47"
TRADE_TYPE_VALUE = "1"  # Открытый аукцион
MAX_LOTS = 25


def init_driver():
    logger.info("🧠 Запускаем браузер...")
    opts = webdriver.ChromeOptions()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    })
    driver.maximize_window()
    return driver


def apply_filters(driver):
    wait = WebDriverWait(driver, 20)
    driver.get(SEARCH_URL)
    logger.info("🔎 Открыли страницу поиска")

    # Регион
    sel = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphBody_ucRegion_ddlBoundList")))
    for o in sel.find_elements(By.TAG_NAME, "option"):
        if REGION_NAME in o.text:
            o.click()
            break
    logger.info(f"✅ Регион: {REGION_NAME}")

    # Вид торгов
    sel = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphBody_ucTradeType_ddlBoundList")))
    for o in sel.find_elements(By.TAG_NAME, "option"):
        if o.get_attribute("value") == TRADE_TYPE_VALUE:
            o.click()
            break
    logger.info("✅ Вид торгов: Открытый аукцион")

    # Ключевые слова
    kw = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cphBody_tbTradeObject")))
    kw.clear()
    kw.send_keys(KEYWORD)
    logger.info(f"✅ Ключевые слова: '{KEYWORD}'")

    # Поиск
    driver.find_element(By.ID, "ctl00_cphBody_btnTradeSearch").click()
    logger.info("▶️ Поиск выполнен — ждём обновления таблицы…")
    wait.until(EC.presence_of_element_located((By.ID, "ctl00_cphBody_gvTradeList")))
    logger.info("📋 Таблица результатов загружена")


def collect_urls(driver, limit=MAX_LOTS):
    wait = WebDriverWait(driver, 20)
    urls = []
    page = 1

    while len(urls) < limit:
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_cphBody_gvTradeList")))
        logger.info(f"--- Собираем URL-ы, страница {page} ({len(urls)}/{limit})")
        for a in driver.find_elements(By.XPATH,
                "//table[@id='ctl00_cphBody_gvTradeList']//a[contains(@href,'TradeCard.aspx')]"):
            h = a.get_attribute("href")
            if h not in urls:
                urls.append(h)
                if len(urls) >= limit:
                    break

        if len(urls) >= limit:
            break

        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(text(), '>')]")
            nxt.click()
            logger.info("▶️ Переходим на следующую страницу…")
            wait.until(EC.staleness_of(nxt))
            page += 1
            time.sleep(1)
        except Exception:
            logger.info("ℹ️ Дальше страниц нет, или кнопка “>” не найдена")
            break

    urls = urls[:limit]
    logger.info(f"ℹ️ Всего URL-ов собрано: {len(urls)}")
    return urls


def scrape_descriptions(driver, urls):
    wait = WebDriverWait(driver, 10)
    data = []

    for url in urls:
        logger.info(f"📝 Обрабатываю: {url}")
        driver.get(url)
        time.sleep(1)

        # находим блок предмета торгов
        block = None
        try:
            block = driver.find_element(By.CSS_SELECTOR, "td[id$='tdTradeObject']")
        except Exception:
            logger.warning("   ⚠️ Не найден блок предмета торгов")
        desc = ""

        if block:
            # ищем внутри блока ссылку подробнее
            try:
                more = block.find_element(By.XPATH, ".//a[contains(.,'подробнее')]")
                onclick = more.get_attribute("onclick") or ""
                m = re.search(r"openNewWin\('([^']+)'", onclick)
                if m:
                    detail_url = "https://old.bankrot.fedresurs.ru" + m.group(1)
                    driver.execute_script("window.open(arguments[0]);", detail_url)
                    driver.switch_to.window(driver.window_handles[-1])
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "td[style*='padding']")))
                    desc = driver.find_element(By.CSS_SELECTOR, "td[style*='padding']").text.strip()
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    logger.info("   → Описание из подробной страницы")
                else:
                    raise Exception
            except Exception:
                # fallback — просто текст внутри блока
                try:
                    desc = block.text.strip()
                    logger.info("   → Описание из текста блока")
                except Exception:
                    logger.warning("   ⚠️ Описание не найдено")

        data.append({"URL": url, "Description": desc})

    return data


def save_to_excel(rows, path="lot_data.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Открытый аукцион"
    ws.append(["URL", "Description"])
    for r in rows:
        ws.append([r["URL"], r["Description"]])
    wb.save(path)
    logger.info(f"💾 Сохранено в «{path}»")


def main():
    driver = init_driver()
    try:
        apply_filters(driver)
        urls = collect_urls(driver)
        recs = scrape_descriptions(driver, urls)
        save_to_excel(recs)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
