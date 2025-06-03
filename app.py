import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import threading
import time
import os
import subprocess
import sys
from datetime import datetime
import random
import requests
from urllib.parse import urlparse
import shutil
import uuid


class AvitoParser:
    def __init__(self, progress_callback=None):
        self.driver = None
        self.progress_callback = progress_callback
        self.stop_flag = False
        self.base_dir = ""
        self.query = ""

    def init_driver(self):
        """Инициализация веб-драйвера с настройками"""
        chrome_options = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

    def update_progress(self, message, value=None):
        """Обновление прогресса"""
        if self.progress_callback:
            self.progress_callback(message, value)

    def create_result_directory(self, query):
        """Создание директории для результатов"""
        self.query = query
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        dir_name = f"{query}_{now}".replace(" ", "_")[:50]
        self.base_dir = os.path.join(os.getcwd(), dir_name)
        
        os.makedirs(os.path.join(self.base_dir, "excel_docs"), exist_ok=True)
        os.makedirs(os.path.join(self.base_dir, "photos"), exist_ok=True)
        
        return os.path.join(self.base_dir, "excel_docs")

    def download_image(self, url, filepath):
        """Скачивание одного изображения"""
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': 'https://www.avito.ru/'
            }
            
            response = requests.get(url, headers=headers, stream=True, timeout=30)
            if response.status_code == 200:
                with open(filepath, 'wb') as f:
                    response.raw.decode_content = True
                    shutil.copyfileobj(response.raw, f)
                return True
        except Exception as e:
            print(f"Ошибка загрузки изображения {url}: {str(e)}")
        return False

    def download_images(self, ad_id, image_urls):
        """Скачивание изображений объявления"""
        if not image_urls:
            return ""
            
        ad_photo_dir = os.path.join(self.base_dir, "photos", f"ad_{ad_id}")
        os.makedirs(ad_photo_dir, exist_ok=True)
        
        saved_images = []
        for i, url in enumerate(image_urls[:5]):  # Ограничим 5 фото на объявление
            try:
                # Очищаем URL от параметров
                clean_url = url.split('?')[0]
                
                # Определяем расширение файла
                ext = os.path.splitext(urlparse(clean_url).path)[1].lower()
                if not ext or ext not in ['.jpg', '.jpeg', '.png', '.webp']:
                    ext = '.jpg'
                
                filename = f"photo_{i+1}{ext}"
                filepath = os.path.join(ad_photo_dir, filename)
                
                if self.download_image(clean_url, filepath):
                    saved_images.append(f"photos/ad_{ad_id}/{filename}")
            except Exception as e:
                print(f"Ошибка обработки изображения {url}: {str(e)}")
        
        return "; ".join(saved_images) if saved_images else ""

    def get_ad_images(self):
        """Получение URL изображений объявления"""
        try:
            # Ждем загрузки галереи
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-marker="image-frame/image-wrapper"]'))
            )
            
            # Получаем все элементы изображений
            image_elements = self.driver.find_elements(By.CSS_SELECTOR, '[data-marker="image-frame/image-wrapper"] img')
            
            # Собираем URL изображений
            image_urls = []
            for img in image_elements:
                src = img.get_attribute("src")
                if src and src.startswith('http'):
                    # Пытаемся получить URL в лучшем качестве
                    high_quality_url = src.replace('64x48', '640x480').replace('128x96', '640x480')
                    image_urls.append(high_quality_url)
            
            return list(set(image_urls))  # Удаляем дубликаты
            
        except Exception as e:
            print(f"Ошибка при получении изображений: {str(e)}")
            return []

    def is_page_available(self):
        """Проверка доступности страницы"""
        try:
            error_messages = [
                'Такого товара не существует',
                'по вашему запросу ничего не найдено',
                'Произошла ошибка'
            ]
            
            for message in error_messages:
                elements = self.driver.find_elements(By.XPATH, f'//*[contains(text(), "{message}")]')
                if elements:
                    return False
            return True
        except:
            return True

    def search_ads(self, query, region_id=621540, max_pages=1):
        """Поиск объявлений"""
        if not self.driver:
            self.init_driver()

        results = []
        try:
            # Создаем директорию для результатов
            excel_dir = self.create_result_directory(query)
            
            for page in range(1, max_pages + 1):
                if self.stop_flag:
                    break

                self.update_progress(f"Обработка страницы {page}/{max_pages}", (page/max_pages)*50)
                
                url = f"https://www.avito.ru/all?q={query.replace(' ', '+')}"
                if page > 1:
                    url += f"&p={page}"
                
                self.driver.get(url)
                time.sleep(random.uniform(2, 4))

                if not self.is_page_available():
                    self.update_progress("Avito сообщает, что товар не существует")
                    return []

                try:
                    WebDriverWait(self.driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '[data-marker="item"]'))
                    )
                except:
                    self.update_progress("Не удалось найти объявления на странице")
                    continue

                self.scroll_page()
                ads = self.driver.find_elements(By.CSS_SELECTOR, '[data-marker="item"]')
                
                for i, ad in enumerate(ads):
                    if self.stop_flag:
                        break
                        
                    self.update_progress(f"Обработка объявления {i+1}/{len(ads)}", 50 + (i/len(ads))*50)
                    
                    if ad_data := self.parse_ad(ad):
                        results.append(ad_data)
                
                time.sleep(random.uniform(2, 4))

        except Exception as e:
            self.update_progress(f"Ошибка: {str(e)}")
            return []
        
        return results

    def scroll_page(self):
        """Прокрутка страницы"""
        for _ in range(3):
            self.driver.execute_script("window.scrollBy(0, 500);")
            time.sleep(random.uniform(0.5, 1.5))

    def parse_ad(self, ad_element):
        """Парсинг одного объявления с текстом и фото"""
        try:
            # Основная информация
            title = ad_element.find_element(By.CSS_SELECTOR, '[itemprop="name"]').text
            link = ad_element.find_element(By.CSS_SELECTOR, 'a[data-marker="item-title"]').get_attribute("href")
            price_elem = ad_element.find_element(By.CSS_SELECTOR, '[itemprop="price"]')
            price = price_elem.get_attribute("content") if price_elem else "Цена не указана"
            
            # Генерируем уникальный ID для объявления
            ad_id = str(uuid.uuid4())[:8]
            
            # Переходим на страницу объявления
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.get(link)
            time.sleep(random.uniform(2, 3))
            
            # Парсинг описания
            description = ""
            try:
                desc_elem = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[itemprop="description"], [data-marker="item-view/item-description"]'))
                )
                description = desc_elem.text
            except:
                description = "Описание не найдено"
            
            # Парсинг фото
            image_urls = self.get_ad_images()
            
            # Скачиваем фото
            photos_path = self.download_images(ad_id, image_urls)
            
            # Закрываем вкладку с объявлением
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            
            return {
                "id": ad_id,
                "title": title,
                "price": price,
                "description": description,
                "photos": photos_path,
                "link": link
            }
            
        except Exception as e:
            print(f"Ошибка парсинга объявления: {str(e)}")
            try:
                if len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
            except:
                pass
            return None

    def save_to_excel(self, data, filename):
        """Сохранение в Excel"""
        try:
            self.update_progress("Сохранение в Excel...", 90)
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Результаты"
            
            # Заголовок с датой
            now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            ws.append(["Дата и время сохранения:", now])
            ws.append(["Запрос:", self.query])
            ws.append([])
            
            # Заголовки таблицы
            headers = ["ID", "Название", "Цена (руб)", "Описание", "Фотографии", "Ссылка"]
            ws.append(headers)
            
            # Форматирование заголовков
            bold_font = Font(bold=True)
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=4, column=col)
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center")
            
            # Заполнение данных
            for i, item in enumerate(data, start=5):
                ws.append([
                    item["id"],
                    item["title"],
                    item["price"],
                    item["description"],
                    item["photos"],
                    item["link"]
                ])
            
            # Автоподбор ширины столбцов
            for col in ws.columns:
                column = col[0].column_letter
                max_length = 0
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width
            
            # Сохранение файла
            excel_dir = os.path.join(self.base_dir, "excel_docs")
            final_filename = os.path.join(excel_dir, os.path.basename(filename))
            
            if os.path.exists(final_filename):
                base, ext = os.path.splitext(final_filename)
                final_filename = f"{base}_{datetime.now().strftime('%H%M%S')}{ext}"
            
            wb.save(final_filename)
            self.update_progress("Готово!", 100)
            return True
            
        except Exception as e:
            self.update_progress(f"Ошибка сохранения: {str(e)}")
            return False

    def close(self):
        """Закрытие драйвера"""
        if self.driver:
            self.driver.quit()


class AvitoParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Avito Parser Pro")
        self.root.geometry("800x450")
        self.parser = None
        self.stop_parsing = False
        
        self.create_widgets()
        self.load_regions()
        
    def create_widgets(self):
        """Создание интерфейса"""
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TFrame", padding=10)
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        input_frame = ttk.LabelFrame(main_frame, text="Параметры поиска", padding=10)
        input_frame.pack(fill="x", pady=5)
        
        settings_frame = ttk.LabelFrame(main_frame, text="Дополнительные настройки", padding=10)
        settings_frame.pack(fill="x", pady=5)
        
        ttk.Label(input_frame, text="Введите запрос:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.query_entry = ttk.Entry(input_frame, width=40)
        self.query_entry.grid(row=0, column=1, padx=5, pady=5)
        self.query_entry.insert(0, "айфон 13 128 гб")
        
        ttk.Label(input_frame, text="Имя файла для сохранения:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.filename_entry = ttk.Entry(input_frame, width=40)
        self.filename_entry.grid(row=1, column=1, padx=5, pady=5)
        self.filename_entry.insert(0, "avito_results.xlsx")
        ttk.Button(input_frame, text="Обзор...", command=self.browse_file).grid(row=1, column=2, padx=5)
        
        ttk.Label(settings_frame, text="Регион:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.region_combobox = ttk.Combobox(settings_frame, values=[], width=37)
        self.region_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.region_combobox.set("Москва")
        
        ttk.Label(settings_frame, text="Кол-во страниц:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.pages_spinbox = ttk.Spinbox(settings_frame, from_=1, to=10, width=5)
        self.pages_spinbox.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.pages_spinbox.set(1)
        
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill="x", pady=10)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.pack(fill="x")
        
        self.progress_label = ttk.Label(self.progress_frame, text="Готов к работе")
        self.progress_label.pack()
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        self.start_btn = ttk.Button(btn_frame, text="Начать парсинг", command=self.start_parsing)
        self.start_btn.pack(side="left", padx=10)
        
        self.stop_btn = ttk.Button(btn_frame, text="Остановить", command=self.stop_parsing_process, state="disabled")
        self.stop_btn.pack(side="left", padx=10)
        
    def load_regions(self):
        """Загрузка списка регионов"""
        regions = {
            621540: "Москва",
            621630: "Санкт-Петербург",
            622640: "Новосибирск",
            621580: "Екатеринбург",
            621570: "Казань"
        }
        self.region_combobox["values"] = list(regions.values())
        self.regions_dict = {v: k for k, v in regions.items()}
        
    def browse_file(self):
        """Выбор файла для сохранения"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Сохранить результаты как"
        )
        if filename:
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, filename)
    
    def update_progress(self, message, value=None):
        """Обновление прогресса"""
        if value is not None:
            self.progress_bar["value"] = value
        self.progress_label["text"] = message
        self.root.update_idletasks()
    
    def start_parsing(self):
        """Запуск парсинга"""
        query = self.query_entry.get()
        filename = self.filename_entry.get()
        region_name = self.region_combobox.get()
        pages = int(self.pages_spinbox.get())
        
        if not query:
            messagebox.showerror("Ошибка", "Введите поисковый запрос!")
            return
        
        if not filename:
            messagebox.showerror("Ошибка", "Укажите файл для сохранения!")
            return
        
        region_id = self.regions_dict.get(region_name, 621540)
        
        self.start_btn["state"] = "disabled"
        self.stop_btn["state"] = "normal"
        self.stop_parsing = False
        
        threading.Thread(
            target=self.run_parsing,
            args=(query, filename, region_id, pages),
            daemon=True
        ).start()

    def open_result_folder(self, path):
        """Кросс-платформенное открытие папки"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS, Linux
                if sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', path])
                else:  # Linux
                    subprocess.run(['xdg-open', path])
        except Exception as e:
            print(f"Не удалось открыть папку: {str(e)}")
            messagebox.showinfo("Информация", f"Результаты сохранены в:\n{path}")
    
    def run_parsing(self, query, filename, region_id, pages):
        """Основная логика парсинга"""
        self.parser = AvitoParser(progress_callback=self.update_progress)
        
        try:
            results = self.parser.search_ads(query, region_id, pages)
            
            if results and not self.stop_parsing:
                if self.parser.save_to_excel(results, filename):
                    messagebox.showinfo("Успех", f"Сохранено {len(results)} объявлений")
                    # Кросс-платформенное открытие папки с результатами
                    result_dir = os.path.dirname(os.path.join(self.parser.base_dir, "excel_docs", filename))
                    self.open_result_folder(result_dir)
                else:
                    messagebox.showerror("Ошибка", "Не удалось сохранить результаты")
            elif not self.stop_parsing:
                messagebox.showinfo("Информация", "Объявления не найдены")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            self.parser.close()
            self.start_btn["state"] = "normal"
            self.stop_btn["state"] = "disabled"
    
    def stop_parsing_process(self):
        """Остановка парсинга"""
        self.stop_parsing = True
        if self.parser:
            self.parser.stop_flag = True
        self.update_progress("Остановка...")
        self.stop_btn["state"] = "disabled"


if __name__ == "__main__":
    root = tk.Tk()
    app = AvitoParserApp(root)
    root.mainloop()
