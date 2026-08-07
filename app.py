from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, jsonify
from flask import send_from_directory
import requests
import xml.etree.ElementTree as ET
from openpyxl.drawing.image import Image as XLImage
import openpyxl
from datetime import date, datetime
import hashlib
import re
import os
import secrets
from openpyxl import Workbook
import urllib.request
import json
from threading import Lock
from local_config import ADD_BOOK_PASSWORD, API_KEY, SECRET_KEY

EXCEL_FILE = "books.xlsx"
EXCEL_LOCK = Lock()

JOURNAL_COLUMNS = {
    'name': 1,
    'author': 3,
    'pages': 5,
    'categories': 7,
    'date_started': 11,
    'date_finished': 12,
    'journal_rating': 13,
    'romance': 14,
    'friendship': 15,
    'humor': 16,
    'heartbreak': 17,
    'plot': 18,
    'chemistry': 19,
    'thoughts': 20,
    'tropes': 21,
}

JOURNAL_HEADERS = [
    'Дата начала прочтения',
    'Дата окончания прочтения',
    'Рейтинг',
    'Романтично',
    'Дружба',
    'Юмор',
    'Стекло',
    'Сюжет',
    'Химия и напряжение',
    'Мои мысли',
    'Популярные тропы',
]

app = Flask(__name__)

app.secret_key = SECRET_KEY
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024

GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1/volumes" 

@app.route('/favicon.ico')
def favicon():
    return send_from_directory('static', 'favicon.ico')

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        sheet = wb.active
        sheet.append(['Название книги', 'ID в Google Books', 'Автор(ы)', 'Год издания', 
                      'Количество страниц', 'Рейтинг', 'Жанры/Категории', 'Добавил пользователь', 'Картинка', 'Избранное'] + JOURNAL_HEADERS)
        sheet.row_dimensions[1].height = 100
        sheet.column_dimensions['J'].width = 15
        wb.save(EXCEL_FILE)

init_excel()

def make_cover_filename(book_name, book_id, extension='jpg'):
    """Формирует единое имя обложки: название_книги_ID.расширение."""
    safe_name = re.sub(r'[\\/*?:"<>|\.]', '_', str(book_name or ''))
    safe_name = re.sub(r'\s+', '_', safe_name)
    safe_id = re.sub(r'[\\/*?:"<>|\.\s]', '_', str(book_id or ''))
    return f'{safe_name}_{safe_id}.{extension}'

def find_cover_image(book_name, book_id):
    """Возвращает web-путь обложки, имя которой соответствует книге и её ID."""
    for extension in ('jpg', 'jpeg', 'png', 'webp'):
        filename = make_cover_filename(book_name, book_id, extension)
        image_path = os.path.join('static', 'images', filename)
        if os.path.exists(image_path):
            return f'images/{filename}'
    return None

def make_manual_book_id(book_name):
    """Создаёт URL-безопасный уникальный ID для книги, добавленной вручную."""
    normalized_name = re.sub(r'\s+', ' ', book_name.strip()).casefold()
    name_hash = hashlib.sha256(normalized_name.encode('utf-8')).hexdigest()[:8]
    return f'manual-{name_hash}-{secrets.token_hex(3)}'

def save_uploaded_cover(uploaded_file, book_name, book_id):
    """Проверяет пользовательскую обложку и сохраняет её в общем формате JPEG."""
    if not uploaded_file or not uploaded_file.filename:
        return None

    extension = uploaded_file.filename.rsplit('.', 1)[-1].lower() if '.' in uploaded_file.filename else ''
    if extension not in {'jpg', 'jpeg', 'png', 'webp'}:
        raise ValueError('Допустимы только изображения JPG, PNG или WebP')

    from PIL import Image as PILImage
    from PIL import ImageOps

    try:
        image = PILImage.open(uploaded_file.stream)
        image.verify()
        uploaded_file.stream.seek(0)
        image = PILImage.open(uploaded_file.stream)
        image = ImageOps.exif_transpose(image)
    except Exception as error:
        raise ValueError('Не удалось прочитать загруженную обложку') from error

    if image.mode in ('RGBA', 'P', 'LA'):
        background = PILImage.new('RGB', image.size, (255, 255, 255))
        if image.mode == 'P':
            image = image.convert('RGBA')
        background.paste(image, mask=image.getchannel('A') if image.mode in ('RGBA', 'LA') else None)
        image = background
    elif image.mode != 'RGB':
        image = image.convert('RGB')

    static_img_dir = 'static/images'
    os.makedirs(static_img_dir, exist_ok=True)
    filename = make_cover_filename(book_name, book_id)
    image.save(os.path.join(static_img_dir, filename), 'JPEG', quality=95, optimize=True)
    return f'images/{filename}'

def download_and_insert_image(sheet, row_num, image_url, book_name, book_id):
    """Скачивает картинку обложки и вставляет в Excel"""
    try:
        from PIL import Image as PILImage
        import io

        static_img_dir = 'static/images'
        if not os.path.exists(static_img_dir):
            os.makedirs(static_img_dir)
        
        img_filename = make_cover_filename(book_name, book_id)
        img_path = os.path.join(static_img_dir, img_filename)
        
        session_req = requests.Session()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = session_req.get(image_url, headers=headers, timeout=(10, 60))
        response.raise_for_status()
        
        img_pil = PILImage.open(io.BytesIO(response.content))
        
        if img_pil.mode in ('RGBA', 'P', 'LA'):
            background = PILImage.new('RGB', img_pil.size, (255, 255, 255))
            if img_pil.mode == 'P':
                img_pil = img_pil.convert('RGBA')
            background.paste(
                img_pil, 
                mask=img_pil.split()[-1] if img_pil.mode in ('RGBA', 'LA') else None
            )
            img_pil = background
        elif img_pil.mode != 'RGB':
            img_pil = img_pil.convert('RGB')
        
        img_pil.save(img_path, 'JPEG', quality=95, optimize=True)
        
        xl_img = XLImage(img_path)
        xl_img.width = 60
        xl_img.height = 80
        sheet.add_image(xl_img, f'I{row_num}')
        sheet.row_dimensions[row_num].height = 90
        sheet.column_dimensions['I'].width = 12
        
        return f'images/{img_filename}'
        
    except Exception as e:
        print(f"Ошибка при загрузке обложки: {e}")
        return None

def get_all_categories():
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    categories = set()
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        if row[6]:
            for category in row[6].split("; "):
                if category and category.strip():
                    categories.add(category.strip())
    return sorted(list(categories))
    
def get_all_authors():
    """Собирает всех уникальных авторов из Excel файла"""
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    authors = set()
    for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):
        if row[2]:  # колонка 3 (столбец C) - авторы
            # Разделяем, если несколько авторов через запятую
            for author in str(row[2]).split(','):
                author = author.strip()
                if author and author != 'Автор не указан':
                    authors.add(author)
    return sorted(list(authors))

@app.route('/')
def index():
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    books_count = 0
    featured_books = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[0]
        if not name:
            continue

        books_count += 1
        book_id = row[1]
        image_path = find_cover_image(name, book_id)
        if image_path and len(featured_books) < 2:
            featured_books.append({
                'name': name,
                'id': book_id,
                'image_path': image_path,
            })

    workbook.close()
    return render_template('index.html', books_count=books_count, featured_books=featured_books)

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)

@app.route('/books')
def books():
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    books_list = []
    for row_idx in range(2, sheet.max_row + 1):
        name = sheet.cell(row=row_idx, column=1).value
        if name:
            book_id = sheet.cell(row=row_idx, column=2).value
            author = sheet.cell(row=row_idx, column=3).value
            year = sheet.cell(row=row_idx, column=4).value
            pages = sheet.cell(row=row_idx, column=5).value
            rating = sheet.cell(row=row_idx, column=6).value
            categories = sheet.cell(row=row_idx, column=7).value
            added_by = sheet.cell(row=row_idx, column=8).value
            favorite = sheet.cell(row=row_idx, column=10).value
            
            image_filename = find_cover_image(name, book_id)
            
            books_list.append({
                'row_index': row_idx,
                'name': name,
                'id': book_id,
                'author': author,
                'year': year,
                'pages': pages,
                'rating': rating,
                'categories': categories,
                'added_by': added_by,
                'image_path': image_filename,
                'favorite': str(favorite).strip() if favorite else ''
            })
    
    return render_template('books.html', books=books_list)

def format_excel_date(value):
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    if value:
        return str(value)
    return ''

def parse_form_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None

def parse_score(value):
    if value in (None, ''):
        return None
    try:
        score = int(value)
    except (TypeError, ValueError):
        return None
    return max(0, min(5, score))

def parse_non_negative_int(value):
    if value in (None, ''):
        return None
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return None

@app.route('/books/<book_id>', methods=['GET', 'POST'])
def book_journal(book_id):
    if not session.get('password_ok'):
        if request.method == 'POST':
            password = request.form.get('password')
            if password == ADD_BOOK_PASSWORD:
                session['password_ok'] = True
                flash('Пароль верный! Дневник книги открыт', 'success')
                return redirect(url_for('book_journal', book_id=book_id))

            flash('Неверный пароль!', 'error')

        return render_template(
            'password_check.html',
            access_title='Доступ к читательскому дневнику',
            access_heading='Открыть читательский дневник',
            access_intro='Введите пароль владельца, чтобы просматривать и редактировать записи о книге.',
            access_copy='После подтверждения откроется читательский дневник выбранной книги.',
        )

    with EXCEL_LOCK:
        workbook = openpyxl.load_workbook(EXCEL_FILE)
        sheet = workbook.active

        row_index = next(
            (
                row
                for row in range(2, sheet.max_row + 1)
                if str(sheet.cell(row=row, column=2).value or '') == str(book_id)
            ),
            None,
        )

        if row_index is None:
            workbook.close()
            flash('Книга не найдена', 'error')
            return redirect(url_for('books'))

        if request.method == 'POST':
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['name'], value=request.form.get('name', '').strip())
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['author'], value=request.form.get('author', '').strip())
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['pages'], value=parse_non_negative_int(request.form.get('pages')))
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['categories'], value=request.form.get('categories', '').strip())
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['date_started'], value=parse_form_date(request.form.get('date_started')))
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['date_finished'], value=parse_form_date(request.form.get('date_finished')))

            for field in ('journal_rating', 'romance', 'friendship', 'humor', 'heartbreak', 'plot', 'chemistry'):
                sheet.cell(row=row_index, column=JOURNAL_COLUMNS[field], value=parse_score(request.form.get(field)))

            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['thoughts'], value=request.form.get('thoughts', '').strip())
            sheet.cell(row=row_index, column=JOURNAL_COLUMNS['tropes'], value=request.form.get('tropes', '').strip())

            workbook.save(EXCEL_FILE)
            workbook.close()
            flash('Дневник книги сохранён', 'success')
            return redirect(url_for('books'))

        journal = {
            field: sheet.cell(row=row_index, column=column).value
            for field, column in JOURNAL_COLUMNS.items()
        }
        journal['date_started'] = format_excel_date(journal['date_started'])
        journal['date_finished'] = format_excel_date(journal['date_finished'])
        journal['image_path'] = find_cover_image(journal['name'], book_id)

        workbook.close()

    return render_template('book_journal.html', book=journal, book_id=book_id)

@app.route('/export_excel')
def export_excel():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name='books.xlsx')
    else:
        flash('Файл не найден', 'error')
        return redirect(url_for('index'))

@app.route('/add_book', methods=['GET', 'POST'])
def add_book():
    if session.get('password_ok'):
        return render_template('add_book.html')
    
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADD_BOOK_PASSWORD:
            session['password_ok'] = True
            flash('Пароль верный! Теперь можно добавлять книги', 'success')
            return render_template('add_book.html')
        else:
            flash('Неверный пароль!', 'error')
            return render_template('password_check.html')
    
    return render_template('password_check.html')

@app.route('/add_book_manual', methods=['POST'])
def add_book_manual():
    if not session.get('password_ok'):
        flash('Сначала подтвердите пароль владельца', 'error')
        return redirect(url_for('add_book'))

    title = request.form.get('manual_title', '').strip()
    cover = request.files.get('manual_cover')
    if not title:
        flash('Укажите название книги', 'error')
        return redirect(url_for('add_book'))

    cover_path = None
    workbook = None
    try:
        with EXCEL_LOCK:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
            sheet = workbook.active

            existing_ids = {
                str(sheet.cell(row=row, column=2).value or '')
                for row in range(2, sheet.max_row + 1)
            }
            book_id = make_manual_book_id(title)
            while book_id in existing_ids:
                book_id = make_manual_book_id(title)

            if cover and cover.filename:
                cover_path = save_uploaded_cover(cover, title, book_id)

            next_row = sheet.max_row + 1
            sheet.cell(row=next_row, column=1, value=title)
            sheet.cell(row=next_row, column=2, value=book_id)
            sheet.cell(row=next_row, column=3, value='')
            sheet.cell(row=next_row, column=4, value='')
            sheet.cell(row=next_row, column=5, value='')
            sheet.cell(row=next_row, column=6, value='')
            sheet.cell(row=next_row, column=7, value='')
            sheet.cell(row=next_row, column=8, value='web_user')
            sheet.cell(row=next_row, column=9, value=cover_path or '')
            sheet.cell(row=next_row, column=10, value='')
            workbook.save(EXCEL_FILE)

        flash(f"Книга '{title}' успешно добавлена!", 'success')
        return redirect(url_for('books'))
    except ValueError as error:
        flash(str(error), 'error')
    except Exception as error:
        flash(f'Не удалось добавить книгу: {error}', 'error')
    finally:
        if workbook is not None:
            workbook.close()

    if cover_path:
        cover_file = os.path.join('static', cover_path)
        if os.path.exists(cover_file):
            os.remove(cover_file)
    return redirect(url_for('add_book'))

@app.route('/search_books', methods=['POST'])
def search_books():
    search_text = request.form.get('book_name')
    if search_text:
        search_url = f"https://www.googleapis.com/books/v1/volumes?q={search_text.replace(' ', '+')}&key={API_KEY}&langRestrict=ru&maxResults=20"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        try:
            response = requests.get(search_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            
            # 🔧 ДОБАВЬТЕ ЭТУ ЧАСТЬ - ОБРАБОТКА РЕЗУЛЬТАТОВ
            results = []
            for item in data.get('items', []):
                volume_info = item.get('volumeInfo', {})
                title = volume_info.get('title', 'Без названия')
                book_id = item.get('id')
                authors = ', '.join(volume_info.get('authors', ['Автор не указан']))
                
                # Получаем обложку
                image_links = volume_info.get('imageLinks', {})
                thumbnail_url = image_links.get('thumbnail', '')
                
                # Год издания
                published_date = volume_info.get('publishedDate', '')
                year = published_date[:4] if published_date else ''
                
                results.append({
                    'name': title,
                    'id': book_id,
                    'authors': authors,
                    'year': year,
                    'cover_url': thumbnail_url
                })
            
            return render_template('search_results.html', results=results, search_text=search_text)
            # КОНЕЦ ДОБАВЛЕННОЙ ЧАСТИ
            
        except requests.exceptions.Timeout:
            flash('Превышено время ожидания ответа от Google Books API', 'error')
        except requests.exceptions.ConnectionError:
            flash('Ошибка подключения к Google Books API. Проверьте интернет.', 'error')
        except requests.exceptions.HTTPError as e:
            flash(f'Ошибка HTTP: {e}. API может быть временно недоступен.', 'error')
        except Exception as e:
            flash(f'Неизвестная ошибка: {str(e)}', 'error')
            
        return redirect(url_for('add_book'))
    
    return redirect(url_for('add_book'))

@app.route('/add_book_by_id', methods=['POST'])
def add_book_by_id():
    book_id = request.form.get('book_id')
    
    if not book_id:
        flash('Пожалуйста, введите корректный ID книги', 'error')
        return redirect(url_for('add_book'))
    
    # Получаем данные о книге из Google Books API
    book_url = f"https://www.googleapis.com/books/v1/volumes/{book_id}"
    
    params = {'key': API_KEY}
    
    try:
        response = requests.get(book_url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            volume_info = data.get('volumeInfo', {})
            
            title = volume_info.get('title', 'Без названия')
            authors = ', '.join(volume_info.get('authors', ['Автор не указан']))
            published_date = volume_info.get('publishedDate', '')
            year = published_date[:4] if published_date else ''
            page_count = volume_info.get('pageCount', '')
            rating = volume_info.get('averageRating', '')
            categories = '; '.join(volume_info.get('categories', []))
            
            # Получаем обложку
            image_links = volume_info.get('imageLinks', {})
            image_url = image_links.get('medium') or image_links.get('thumbnail')
            
            # Открываем Excel файл один раз
            workbook = openpyxl.load_workbook(EXCEL_FILE)
            sheet = workbook.active
            
            next_row = sheet.max_row + 1
            
            # Добавляем все данные
            sheet.cell(row=next_row, column=1, value=title)
            sheet.cell(row=next_row, column=2, value=book_id)
            sheet.cell(row=next_row, column=3, value=authors)
            sheet.cell(row=next_row, column=4, value=year)
            sheet.cell(row=next_row, column=5, value=page_count)
            sheet.cell(row=next_row, column=6, value=rating)
            sheet.cell(row=next_row, column=7, value=categories)
            sheet.cell(row=next_row, column=8, value='web_user')
            sheet.cell(row=next_row, column=9, value='')  # Картинка
            sheet.cell(row=next_row, column=10, value='')  # Избранное
            
            # Добавляем картинку если есть
            if image_url:
                try:
                    download_and_insert_image(sheet, next_row, image_url, title, book_id)
                except Exception as img_error:
                    print(f"Ошибка картинки: {img_error}")
                    # Продолжаем, даже если картинка не добавилась
            
            # Сохраняем файл только один раз в конце
            workbook.save(EXCEL_FILE)
            
            flash(f"Книга '{title}' успешно добавлена!", 'success')
            return redirect(url_for('books'))
        else:
            flash(f'Ошибка API: статус {response.status_code}', 'error')
            
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')
        print(f"Детали: {e}")
        
    return redirect(url_for('add_book'))

@app.route('/toggle_favorite', methods=['POST'])
def toggle_favorite():
    data = request.get_json()
    book_name = data.get('book_name')
    favorite_value = data.get('favorite')
    
    if not book_name:
        return jsonify({'success': False, 'error': 'Не указана книга'}), 400
    
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    
    found = False
    for row_idx in range(2, sheet.max_row + 1):
        if sheet.cell(row=row_idx, column=1).value == book_name:
            current_value = sheet.cell(row=row_idx, column=10).value or ''
            
            if favorite_value is None:
                new_value = 'нет' if current_value == 'да' else 'да'
            else:
                new_value = favorite_value
            
            sheet.cell(row=row_idx, column=10, value=new_value)
            found = True
            break
    
    if found:
        workbook.save(EXCEL_FILE)
        return jsonify({'success': True, 'favorite': new_value})
    else:
        return jsonify({'success': False, 'error': 'Книга не найдена'}), 404

@app.route('/filter')
def filter_books():
    return render_template('filter.html')

@app.route('/filter/author', methods=['GET', 'POST'])
def filter_by_author():
    authors = get_all_authors()  # Получаем список авторов из Excel
    
    if request.method == 'POST':
        author = request.form.get('author')
        if author:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
            sheet = workbook.active
            
            matching_books = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[2]:  # колонка с авторами
                    book_authors = row[2].split(',')
                    for book_author in book_authors:
                        if author.lower() in book_author.strip().lower():
                            matching_books.append(row[0])
                            break
            
            return render_template('filter_results.html', games=matching_books, 
                                 filter_text=f"по автору '{author}'")
    
    return render_template('filter_author.html', authors=authors)

@app.route('/filter/category', methods=['GET', 'POST'])
def filter_by_category():
    categories = get_all_categories()
    
    if request.method == 'POST':
        category = request.form.get('category')
        if category:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
            sheet = workbook.active
            
            matching_books = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[6]:  # колонка с жанрами
                    book_categories = row[6].split("; ")
                    if category in book_categories:
                        matching_books.append(row[0])  # название книги
            
            return render_template('filter_results.html', games=matching_books, 
                                 filter_text=f"в жанре '{category}'")
    
    return render_template('filter_category.html', categories=categories)

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)

if __name__ == '__main__':
    app.run(
        host='0.0.0.0', 
        port=5001,  # Другой порт, чтобы не конфликтовать с сайтом настолок
        debug=True
    )
