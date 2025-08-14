import os
import random
import pandas as pd
import base64
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Область доступа для Gmail API - используем менее строгий scope
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

class EmailSender:
    def __init__(self):
        self.service = None
        self.templates = [
            """Здравствуйте, {company_name}!""",

            """Привет, {company_name}!"""
        ]

    def authenticate_gmail(self):
        """Аутентификация через Google OAuth"""
        creds = None
        
        # Проверяем наличие файла credentials.json
        if not os.path.exists('credentials.json'):
            print("❌ Файл credentials.json не найден!")
            print("📋 Инструкция:")
            print("1. Зайдите на https://console.cloud.google.com/")
            print("2. Создайте проект и включите Gmail API")
            print("3. Создайте OAuth 2.0 Client ID для Desktop application")
            print("4. Скачайте JSON файл как 'credentials.json'")
            return False
        
        # Файл token.json хранит токены доступа пользователя
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
        # Если нет действительных учетных данных, позволить пользователю войти в систему
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                print("🔄 Обновление токена...")
                creds.refresh(Request())
            else:
                print("\n🔐 АВТОРИЗАЦИЯ GOOGLE:")
                print("=" * 40)
                print("✅ Сейчас откроется браузер")
                print("✅ Войдите в ваш Google аккаунт")
                print("✅ При предупреждении 'Приложение не проверено':")
                print("   → Нажмите 'Дополнительно'")
                print("   → Выберите 'Перейти в testsites (небезопасно)'")
                print("✅ Разрешите доступ к Gmail")
                print("=" * 40)
                input("Нажмите Enter для продолжения...")
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0, open_browser=True)
            
            # Сохранить учетные данные для следующего запуска
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            self.service = build('gmail', 'v1', credentials=creds)
            print("✅ Успешная авторизация в Gmail!")
            return True
        except Exception as error:
            print(f"❌ Ошибка подключения к Gmail: {error}")
            return False

    def create_message(self, to_email, subject, message_text):
        """Создание сообщения для отправки"""
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        
        msg = MIMEText(message_text, 'plain', 'utf-8')
        message.attach(msg)
        
        # Кодирование сообщения в base64
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        return {'raw': raw_message}

    def send_message(self, message):
        """Отправка сообщения через Gmail API"""
        try:
            sent_message = self.service.users().messages().send(
                userId="me", body=message).execute()
            return True, sent_message["id"]
        except Exception as error:
            return False, str(error)

    def read_excel_data(self, excel_file_path):
        """Чтение данных из Excel файла"""
        try:
            if not os.path.exists(excel_file_path):
                print(f"❌ Файл {excel_file_path} не найден!")
                return None
            
            # Сначала пробуем открыть "Лист1"
            try:
                df = pd.read_excel(excel_file_path, sheet_name='Лист1')
                print(f"📊 Загружено {len(df)} строк из листа 'Лист1'")
            except:
                # Если "Лист1" не найден, берем первый лист
                df = pd.read_excel(excel_file_path, sheet_name=0)
                print(f"📊 Загружено {len(df)} строк из первого листа")
            
            return df
        except Exception as error:
            print(f'❌ Ошибка при чтении Excel файла: {error}')
            return None

    def preview_data(self, df, company_col, email_col):
        """Предварительный просмотр данных"""
        print("\n📋 Предварительный просмотр данных:")
        print("-" * 50)
        
        valid_rows = 0
        for index, row in df.iterrows():
            company_name = row.get(company_col, '')
            email = row.get(email_col, '')
            
            if pd.notna(company_name) and pd.notna(email) and company_name and email:
                valid_rows += 1
                if valid_rows <= 5:  # Показываем только первые 5 валидных строк
                    print(f"{valid_rows}. {company_name} → {email}")
        
        if valid_rows > 5:
            print(f"... и еще {valid_rows - 5} записей")
        
        print(f"\n✅ Всего валидных записей для отправки: {valid_rows}")
        print("-" * 50)
        
        return valid_rows

    def send_bulk_emails(self, excel_file_path, subject="Предложение сотрудничества", delay=2):
        """Массовая отправка писем"""
        print("🚀 Запуск программы отправки писем")
        print("=" * 50)
        
        # Аутентификация
        if not self.authenticate_gmail():
            return
        
        # Чтение данных из Excel
        df = self.read_excel_data(excel_file_path)
        if df is None:
            return
        
        # Проверяем наличие нужных столбцов
        # Ищем столбцы с названиями компаний и email
        company_col = None
        email_col = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'компан' in col_lower or 'название' in col_lower or col == 'C':
                company_col = col
            elif 'email' in col_lower or 'почт' in col_lower or 'mail' in col_lower or col == 'D':
                email_col = col
        
        if company_col is None:
            print("❌ Ошибка: Не найден столбец с названиями компаний")
            print(f"Ищу столбцы содержащие: 'компан', 'название' или 'C'")
            print(f"Доступные столбцы: {list(df.columns)}")
            return
        
        if email_col is None:
            print("❌ Ошибка: Не найден столбец с email адресами")
            print(f"Ищу столбцы содержащие: 'email', 'почт', 'mail' или 'D'")
            print(f"Доступные столбцы: {list(df.columns)}")
            return
        
        print(f"✅ Используемые столбцы:")
        print(f"Компании: '{company_col}'")
        print(f"Email: '{email_col}'")
        
        # Предварительный просмотр
        valid_count = self.preview_data(df, company_col, email_col)
        if valid_count == 0:
            print("❌ Нет валидных данных для отправки")
            return
        
        # Подтверждение отправки
        print(f"\n📧 Тема письма: '{subject}'")
        print(f"⏱️ Задержка между отправками: {delay} сек")
        
        confirm = input("\n❓ Продолжить отправку? (y/N): ").lower()
        if confirm not in ['y', 'yes', 'да', 'д']:
            print("❌ Отправка отменена")
            return
        
        sent_count = 0
        failed_count = 0
        
        print("\n📤 Начинаем отправку...")
        print("-" * 50)
        
        # Отправка писем
        for index, row in df.iterrows():
            company_name = row.get(company_col, '')
            email = row.get(email_col, '')
            
            # Пропускаем строки с пустыми данными
            if pd.isna(company_name) or pd.isna(email) or not company_name or not email:
                continue
            
            # Выбираем случайный шаблон
            template = random.choice(self.templates)
            
            # Подставляем название компании
            message_text = template.format(company_name=company_name)
            
            # Создаем и отправляем сообщение
            message = self.create_message(email, subject, message_text)
            
            success, result = self.send_message(message)
            
            if success:
                sent_count += 1
                template_num = self.templates.index(template) + 1
                print(f"✅ {sent_count}. {company_name} ({email}) - Шаблон {template_num}")
            else:
                failed_count += 1
                print(f"❌ {company_name} ({email}) - Ошибка: {result}")
            
            # Задержка между отправками
            if index < len(df) - 1:  # Не ждем после последнего письма
                time.sleep(delay)
        
        print("\n" + "=" * 50)
        print("📈 ИТОГОВЫЙ ОТЧЕТ:")
        print(f"✅ Успешно отправлено: {sent_count}")
        print(f"❌ Ошибок: {failed_count}")
        print(f"📊 Всего обработано: {sent_count + failed_count}")
        print("=" * 50)

def main():
    print("📧 Email Sender - Массовая отправка писем")
    print("=" * 50)
    
    sender = EmailSender()
    
    # Укажите путь к вашему Excel файлу
    excel_file = input("📂 Введите путь к Excel файлу: ").strip().strip('"')
    
    if not excel_file:
        print("❌ Файл не выбран")
        return
    
    # Стандартные настройки
    subject = "Предложение сотрудничества"
    delay = 3  # Увеличиваем задержку для безопасности
    
    print(f"\n📧 Тема письма: {subject}")
    print(f"⏱️ Задержка между отправками: {delay} сек")
    print("📝 Используются 2 случайных шаблона писем")
    
    # Отправка писем
    sender.send_bulk_emails(excel_file, subject, delay)
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
