# 🛠️ Development Guide

This document provides detailed instructions for developers who want to contribute to or modify the Advanced Email Sender project.

## 📋 Prerequisites

### System Requirements
- **Python 3.8+** (recommended: Python 3.9 or 3.10)
- **Git** for version control
- **Google Cloud Console** access for API credentials
- **Code Editor** (VS Code, PyCharm, etc.)

### Knowledge Requirements
- Python programming fundamentals
- Basic understanding of OAuth2 authentication
- Familiarity with Gmail API
- Excel/CSV data manipulation with pandas

## 🏗️ Project Structure

```
advanced-email-sender/
├── email_sender_advanced.py    # Main application file
├── requirements.txt             # Python dependencies
├── credentials.json            # Gmail API credentials (not in repo)
├── token.json                  # OAuth token cache (auto-generated)
├── README.md                   # Project documentation
├── DEVELOPMENT.md              # This file
├── LICENSE                     # MIT license
└── .gitignore                  # Git ignore rules
```

## 🚀 Setup Development Environment

### 1. Clone and Setup
```bash
# Clone the repository
git clone https://github.com/theknyazzev/py-send-gmail.git
cd advanced-email-sender

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Install development dependencies
pip install pytest black flake8 mypy
```

### 2. Configure Gmail API

#### Google Cloud Console Setup
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Gmail API:
   ```
   APIs & Services → Library → Search "Gmail API" → Enable
   ```
4. Create OAuth 2.0 credentials:
   ```
   APIs & Services → Credentials → Create Credentials → OAuth client ID
   Application type: Desktop Application
   Name: Advanced Email Sender
   ```
5. Download the JSON file and rename to `credentials.json`
6. Place in project root directory

#### OAuth Consent Screen
Configure the OAuth consent screen:
- **Application name**: Advanced Email Sender
- **User support email**: Your email
- **Scopes**: Add Gmail compose scope
- **Test users**: Add your Gmail address

### 3. Environment Variables (Optional)
Create `.env` file for development:
```bash
DEBUG=True
LOG_LEVEL=DEBUG
DEFAULT_DELAY=5
MAX_RETRIES=3
```

## 🧪 Testing

### Unit Tests
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=email_sender_advanced

# Run specific test file
pytest tests/test_authentication.py
```

### Manual Testing
1. **Create test data**:
   ```bash
   python create_test_data.py
   ```

2. **Test authentication**:
   ```bash
   python -c "from email_sender_advanced import EmailSender; sender = EmailSender(); sender.authenticate_gmail()"
   ```

3. **Test email sending** (use your own email):
   ```bash
   python test_send_single.py
   ```

## 🔧 Code Style and Standards

### Code Formatting
```bash
# Format code with Black
black email_sender_advanced.py

# Check with flake8
flake8 email_sender_advanced.py

# Type checking with mypy
mypy email_sender_advanced.py
```

### Coding Standards
- **PEP 8** compliance
- **Type hints** for all functions
- **Docstrings** for all classes and methods
- **Error handling** for all external API calls
- **Logging** instead of print statements

### Example Function Format
```python
def send_email(self, recipient: str, subject: str, body: str) -> bool:
    """
    Send an email to a single recipient.
    
    Args:
        recipient: Email address of the recipient
        subject: Email subject line
        body: Email body content
        
    Returns:
        bool: True if email was sent successfully, False otherwise
        
    Raises:
        GoogleAPIError: If Gmail API call fails
        ValueError: If recipient email is invalid
    """
    try:
        # Implementation here
        return True
    except Exception as e:
        self.logger.error(f"Failed to send email: {e}")
        return False
```

## 🏗️ Architecture Overview

### Class Structure
```python
class EmailSender:
    def __init__(self):
        # Initialize Gmail service, templates, logging
        
    def authenticate_gmail(self) -> bool:
        # Handle OAuth2 authentication
        
    def load_recipients(self, file_path: str) -> pd.DataFrame:
        # Load and validate recipient data
        
    def send_bulk_emails(self, recipients: pd.DataFrame) -> dict:
        # Main bulk sending logic
        
    def send_single_email(self, recipient: dict) -> bool:
        # Send individual email
        
    def get_template(self, template_id: int) -> str:
        # Retrieve and format email template
```

### Key Components

1. **Authentication Module**
   - OAuth2 flow handling
   - Token storage and refresh
   - Error handling for auth failures

2. **Data Processing Module**
   - Excel/CSV file parsing
   - Data validation
   - Error handling for malformed data

3. **Email Engine**
   - Template processing
   - Rate limiting
   - Retry logic
   - Progress tracking

4. **Logging System**
   - Structured logging
   - Error categorization
   - Performance metrics

## 🔄 API Integration

### Gmail API Endpoints Used
- **gmail.users.messages.send**: Send emails
- **gmail.users.getProfile**: Get user profile info

### Rate Limiting Strategy
```python
import time
import random

def rate_limit_delay(self):
    """Implement smart rate limiting"""
    base_delay = 5  # Base delay in seconds
    jitter = random.uniform(0, 2)  # Add randomness
    time.sleep(base_delay + jitter)
```

### Error Handling Patterns
```python
from googleapiclient.errors import HttpError

def handle_api_error(self, error: HttpError) -> bool:
    """Handle Gmail API errors gracefully"""
    if error.resp.status == 403:
        # Handle quota exceeded
        return self.handle_quota_error()
    elif error.resp.status == 401:
        # Handle authentication error
        return self.reauthenticate()
    else:
        # Handle other errors
        return self.log_and_continue(error)
```

## 🐛 Debugging

### Enable Debug Mode
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Common Debug Scenarios

1. **Authentication Issues**
   ```python
   # Check credentials file
   import os
   print(f"Credentials exist: {os.path.exists('credentials.json')}")
   
   # Verify token
   creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   print(f"Token valid: {creds.valid}")
   ```

2. **Data Loading Issues**
   ```python
   # Debug Excel loading
   df = pd.read_excel('data.xlsx')
   print(df.info())
   print(df.head())
   ```

3. **API Call Issues**
   ```python
   # Enable HTTP debugging
   import httplib2
   httplib2.debuglevel = 1
   ```

## 📦 Building and Deployment

### Create Distribution Package
```bash
# Install build tools
pip install build twine

# Build package
python -m build

# Upload to PyPI (for public packages)
twine upload dist/*
```

### Create Executable
```bash
# Install PyInstaller
pip install pyinstaller

# Create executable
pyinstaller --onefile email_sender_advanced.py
```

### Docker Container
```dockerfile
FROM python:3.9-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .
CMD ["python", "email_sender_advanced.py"]
```

## 🔒 Security Considerations

### Credential Security
- **Never commit** `credentials.json` or `token.json`
- Use environment variables for sensitive data
- Implement proper secret management in production

### API Security
- Use minimal required scopes
- Implement proper rate limiting
- Add request validation
- Log security events

### Data Privacy
- Encrypt recipient data at rest
- Implement data retention policies
- Add GDPR compliance features
- Secure temporary file handling

## 🚀 Performance Optimization

### Batch Processing
```python
def process_in_batches(self, recipients: list, batch_size: int = 50):
    """Process recipients in batches to optimize performance"""
    for i in range(0, len(recipients), batch_size):
        batch = recipients[i:i + batch_size]
        self.process_batch(batch)
        time.sleep(self.batch_delay)
```

### Memory Management
- Stream large Excel files
- Clear processed data from memory
- Monitor memory usage during bulk operations

### Caching Strategy
- Cache authenticated service objects
- Store processed templates
- Implement smart retry caching

## 📈 Monitoring and Metrics

### Key Metrics to Track
- **Send Rate**: Emails per minute
- **Success Rate**: Percentage of successful sends
- **Error Rate**: Categorized error frequencies
- **API Usage**: Quota consumption tracking

### Logging Best Practices
```python
import logging

# Configure structured logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sender.log'),
        logging.StreamHandler()
    ]
)
```

## 🤝 Contributing Guidelines

### Pull Request Process
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the full test suite
6. Update documentation
7. Submit pull request

### Commit Message Format
```
type(scope): short description

Detailed description if needed

Fixes #123
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

### Code Review Checklist
- [ ] Code follows style guidelines
- [ ] Tests pass and coverage is maintained
- [ ] Documentation is updated
- [ ] No security vulnerabilities introduced
- [ ] Performance impact is considered
- [ ] Backward compatibility is maintained

## 📚 Resources

### Documentation
- [Gmail API Documentation](https://developers.google.com/gmail/api)
- [Google Auth Library](https://google-auth.readthedocs.io/)
- [pandas Documentation](https://pandas.pydata.org/docs/)

### Tools
- [OAuth 2.0 Playground](https://developers.google.com/oauthplayground/)
- [Gmail API Explorer](https://developers.google.com/gmail/api/reference/rest)
- [Python Type Checker](http://mypy-lang.org/)

---

For questions about development, please create an issue or contact the maintainers.
