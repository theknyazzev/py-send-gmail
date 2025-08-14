# 📧 Advanced Email Sender

A powerful Python application for automated email campaigns using Gmail API with Excel integration and smart templates.

![Python](https://img.shields.io/badge/python-v3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Gmail API](https://img.shields.io/badge/Gmail-API-red.svg)

## ✨ Features

- 🔐 **Secure Gmail API Integration** - OAuth2 authentication for safe email sending
- 📊 **Excel/CSV Support** - Load recipient data from spreadsheet files
- 📝 **Smart Templates** - Multiple customizable email templates with variable substitution
- ⏱️ **Rate Limiting** - Built-in delays to respect Gmail API limits
- 🎯 **Targeted Campaigns** - Personalized emails with company name insertion
- 📈 **Progress Tracking** - Real-time sending progress with detailed logging
- 🛡️ **Error Handling** - Comprehensive error handling and retry mechanisms
- 🎨 **User-Friendly Interface** - Interactive command-line interface

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Google Cloud Console account
- Gmail account

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/advanced-email-sender.git
   cd advanced-email-sender
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Gmail API credentials**
   - Visit [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one
   - Enable Gmail API
   - Create OAuth 2.0 credentials (Desktop Application)
   - Download `credentials.json` and place it in the project root

### Usage

1. **Prepare your data**
   - Create an Excel file with columns: `email` and `company_name`
   - Example:
     ```
     email                    | company_name
     contact@example.com      | Example Corp
     info@business.com        | Business Inc
     ```

2. **Run the application**
   ```bash
   python email_sender_advanced.py
   ```

3. **Follow the interactive prompts**
   - Select your Excel/CSV file
   - Choose email template
   - Configure sending parameters
   - Start the campaign

## 📋 Excel File Format

Your Excel file should contain the following columns:

| Column | Description | Required |
|--------|-------------|----------|
| `email` | Recipient email address | ✅ Yes |
| `company_name` | Company name for personalization | ✅ Yes |

**Example Excel structure:**
```
Row 1: email, company_name
Row 2: john@techcorp.com, TechCorp Solutions
Row 3: mary@innovate.io, Innovate Labs
```

## 🎨 Email Templates

The application includes several professional email templates:

### Template 1: Follow-up Style
```
Hello {company_name}!

I understand you're busy, and my message might have gotten lost. 
I wanted to follow up: is my proposal still relevant to you?

To help you evaluate my approach, I'm ready to prepare a test version 
of your homepage design. This will take a couple of days, and you can 
review the result without any obligations.

If you like it — we'll work together. If not — you don't lose anything.

What do you think? Thank you for your response!
```

### Template 2: Brief Approach
```
Hello {company_name}!

I don't want to bother you, but I wanted to clarify: is there interest 
in my proposal? I understand you're busy, so I'll be brief.

I can quickly create a trial design for your homepage — this way you 
can evaluate the quality without unnecessary discussions. No obligations, 
just results.

If it works — great! If not — just let me know.

How does this option sound to you? Thank you!
```

## ⚙️ Configuration

### Rate Limiting
- Default delay: 5-10 seconds between emails
- Customizable through interactive interface
- Respects Gmail API quotas

### Authentication
- Secure OAuth2 flow
- Automatic token refresh
- Credentials stored locally

## 🔧 Advanced Usage

### Custom Templates
You can modify templates in the `EmailSender` class:

```python
self.templates = [
    "Your custom template with {company_name} placeholder",
    "Another template for {company_name}"
]
```

### Batch Processing
The application supports processing large lists with automatic batching and progress tracking.

## 📊 Monitoring & Logging

- Real-time progress display
- Detailed error logging
- Success/failure statistics
- Automatic retry on temporary failures

## 🛠️ Troubleshooting

### Common Issues

**"credentials.json not found"**
- Download OAuth2 credentials from Google Cloud Console
- Place the file in the project root directory

**"Access denied" errors**
- Ensure Gmail API is enabled in Google Cloud Console
- Check OAuth consent screen configuration
- Verify application type is set to "Desktop Application"

**Rate limiting errors**
- Increase delay between emails
- Reduce batch size
- Check Gmail API quotas

### Debug Mode
Run with debug information:
```bash
python email_sender_advanced.py --debug
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup
See [DEVELOPMENT.md](DEVELOPMENT.md) for detailed development instructions.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Google Gmail API for secure email sending
- pandas for Excel/CSV processing
- The Python community for excellent libraries

## 📞 Support

If you encounter any issues or have questions:

1. Check the [Issues](https://github.com/yourusername/advanced-email-sender/issues) page
2. Create a new issue with detailed description
3. Include error messages and system information

---

**⚠️ Disclaimer:** This tool is for legitimate email marketing purposes only. Please ensure you comply with applicable laws and regulations regarding email marketing in your jurisdiction. Always obtain proper consent before sending marketing emails.
