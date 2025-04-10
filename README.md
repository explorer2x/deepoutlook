# 📧 Outlook Email Assistant (VBA)

A VBA macro for Microsoft Outlook that helps with email drafting, translation, revision, and analysis using OpenAI's API (or DeepSeek API).

## ✨ Features

- **Extract Selected Text** from an email with non-printable character removal
- **AI-Powered Email Assistance**:
  - ✍️ **Revise Email** - Improve clarity and professionalism
  - 🌍 **Translate Email** - Convert emails to Chinese (while keeping necessary English terms)
  - 🔍 **Analyze Email** - Get insights in Markdown format
  - 📝 **Compose Email** - Generate professional emails from notes
- **Seamless Integration** - Inserts AI responses directly into your email

## ⚙️ Setup

1. **Enable Developer Mode in Outlook**:
   - File → Options → Customize Ribbon → Check "Developer" tab

2. **Add VBA Code**:
   - Open VBA Editor (Alt+F11)
   - Import this module

3. **API Configuration**:
   ```cmd
   setx OPENAI_API_KEY "your-api-key-here"
   ```
   - Supports both OpenAI and DeepSeek APIs

## 🚀 How to Use

1. **Select text** in an open email
2. Run one of these macros from the Developer tab:

| Macro               | Action                                  | Shortcut Recommendation |
|---------------------|-----------------------------------------|-------------------------|
| `translate_My_email` | Translates to Chinese                   | `Ctrl+Shift+T`          |
| `revise_My_email`    | Improves email professionalism          | `Ctrl+Shift+R`          |
| `Analyze_My_email`   | Provides email analysis in Markdown     | `Ctrl+Shift+A`          |
| `write_an_email`     | Generates professional email from notes | `Ctrl+Shift+W`          |

## 🔧 Technical Details

- **Text Processing**:
  - Automatically removes non-printable ASCII characters
  - Handles line breaks and special characters properly

- **API Integration**:
  - Uses `deepseek-chat` model by default (configurable)
  - JSON payload construction with proper escaping
  - 32-bit/64-bit compatible object creation

## 📝 Notes

1. The script will:
   - Warn if no text is selected
   - Notify if API key is missing
   - Handle quota exceeded errors

2. For best results:
   - Select complete sentences/paragraphs
   - Review AI suggestions before sending

## 🤖 Customization

To modify the AI behavior, edit the prompt templates in:
- `translate_My_email()`
- `revise_My_email()`
- `Analyze_My_email()`
- `write_an_email()`

Example prompt modification:
```vba
propt = "Review this email and suggest improvements focusing on conciseness:"
```

## ⚠️ Limitations

- Requires internet access for API calls
- Currently processes ~1000 tokens per request
- DeepSeek API endpoint is default (can switch to OpenAI)

---

**📌 Pro Tip**: Assign these macros to Quick Access Toolbar buttons for one-click access while composing emails!
