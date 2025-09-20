# Bulk Text Replacement for Word
🎯 Easy-to-use tool for replacing text in Word documents. Handles hyperlinks, shapes, headers, footers safely. Previews changes and processes multiple files at once. Perfect for bulk document updates. Download exe - no installation needed! Made for efficiency ⚡

## 🚀 Quick start

### Option 1: Download executable
1. Go to [**exe**](exe) and download `WordTextReplacer.exe`.
2. Double-click to run. **no installation required!**
3. Add your Word documents and start replacing text.

### Option 2: Download executable and add it to your context menu (recommended for the fastest approach)
1. Go to [**exe**](exe) and download `WordTextReplacer.exe`.
2. Save it to your preferred path. **no installation required!**
3. Go to [**c_m_e**](c_m_e) and download `add_word_text_replacer_to_your_context_menu.reg`.
4. Adapt the .reg file by inserting the path where you saved `WordTextReplacer.exe` (step 2) in line 8 and 15.
   (Optional: Edit the icon in line 4 & 11 and the text of the context menu line 5 & 12.)
6. Run the .reg file and confirm that you want to add it to registry editor. Afterwards, you can delete the .reg file.
7. Right-click on any .doc/.docx/.docm file and click on your context menu entry to run the script.
   (Optional: Add more files in which you want to run the search and replace workflow.)
9. Start replacing text.

### Option 3: You're a developer? Use the code and have fun! 
It's not much code, just download the three scripts from [**src**](src) and play around with it. :)

## ✨ Features and best practices

| Feature | Standard Replace | Advanced Replace |
|---------|------------------|------------------|
| **Speed** | ⚡ Very fast | 🔧 Comprehensive |
| **Main content** | ✅ Paragraphs & tables | ✅ Paragraphs & tables |
| **Shapes & text boxes** | ❌ | ✅ |
| **Headers & footers** | ❌ | ✅ |
| **Footnotes & endnotes** | ❌ | ✅ |
| **Form fields** | ❌ | ✅ |
| **Formatting preserved?** | ✅ | ✅ |
| **Hyperlinks** | ⚠️ Replaces text, but removes link. | ✅ Replaces text and preserves links. |

#### When to use Standard Replace
Large batch operations (100+ files)/Simple text in paragraphs and tables/No hyperlinks in replacement text/Speed is important/Basic formatting preservation
- **Remember:** If your file has "special content" (shapes, text boxes, headers, footers, footnotes, endnotes, form fields, and hyperlinks), but you are not going to replace any of this text, you can still use the very fast Standard Replace. It won't break any of your "special content".

#### When to use Advanced Replace
If you are going to replace "special content", use Advanced Replace as it will work for both standard content, such as paragraphs and tables, and "special content". It is just slower than the Standard Replace

#### Why does the tool offer Standard Replace if Advanced Replace covers everything?
Sometimes you just want to replace a text snippet which is not present in any "special content" and you want to do that as fast as possible. Therefore, I added the Standard Replace (which will maintain formatting, don't worry!).

### Additional features
- 👀 **Preview**: See changes before running the search and replace workflow.
- 🔗 **Hyperlink check**: Analyze documents for hyperlinks.
- 📝 **Non-breaking space support**: Use `_nbsp_` or Shift+Space.
- 📁 **Batch processing**: Handle multiple documents at once.
- 💾 **Backup creation**: Optional backup files.
- ⌨️ **Keyboard shortcuts**: Efficient workflow.

## The tool

### 📸 Screenshot

<img width="526" height="556" alt="grafik" src="https://github.com/user-attachments/assets/52924273-4293-4d2b-aabb-405051283cb6" />

### ▶️ YouTube

This video, with a slightly outdated UI, shows the [tool in action](https://www.youtube.com/watch?v=ff8-k-COUYc).

If you don't want to hear me talking for too long, just jump to [3:30](https://youtu.be/ff8-k-COUYc?si=H4sXY8syJmcAigQc&t=210).

### 🎯 Use cases (examples)

- **Legal documents**: Update contract terms across multiple files.
- **Technical documentation**: Replace product names or versions.
- **Marketing materials**: Update company information or branding.
- **Academic papers**: Standardize terminology or citations.
- **Corporate communications**: Update contact information or policies.

### ⌨️ Keyboard shortcuts (excerpt)

| Shortcut | Action |
|----------|--------|
| `F5` | Preview changes |
| `Ctrl+Enter` | Standard Replace |
| `Shift+Enter` | Advanced Replace |
| `Ctrl+O` | Add more files |
| `Shift+Space` | Insert non-breaking space |
| `Ctrl+1/2` | Focus search/replace boxes |
| `Tab` | Switch between search & replace boxes |

### 📋 System requirements

- **OS**: Windows 7/8/10/11
- **Software**: Microsoft Word (for Advanced Replace features)
- **Installation**: None required for executable version

## Project structure
```
WordTextReplacer/
├── c_m_e/                     # .reg file
├── doc/                       # User Manual
├── exe/                       # .exe file
└── src/                       # Source Code (one .py & 2 .vbs)
```

## 🤝 Contributing

Contributions are welcome! I am not an experienced developer. If you have ideas/upgrades/improvements, I am happy to see your code!

1. Fork the repository.
2. Create a feature branch.
3. Make your changes.
4. Submit a pull request.

## 📄 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) for details.

## 🫶 Support

The tool is free and I am happy with sharing it with whoever wants to use it. If you want, you can [buy me a coffee](https://buymeacoffee.com/abbatem). ☕🙏 And connect with [me](https://www.linkedin.com/in/mario-abbate-601885150/)!

## 🏆 Acknowledgments

- Built with Python and tkinter for the GUI.
- VBScript integration for advanced Word automation.
- python-docx for fast document processing.
- I used AI to create this tool. :)

---

⬛🟦⬛ **Forza Inter!** ⬛🟦⬛ 

*Made with enthusiasm for efficient document processing.*

