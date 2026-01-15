# 📊 Vertify Mapping Spreadsheet Generator

Web application to convert Vertify mapping JSON files into formatted Excel spreadsheets.

## 🚀 Features

- ✅ **Automatic generation** - Upload JSON and download Excel automatically
- ✅ **No installation required** - Web-based interface
- ✅ **Visual preview** - Preview ObjectMaps before download
- ✅ **Formatted output** - Professional Excel spreadsheet with multiple tabs
- ✅ **Free hosting** - Deploy on Streamlit Cloud at no cost

## 🏗️ Project Structure

```
Vertify/
├── src/
│   ├── app.py          # Streamlit web interface
│   ├── generator.py    # Excel generation logic
│   ├── styles.py       # Excel styling and formatting
│   └── __init__.py     # Python module initialization
├── requirements.txt    # Python dependencies
├── .gitignore         # Git ignore configuration
└── README.md          # This file
```

## 🛠️ Technologies

- **[Streamlit](https://streamlit.io/)** - Web framework
- **[OpenPyXL](https://openpyxl.readthedocs.io/)** - Excel manipulation
- **Python 3.9+**

## 💻 Local Development

### Prerequisites

- Python 3.9 or higher
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/your-username/vertify-mapping-generator.git
cd vertify-mapping-generator

# Install dependencies
pip install -r requirements.txt
```

### Run Locally

```bash
streamlit run src/app.py
```

The app will open automatically at `http://localhost:8501`

## 🚀 Deploy on Streamlit Cloud

### Step by step:

1. **Create a GitHub repository**
   - Go to [github.com/new](https://github.com/new)
   - Name the repository (e.g., `vertify-mapping-generator`)
   - Create the repository

2. **Push the code**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/your-username/vertify-mapping-generator.git
   git push -u origin main
   ```

3. **Deploy on Streamlit Cloud**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click "New app"
   - Select:
     - **Repository**: `your-username/vertify-mapping-generator`
     - **Branch**: `main`
     - **Main file path**: `src/app.py`
   - Click "Deploy!"

4. **Wait** ~2-3 minutes and your app will be live! 🎉

## 📖 How to Use

1. Access the web application
2. Upload the Vertify mapping JSON file
3. Review the displayed information
4. The Excel spreadsheet is generated automatically
5. Download the XLSX file

## 📊 Generated Spreadsheet Structure

The generated Excel spreadsheet contains:

- **Tab 1**: `Movements to migrate` - Summary of all ObjectMaps
- **Tabs 2-N**: Details of each ObjectMap including:
  - API Request configuration
  - Merge rules
  - Filter conditions
  - Field mappings (Properties Map)

## 🎯 Modular Architecture

The project follows a clean, modular architecture:

- **`src/app.py`** - User interface (Streamlit)
- **`src/generator.py`** - Business logic (Excel generation)
- **`src/styles.py`** - Formatting and styling

This separation ensures:
- ✅ Easy maintenance
- ✅ Testable components
- ✅ Reusable code
- ✅ Clear responsibilities

## 🤝 Contributing

Contributions are welcome! Feel free to:

1. Fork the project
2. Create a feature branch (`git checkout -b feature/MyFeature`)
3. Commit your changes (`git commit -m 'Add MyFeature'`)
4. Push to the branch (`git push origin feature/MyFeature`)
5. Open a Pull Request

## 📄 License

This project is under the MIT License.

---

**Made with ❤️ by Digibee**
- Configurações (Sandbox, Credentials, etc.)
- Notes

### **Abas 2-N: Detalhes de cada ObjectMap**
Cada aba contém:
1. **API Request** - Informações dos sistemas
2. **Merge** - Regras de merge
3. **Filter** - Filtros aplicados
4. **Field Mapping** - Mapeamento completo de campos

---

## 🚀 Recomendação

Use a **versão 2.0** para novos projetos - ela é mais flexível e profissional!
