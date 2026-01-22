# 🛠️ SuiteTools - Spreadsheet, Document & Slides Editor

A developer-friendly, browser-based tool to create, view, edit, and query spreadsheets, documents, and presentations.

## Features

### 📊 Spreadsheet Features
- ✅ **Create New Spreadsheets** - Start from scratch and build your own
- ✅ **SQL Queries** - Filter and analyze data using SQL syntax
- ✅ **Edit Cells** - Double-click to edit any cell
- ✅ **Add/Remove Rows & Columns** - Dynamically modify your data
- ✅ **Rename Columns** - Double-click headers to rename
- ✅ **Multi-Sheet Support** - View and switch between Excel sheets
- ✅ **CSV Support** - Handles CSV with various delimiters
- ✅ **Excel Support** - Reads and writes .xlsx and .xls files
- ✅ **Export Options** - Convert between CSV, Excel, JSON, HTML

### 📄 Document Features
- ✅ **Create New Documents** - Start with a blank document
- ✅ **Rich Text Editing** - Bold, italic, underline, strikethrough
- ✅ **Font Styling** - Change font family, size, and color
- ✅ **Text Alignment** - Left, center, right, justify
- ✅ **Lists** - Bullet and numbered lists
- ✅ **Links & Images** - Insert hyperlinks and images
- ✅ **Open DOCX Files** - View and edit Word documents
- ✅ **Export Options** - Save as DOCX, HTML, or plain text

### 📽️ Slides/Presentation Features
- ✅ **Create New Presentations** - Start with a blank slide
- ✅ **Import PPTX Files** - Open and view PowerPoint presentations
- ✅ **Import PPT Files** - Basic support for legacy PowerPoint format
- ✅ **Edit Slides** - Click on text to edit titles and content
- ✅ **Slide Styling** - Draggable panel with text color, background colors, gradients, and image upload
- ✅ **Toggle Layouts** - Switch between title and content layouts
- ✅ **Slide Backgrounds** - Preserves colors, gradients, and background images
- ✅ **Slide Navigation** - Use arrow keys or buttons to navigate
- ✅ **Thumbnail Sidebar** - Quick slide overview and selection
- ✅ **Add/Delete/Duplicate Slides** - Full slide management
- ✅ **Move Slides** - Reorder slides with up/down buttons
- ✅ **Title & Content Layouts** - Support for different slide layouts
- ✅ **Embedded Images** - Shows images within slides
- ✅ **Export to PPTX** - Save presentations as PowerPoint files
- ✅ **Export to PDF** - High-quality PDF with preserved backgrounds and styling
- ✅ **Keyboard Shortcuts** - Ctrl+S save, Ctrl+D duplicate, arrows navigate
- ✅ **Fullscreen Slideshow Mode** - Present slides in fullscreen with:
  - ▶️ Click "Slideshow" button or press F5 to start
  - ⬅️ ➡️ Arrow keys, Space, PageUp/Down for navigation
  - 🏠 Home/End to jump to first/last slide
  - ⏯️ Auto-play mode with 5-second intervals (press P to toggle)
  - 📊 Progress bar showing presentation progress
  - 🔢 Slide counter display
  - 🖱️ Click navigation areas on sides of screen
  - ⎋ Press Escape or click ✕ to exit

### 🎨 General Features
- ✅ **Dark/Light Mode** - Toggle between themes (auto-saves preference)
- ✅ **Drag & Drop** - Simply drag files onto the page
- ✅ **Keyboard Shortcuts** - Ctrl+S to save, Ctrl+B/I/U for formatting
- ✅ **No Upload Required** - Files are processed locally in your browser
- ✅ **Auto-Save** - Enable auto-save for spreadsheets (when opened via Browse)
- ✅ **Home Navigation** - Click the SuiteTools logo to return home (with save prompt)

## SQL Query Examples

### SELECT Queries
```sql
-- Select all data from current sheet
SELECT * FROM data

-- Each sheet is also a table (e.g., Sheet1, Employees, Sales)
SELECT * FROM Employees

-- Filter rows
SELECT * FROM data WHERE Status = 'Active'

-- Search text (LIKE)
SELECT * FROM data WHERE Name LIKE '%john%'

-- Sort data
SELECT * FROM data ORDER BY Salary DESC

-- Aggregations
SELECT Department, COUNT(*) as Count, AVG(Salary) as AvgSalary 
FROM data GROUP BY Department
```

### UPDATE, INSERT, DELETE Queries
```sql
-- Update rows
UPDATE data SET Status = 'Inactive' WHERE Department = 'Sales'

-- Update multiple columns
UPDATE Employees SET Salary = Salary * 1.1, Status = 'Reviewed' 
WHERE Department = 'Engineering'

-- Insert new row
INSERT INTO data (Name, Email, Department) 
VALUES ('John Doe', 'john@example.com', 'Marketing')

-- Delete rows
DELETE FROM data WHERE Status = 'Inactive'

-- Delete with multiple conditions  
DELETE FROM Employees WHERE Department = 'Temp' AND Salary < 30000
```

**Note:** UPDATE, INSERT, DELETE queries modify your data directly. Changes can be saved with `Ctrl+S` or auto-save.

### JOIN Queries (Multi-Sheet)
```sql
-- Inner Join between two sheets
SELECT e.*, d.Department_Name 
FROM Employees e 
JOIN Departments d ON e.Dept_ID = d.ID

-- Left Join
SELECT * FROM Orders o 
LEFT JOIN Customers c ON o.Customer_ID = c.ID

-- Cross Join with filter
SELECT a.*, b.* FROM Sheet1 a, Sheet2 b 
WHERE a.ID = b.Reference_ID

-- Union (combine rows from multiple sheets)
SELECT Name, Email FROM Employees
UNION
SELECT Name, Email FROM Contractors
```

### Table Naming
- `data` - Always refers to the **currently selected sheet**
- Each sheet becomes a table using its name (spaces → underscores)
  - "Employee List" → `Employee_List`
  - "Q1 Sales" → `Q1_Sales`

## Quick Start

### Option 1: Python (Recommended)
```bash
python server.py              # Default port 8080
python server.py 3000         # Custom port 3000
python server.py -p 5000      # Using -p flag
python server.py --port 9000  # Using --port flag
python server.py --no-browser # Don't auto-open browser
```

### Option 2: Python 3 http.server
```bash
python -m http.server 8080
```

### Option 3: Node.js
```bash
npx serve -p 8080
```

Then open your browser to: **http://localhost:PORT**

## Project Structure

```
├── index.html          # Main HTML page
├── styles.css          # Styling
├── js/                 # Modular JavaScript files
│   ├── state.js        # Global state management
│   ├── utils.js        # Utility functions
│   ├── theme.js        # Dark/light theme handling
│   ├── dom.js          # DOM element references
│   ├── filter.js       # Quick filter functionality
│   ├── sql.js          # SQL query engine
│   ├── spreadsheet.js  # Spreadsheet operations
│   ├── document.js     # Document editor
│   ├── slides.js       # Presentation/slides editor
│   ├── file-handler.js # File import/export
│   └── main.js         # App initialization & event handlers
├── server.py           # Simple Python server
├── sample.csv          # Sample data for testing
└── README.md           # This file
```

### Module Descriptions

| Module | Description |
|--------|-------------|
| `state.js` | Centralized state management for the entire app |
| `utils.js` | Helper functions (formatting, parsing, downloads) |
| `theme.js` | Dark/light mode toggle functionality |
| `dom.js` | Cached DOM element references |
| `filter.js` | Quick filter panel and search functionality |
| `sql.js` | SQL query execution and table management |
| `spreadsheet.js` | Table rendering, cell editing, row/column operations |
| `document.js` | Rich text editor for documents |
| `slides.js` | Presentation viewer and editor with PPTX support |
| `file-handler.js` | File parsing (CSV, Excel, DOCX, PPTX) and export |
| `main.js` | Event listeners and app initialization |

## How It Works

### Spreadsheets
1. Open the app in your browser
2. Click **"✨ Create New → New Spreadsheet"** or drag & drop a CSV/Excel file
3. Add columns and rows using the toolbar buttons
4. Double-click cells to edit, double-click headers to rename
5. Use SQL queries to filter and analyze data
6. Export in your preferred format

### Documents
1. Click **"✨ Create New → New Document"** or drag & drop a DOCX file
2. Start typing - it's a full rich text editor
3. Use the toolbar to format text (bold, italic, fonts, colors, etc.)
4. Insert links and images as needed
5. Save as DOCX, HTML, or plain text

### Presentations/Slides
1. Click **"✨ Create New → New Presentation"** or drag & drop a PPT/PPTX file
2. **Click on any text** to edit titles, subtitles, and content
3. Use the **🎨 button** (top-right of slide) to change background color
4. Use the **layout button** to toggle between title and content layouts
5. Add, duplicate, delete, or reorder slides using the toolbar
6. Use arrow keys (← →) or navigation buttons to move between slides
7. Press **Ctrl+S** to save, **Ctrl+D** to duplicate current slide
8. Click **▶️ Slideshow** or press **F5** to present in fullscreen mode
9. Export as PPTX PowerPoint file

## Browser Compatibility

Works in all modern browsers:
- Chrome, Firefox, Safari, Edge

## Dependencies

All dependencies are loaded from CDN:
- [SheetJS](https://sheetjs.com/) - For Excel file parsing
- [AlaSQL](https://alasql.org/) - For SQL query support
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js) - For reading DOCX files
- [docx](https://docx.js.org/) - For creating DOCX files
- [JSZip](https://stuk.github.io/jszip/) - For reading PPTX files
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) - For creating PPTX files
- [html2canvas](https://html2canvas.hertzen.com/) - For rendering slides to canvas
- [jsPDF](https://github.com/parallax/jsPDF) - For generating PDF files

## License

MIT License - Feel free to use and modify!