# 📊 SQL-First Excel Viewer

A developer-friendly, browser-based tool to view and query CSV/Excel files using SQL.

## Features

- ✅ **SQL Queries** - Filter and analyze data using SQL syntax
- ✅ **Dark/Light Mode** - Toggle between themes (auto-saves preference)
- ✅ **Drag & Drop** - Simply drag files onto the page
- ✅ **Multi-Sheet Support** - View and switch between Excel sheets
- ✅ **CSV Support** - Handles CSV with various delimiters
- ✅ **Excel Support** - Reads .xlsx and .xls files
- ✅ **Export Options** - Convert between CSV, Excel, JSON, HTML
- ✅ **Export SQL Results** - Export query results in any format
- ✅ **Search/Filter** - Quick text search across all columns
- ✅ **No Upload Required** - Files are processed locally in your browser

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
├── index.html      # Main HTML page
├── styles.css      # Styling
├── app.js          # Application logic
├── server.py       # Simple Python server
├── sample.csv      # Sample data for testing
└── README.md       # This file
```

## How It Works

1. Open the app in your browser
2. Drag & drop a CSV/Excel file or click "Browse Files"
3. View your data in a clean, sortable table
4. Use the search box to filter rows
5. Click "Clear" to reset and load a new file

## Browser Compatibility

Works in all modern browsers:
- Chrome, Firefox, Safari, Edge

## Dependencies

- [SheetJS](https://sheetjs.com/) - For Excel file parsing (loaded from CDN)

## License

MIT License - Feel free to use and modify!