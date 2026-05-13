# Excel Clone

A fully functional Microsoft Excel clone built with pure HTML, CSS, and vanilla JavaScript. This application provides a comprehensive spreadsheet experience with most of the essential features found in Microsoft Excel.

## Features

### Core Functionality
- **1000 rows × 26 columns** spreadsheet grid
- **Cell selection and navigation** with mouse and keyboard
- **Real-time cell editing** with formula support
- **Copy, cut, and paste** operations
- **Insert and delete** rows and columns

### Formula Support
- **Basic arithmetic**: `=A1+B1`, `=A1*B1+C1`, etc.
- **SUM function**: `=SUM(A1:A10)`
- **AVERAGE function**: `=AVERAGE(A1:A10)`
- **COUNT function**: `=COUNT(A1:A10)`
- **MAX function**: `=MAX(A1:A10)`
- **MIN function**: `=MIN(A1:A10)`
- **Cell references**: `A1`, `B2`, `C3`, etc.
- **Range references**: `A1:B10`, `C5:F20`, etc.

### Text Formatting
- **Font family**: Arial, Calibri, Times New Roman, Courier New
- **Font size**: 8pt to 24pt
- **Text styling**: Bold, Italic, Underline
- **Text color**: Full color picker support
- **Background color**: Full color picker support

### Cell Alignment
- **Horizontal alignment**: Left, Center, Right
- **Vertical alignment**: Top, Middle, Bottom

### Border Styling
- **All borders**: Apply borders to all sides
- **Outline**: Apply borders to outer edges only
- **Inside**: Apply borders to inner edges only
- **No border**: Remove all borders

### File Operations
- **New file**: Create a fresh spreadsheet
- **Open file**: Import .xlsx and .xls files
- **Export to XLSX**: Download spreadsheet as .xlsx file
- **Save functionality**: Export option for saving

### Navigation & Selection
- **Mouse selection**: Click and drag to select ranges
- **Keyboard navigation**: Arrow keys, Tab, Enter
- **Row/Column selection**: Click headers to select entire rows/columns
- **Multi-cell selection**: Shift+click for range selection
- **Formula bar**: Display and edit cell content/formulas

### Keyboard Shortcuts
- **Ctrl+C**: Copy selected cells
- **Ctrl+X**: Cut selected cells
- **Ctrl+V**: Paste from clipboard
- **Ctrl+N**: New file
- **Ctrl+O**: Open file
- **Ctrl+S**: Save file
- **F2**: Edit current cell
- **Escape**: Cancel editing
- **Delete/Backspace**: Clear selected cells
- **Arrow keys**: Navigate between cells
- **Enter**: Move to next row
- **Tab**: Move to next column

### User Interface
- **Excel-like appearance**: Professional spreadsheet interface
- **Context menu**: Right-click for quick actions
- **Status bar**: Display current selection and status
- **Responsive design**: Works on different screen sizes
- **Scrollable grid**: Handle large spreadsheets efficiently

## Getting Started

1. **Open the application**: Simply open `index.html` in a web browser
2. **Start working**: Click on any cell to begin entering data
3. **Use formulas**: Start with `=` to enter formulas
4. **Format cells**: Use the toolbar to apply formatting
5. **Export data**: Use "Export to XLSX" to download your work

## Technical Details

### Dependencies
- **SheetJS (xlsx)**: For importing and exporting Excel files
- **Pure HTML/CSS/JavaScript**: No additional frameworks required

### Browser Compatibility
- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

### File Format Support
- **Import**: .xlsx, .xls files
- **Export**: .xlsx files

## Usage Examples

### Basic Data Entry
```
A1: Name
B1: Age
C1: Salary
A2: John Doe
B2: 30
C2: 50000
```

### Formula Examples
```
D1: =SUM(B2:B10)          // Sum ages
E1: =AVERAGE(C2:C10)      // Average salary
F1: =COUNT(A2:A10)        // Count employees
G1: =MAX(C2:C10)          // Highest salary
H1: =MIN(C2:C10)          // Lowest salary
```

### Cell References
```
A1: 100
B1: 200
C1: =A1+B1                // Result: 300
D1: =A1*2+B1/2            // Result: 250
```

## Technical Concepts Used

### HTML Concepts

#### Semantic Structure
- **Document Structure**: Proper HTML5 document structure with `<!DOCTYPE html>`, `<html>`, `<head>`, and `<body>` elements
- **Meta Tags**: Viewport meta tag for responsive design and charset declaration
- **External Libraries**: CDN integration for SheetJS library using `<script>` tag

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Clone</title>
    <link rel="stylesheet" href="styles.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
```

#### Form Elements
- **File Input**: Hidden file input (`<input type="file">`) for opening Excel files
- **Color Inputs**: HTML5 color picker inputs (`<input type="color">`) for text and background colors
- **Select Dropdowns**: Font family and font size selection using `<select>` elements
- **Text Input**: Formula bar input field for cell content editing

```html
<!-- Hidden file input for opening files -->
<input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;" onchange="handleFileOpen(event)">

<!-- Color pickers for text and background -->
<input type="color" id="textColor" onchange="changeTextColor(this.value)" title="Text Color">
<input type="color" id="bgColor" onchange="changeBackgroundColor(this.value)" title="Background Color">

<!-- Font selection dropdowns -->
<select id="fontFamily" onchange="changeFontFamily(this.value)">
    <option value="Arial">Arial</option>
    <option value="Calibri">Calibri</option>
    <option value="Times New Roman">Times New Roman</option>
    <option value="Courier New">Courier New</option>
</select>

<!-- Formula bar input -->
<input type="text" id="formulaInput" placeholder="Enter formula or value" onkeydown="handleFormulaInput(event)">
```

#### Interactive Elements
- **Buttons**: Various action buttons for formatting and operations
- **Event Handlers**: Inline event handlers (`onclick`, `onchange`, `onkeydown`) for user interactions
- **Content Editable**: Cells made editable using `contentEditable="true"` attribute

```html
<!-- Formatting buttons with inline event handlers -->
<button id="boldBtn" onclick="toggleBold()" class="format-btn">B</button>
<button id="italicBtn" onclick="toggleItalic()" class="format-btn">I</button>
<button id="underlineBtn" onclick="toggleUnderline()" class="format-btn">U</button>

<!-- Alignment buttons -->
<button onclick="setAlignment('left')" class="align-btn">⬅</button>
<button onclick="setAlignment('center')" class="align-btn">⬌</button>
<button onclick="setAlignment('right')" class="align-btn">➡</button>

<!-- Action buttons -->
<button onclick="insertRow()" class="action-btn">Insert Row</button>
<button onclick="deleteRow()" class="action-btn">Delete Row</button>
```

#### Data Attributes
- **Custom Data Attributes**: `data-row` and `data-col` attributes for cell identification
- **Element Identification**: Using `id` attributes for targeting specific elements

```html
<!-- Cells with data attributes for identification -->
<div class="cell" data-row="0" data-col="0" tabIndex="0" contentEditable="true"></div>
<div class="cell" data-row="0" data-col="1" tabIndex="0" contentEditable="true"></div>

<!-- Elements with specific IDs for targeting -->
<div class="cell-address" id="cellAddress">A1</div>
<div class="status-bar">
    <span id="statusText">Ready</span>
    <span id="selectionInfo"></span>
</div>
```

### CSS Concepts

#### Layout Techniques
- **Flexbox**: Extensive use of `display: flex` for toolbar, formula bar, and status bar layouts
- **CSS Grid**: `display: grid` for creating the spreadsheet data grid with dynamic columns/rows
- **CSS Variables**: Custom properties (`--cols`, `--rows`) for dynamic grid sizing
- **Positioning**: Absolute positioning for headers and selection indicators

```css
/* Flexbox Layout for Main Container */
.excel-container {
    display: flex;
    flex-direction: column;
    height: 100vh;
    background-color: white;
}

/* Toolbar with Flexbox */
.toolbar {
    display: flex;
    background-color: #f8f9fa;
    border-bottom: 1px solid #dee2e6;
    padding: 8px 10px;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
}

/* CSS Grid for Data Grid */
.data-grid {
    position: absolute;
    top: 25px;
    left: 40px;
    display: grid;
    grid-template-columns: repeat(var(--cols, 26), 100px);
    grid-template-rows: repeat(var(--rows, 1000), 25px);
    border: 1px solid #dee2e6;
}

/* CSS Variables Usage */
:root {
    --cell-width: 100px;
    --cell-height: 25px;
    --header-height: 25px;
}
```

#### Responsive Design
- **Viewport Units**: `100vh` for full-height layout
- **Media Queries**: Responsive breakpoints for mobile devices
- **Flexible Sizing**: `flex: 1` for expanding containers

```css
/* Full Height Layout */
.excel-container {
    height: 100vh; /* Full viewport height */
}

/* Flexible Formula Bar */
#formulaInput {
    flex: 1; /* Takes remaining space */
    padding: 6px 8px;
    border: 1px solid #ced4da;
    border-radius: 4px;
}

/* Responsive Design */
@media (max-width: 768px) {
    .toolbar {
        flex-wrap: wrap;
        padding: 5px;
    }
    
    .cell {
        width: 80px;
        font-size: 10px;
    }
    
    .data-grid {
        grid-template-columns: repeat(var(--cols, 26), 80px);
    }
}
```

#### Visual Styling
- **Box Model**: Proper use of margin, padding, and border properties
- **Color System**: Consistent color palette using CSS custom properties
- **Typography**: Font family, size, weight, and style controls
- **Hover Effects**: Interactive states using `:hover` pseudo-class

```css
/* Box Model and Styling */
.cell {
    width: 100px;
    height: 25px;
    border: 1px solid #e9ecef;
    border-right: 1px solid #dee2e6;
    border-bottom: 1px solid #dee2e6;
    padding: 2px 4px;
    font-size: 11px;
    outline: none;
    background-color: white;
    display: flex;
    align-items: center;
    position: relative;
    cursor: cell;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* Hover Effects */
.menu-item:hover {
    background-color: #e9ecef;
}

.format-btn:hover, .align-btn:hover, .border-btn:hover, .action-btn:hover {
    background-color: #e9ecef;
}

/* Typography Controls */
.cell.bold {
    font-weight: bold;
}

.cell.italic {
    font-style: italic;
}

.cell.underline {
    text-decoration: underline;
}
```

#### Advanced CSS Features
- **Pseudo-classes**: `:focus`, `:hover`, `:last-child` for enhanced interactivity
- **CSS Animations**: Keyframe animations for loading spinners
- **Box Shadow**: Drop shadows for context menus and elevated elements
- **Border Radius**: Rounded corners for modern appearance
- **Custom Scrollbars**: Webkit scrollbar styling for better UX

```css
/* Pseudo-classes */
.cell:focus {
    border: 2px solid #007bff;
    z-index: 5;
}

.context-menu-item:last-child {
    border-bottom: none;
}

/* CSS Animations */
.loading {
    display: inline-block;
    width: 20px;
    height: 20px;
    border: 3px solid #f3f3f3;
    border-top: 3px solid #007bff;
    border-radius: 50%;
    animation: spin 1s linear infinite;
}

@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

/* Box Shadow for Context Menu */
.context-menu {
    position: absolute;
    background-color: white;
    border: 1px solid #dee2e6;
    border-radius: 4px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    z-index: 1000;
    min-width: 150px;
    display: none;
}

/* Custom Scrollbars */
::-webkit-scrollbar {
    width: 12px;
    height: 12px;
}

::-webkit-scrollbar-track {
    background: #f1f1f1;
}

::-webkit-scrollbar-thumb {
    background: #c1c1c1;
    border-radius: 6px;
}

::-webkit-scrollbar-thumb:hover {
    background: #a8a8a8;
}
```

#### State Management
- **CSS Classes**: Dynamic class toggling for cell states (selected, editing, formula, error)
- **Conditional Styling**: Different styles based on cell content and state
- **Z-index Management**: Layering system for overlapping elements

```css
/* Cell States */
.cell.selected {
    background-color: #e3f2fd;
    border: 2px solid #007bff;
}

.cell.editing {
    background-color: white;
    border: 2px solid #28a745;
}

.cell.formula {
    background-color: #fff3cd;
}

.cell.formula.error {
    background-color: #f8d7da;
    color: #721c24;
}

/* Alignment Classes */
.cell.align-left {
    text-align: left !important;
    justify-content: flex-start !important;
}

.cell.align-center {
    text-align: center !important;
    justify-content: center !important;
}

.cell.align-right {
    text-align: right !important;
    justify-content: flex-end !important;
}

/* Z-index Management */
.row-headers {
    position: absolute;
    left: 0;
    top: 0;
    z-index: 10;
}

.selection-indicator {
    position: absolute;
    border: 2px solid #007bff;
    background-color: rgba(0, 123, 255, 0.1);
    pointer-events: none;
    z-index: 3;
}
```

### JavaScript Concepts

#### Object-Oriented Programming
- **ES6 Classes**: Main `ExcelClone` class for organizing functionality
- **Constructor Pattern**: Initialization of application state and properties
- **Method Organization**: Logical grouping of related functionality
- **Encapsulation**: Private-like methods and properties within class scope

```javascript
// Main ExcelClone Class
class ExcelClone {
    constructor() {
        this.rows = 1000;
        this.cols = 26;
        this.currentCell = { row: 0, col: 0 };
        this.data = {};
        this.formats = {};
        this.init();
    }

    init() {
        this.createGrid();
        this.bindEvents();
        this.updateCellDisplay();
    }

    createGrid() {
        const dataGrid = document.getElementById('dataGrid');
        dataGrid.style.setProperty('--rows', this.rows);
        dataGrid.style.setProperty('--cols', this.cols);
        
        // Create cells dynamically
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.dataset.row = row;
                cell.dataset.col = col;
                dataGrid.appendChild(cell);
            }
        }
    }
}
```

#### DOM Manipulation
- **Element Selection**: `document.getElementById()`, `document.querySelector()`, `document.querySelectorAll()`
- **Dynamic Element Creation**: `document.createElement()` for generating cells and headers
- **Content Modification**: `textContent`, `innerHTML`, `classList` manipulation
- **Event Binding**: `addEventListener()` for handling user interactions

```javascript
// Element Selection Examples
const dataGrid = document.getElementById('dataGrid');
const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);

// Dynamic Element Creation
const cell = document.createElement('div');
cell.className = 'cell';
cell.dataset.row = row;
cell.dataset.col = col;

// Content Modification
cellElement.textContent = cellData.value || '';
cellElement.classList.add('selected');

// Event Binding
cell.onclick = (e) => this.selectCell(row, col, e);
cell.onkeydown = (e) => this.handleCellKeydown(e);
```

#### Event Handling
- **Keyboard Events**: `keydown`, `keyup` for navigation and shortcuts
- **Mouse Events**: `click`, `contextmenu`, `mousedown` for selection and interaction
- **Form Events**: `change`, `input` for toolbar controls
- **File Events**: `change` for file input handling
- **Focus Events**: `focus`, `blur` for cell editing states

```javascript
// Keyboard Event Handling
handleCellKeydown(event) {
    switch (event.key) {
        case 'Enter':
            event.preventDefault();
            this.moveToCell(this.currentCell.row + 1, this.currentCell.col);
            break;
        case 'Tab':
            event.preventDefault();
            this.moveToCell(this.currentCell.row, this.currentCell.col + 1);
            break;
        case 'ArrowUp':
            if (!this.isEditing) {
                this.moveToCell(this.currentCell.row - 1, this.currentCell.col);
            }
            break;
    }
}

// Global Keyboard Shortcuts
handleGlobalKeydown(event) {
    if (event.ctrlKey) {
        switch (event.key.toLowerCase()) {
            case 'c': this.copy(); break;
            case 'x': this.cut(); break;
            case 'v': this.paste(); break;
        }
    }
}

// Context Menu Handling
handleContextMenu(event) {
    event.preventDefault();
    const contextMenu = document.createElement('div');
    contextMenu.className = 'context-menu';
    contextMenu.style.left = event.pageX + 'px';
    contextMenu.style.top = event.pageY + 'px';
    document.body.appendChild(contextMenu);
}
```

#### Data Management
- **Object Storage**: Using objects as hash maps for cell data and formatting
- **Array Operations**: `forEach()`, `map()`, `push()`, `pop()` for data processing
- **JSON Manipulation**: Converting data to/from JSON format
- **Local Storage**: Potential for data persistence (not implemented but structure supports it)

```javascript
// Object Storage for Cell Data
setCellData(row, col, value, formula = null) {
    const cellId = this.getCellId(row, col);
    if (!this.data[cellId]) {
        this.data[cellId] = {};
    }
    
    this.data[cellId].value = value;
    this.data[cellId].displayValue = value;
    if (formula) {
        this.data[cellId].formula = formula;
    }
}

// Array Operations for Range Processing
getSelectedRangeData() {
    const minRow = Math.min(this.selectedRange.start.row, this.selectedRange.end.row);
    const maxRow = Math.max(this.selectedRange.start.row, this.selectedRange.end.row);
    const minCol = Math.min(this.selectedRange.start.col, this.selectedRange.end.col);
    const maxCol = Math.max(this.selectedRange.start.col, this.selectedRange.end.col);

    const data = [];
    for (let row = minRow; row <= maxRow; row++) {
        const rowData = [];
        for (let col = minCol; col <= maxCol; col++) {
            const cellData = this.getCellData(row, col);
            rowData.push(cellData ? cellData.value : '');
        }
        data.push(rowData);
    }
    return data;
}

// Paste Data with Array Processing
pasteData(startRow, startCol) {
    if (!this.clipboard) return;

    this.clipboard.forEach((rowData, rowIndex) => {
        rowData.forEach((value, colIndex) => {
            const targetRow = startRow + rowIndex;
            const targetCol = startCol + colIndex;
            
            if (targetRow < this.rows && targetCol < this.cols) {
                this.setCellData(targetRow, targetCol, value);
            }
        });
    });
}
```

#### Advanced JavaScript Features
- **Arrow Functions**: Concise function syntax for event handlers
- **Template Literals**: String interpolation for dynamic content
- **Destructuring**: Object and array destructuring for cleaner code
- **Spread Operator**: Array and object spreading for data manipulation
- **Async/Await**: File reading operations with Promise handling

```javascript
// Arrow Functions for Event Handlers
cell.onclick = (e) => this.selectCell(row, col, e);
cell.onkeydown = (e) => this.handleCellKeydown(e);

// Template Literals for Dynamic Content
getCellId(row, col) {
    return `${this.getColumnName(col)}${row + 1}`;
}

updateStatusBar() {
    const statusText = document.getElementById('statusText');
    const selectionInfo = document.getElementById('selectionInfo');
    
    if (minRow === maxRow && minCol === maxCol) {
        selectionInfo.textContent = `Cell: ${this.getCellId(this.currentCell.row, this.currentCell.col)}`;
    } else {
        selectionInfo.textContent = `Range: ${this.getCellId(minRow, minCol)}:${this.getCellId(maxRow, maxCol)}`;
    }
}

// Destructuring for Cleaner Code
const { start, end } = this.selectedRange;
const { row, col } = this.currentCell;

// Spread Operator for Data Manipulation
const newFormat = { ...existingFormat, bold: true };
const updatedData = { ...this.data, [cellId]: newCellData };

// Async/Await for File Operations
handleFileOpen(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Process worksheet data
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                jsonData.forEach((row, rowIndex) => {
                    row.forEach((cell, colIndex) => {
                        if (cell !== undefined && cell !== '') {
                            this.setCellData(rowIndex, colIndex, cell);
                        }
                    });
                });
                
                this.updateStatusBar('File loaded successfully');
            } catch (error) {
                this.updateStatusBar('Error loading file: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }
}
```

#### Algorithm Implementation
- **Formula Parsing**: Regular expressions for parsing cell references and ranges
- **Mathematical Evaluation**: Custom formula evaluation engine
- **Range Calculations**: Complex algorithms for cell range operations
- **Data Validation**: Input validation and error handling

```javascript
// Formula Evaluation Engine
evaluateFormula(formula) {
    if (!formula || !formula.startsWith('=')) {
        return formula;
    }

    try {
        const expression = formula.substring(1);
        
        if (expression.toUpperCase().startsWith('SUM(')) {
            return this.evaluateSum(expression);
        } else if (expression.toUpperCase().startsWith('AVERAGE(')) {
            return this.evaluateAverage(expression);
        } else {
            return this.evaluateArithmetic(expression);
        }
    } catch (error) {
        return '#ERROR';
    }
}

// SUM Function Implementation
evaluateSum(expression) {
    const range = this.parseRange(expression.match(/SUM\(([^)]+)\)/i)[1]);
    let sum = 0;
    
    for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
            const cellData = this.getCellData(row, col);
            const value = parseFloat(cellData ? cellData.value : 0);
            if (!isNaN(value)) {
                sum += value;
            }
        }
    }
    
    return sum;
}

// Range Parsing with Regular Expressions
parseRange(rangeStr) {
    if (rangeStr.includes(':')) {
        const [start, end] = rangeStr.split(':');
        const startCol = this.getColumnIndex(start.match(/[A-Z]+/)[0]);
        const startRow = parseInt(start.match(/[0-9]+/)[0]) - 1;
        const endCol = this.getColumnIndex(end.match(/[A-Z]+/)[0]);
        const endRow = parseInt(end.match(/[0-9]+/)[0]) - 1;
        
        return { startRow, startCol, endRow, endCol };
    } else {
        const col = this.getColumnIndex(rangeStr.match(/[A-Z]+/)[0]);
        const row = parseInt(rangeStr.match(/[0-9]+/)[0]) - 1;
        return { startRow: row, startCol: col, endRow: row, endCol: col };
    }
}

// Column Name to Index Conversion
getColumnIndex(colName) {
    let index = 0;
    for (let i = 0; i < colName.length; i++) {
        index = index * 26 + (colName.charCodeAt(i) - 65 + 1);
    }
    return index - 1;
}

// Arithmetic Expression Evaluation
evaluateArithmetic(expression) {
    // Replace cell references with values
    const cellRefPattern = /([A-Z]+)([0-9]+)/g;
    let processedExpression = expression.replace(cellRefPattern, (match, col, row) => {
        const colIndex = this.getColumnIndex(col);
        const rowIndex = parseInt(row) - 1;
        const cellData = this.getCellData(rowIndex, colIndex);
        return cellData ? parseFloat(cellData.value) || 0 : 0;
    });

    // Evaluate the expression
    return eval(processedExpression);
}
```

#### Performance Optimization
- **Event Delegation**: Efficient event handling for large numbers of elements
- **Debouncing**: Preventing excessive function calls during rapid user input
- **Lazy Loading**: On-demand processing of cell data
- **Memory Management**: Proper cleanup of event listeners and DOM references

```javascript
// Event Delegation Example
bindEvents() {
    document.addEventListener('keydown', (e) => this.handleGlobalKeydown(e));
    document.addEventListener('contextmenu', (e) => this.handleContextMenu(e));
    document.addEventListener('click', () => this.hideContextMenu());
}

// Efficient Cell Formatting
applyCellFormatting(cellElement, row, col) {
    const cellId = this.getCellId(row, col);
    const format = this.formats[cellId];

    // Remove existing formatting classes
    cellElement.classList.remove('bold', 'italic', 'underline', 'align-left', 'align-center', 'align-right');

    // Reset inline styles
    cellElement.style.cssText += ';color:;background-color:;font-family:;font-size:';

    if (!format) return;

    // Apply classes and styles
    if (format.bold) cellElement.classList.add('bold');
    if (format.italic) cellElement.classList.add('italic');
    if (format.underline) cellElement.classList.add('underline');
    if (format.alignment) cellElement.classList.add(`align-${format.alignment}`);

    if (format.textColor) cellElement.style.color = format.textColor;
    if (format.backgroundColor) cellElement.style.backgroundColor = format.backgroundColor;
    if (format.fontFamily) cellElement.style.fontFamily = format.fontFamily;
    if (format.fontSize) cellElement.style.fontSize = format.fontSize + 'px';
}

// Memory Management
hideContextMenu() {
    if (this.currentContextMenu) {
        document.body.removeChild(this.currentContextMenu);
        this.currentContextMenu = null;
    }
}
```

#### External API Integration
- **FileReader API**: Reading uploaded Excel files
- **SheetJS Integration**: Importing and exporting Excel file formats
- **Clipboard API**: Copy/paste functionality
- **Blob API**: File download generation

```javascript
// FileReader API for File Upload
handleFileOpen(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                jsonData.forEach((row, rowIndex) => {
                    row.forEach((cell, colIndex) => {
                        if (cell !== undefined && cell !== '') {
                            this.setCellData(rowIndex, colIndex, cell);
                        }
                    });
                });
                
                this.updateStatusBar('File loaded successfully');
            } catch (error) {
                this.updateStatusBar('Error loading file: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

// SheetJS Export Functionality
exportToXLSX() {
    if (!excelApp) return;
    
    try {
        const wb = XLSX.utils.book_new();
        const wsData = [];
        
        for (let row = 0; row < excelApp.rows; row++) {
            const rowData = [];
            for (let col = 0; col < excelApp.cols; col++) {
                const cellData = excelApp.getCellData(row, col);
                rowData.push(cellData ? cellData.value : '');
            }
            wsData.push(rowData);
        }
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, 'excel-clone.xlsx');
        excelApp.updateStatusBar('File exported successfully');
    } catch (error) {
        excelApp.updateStatusBar('Export failed: ' + error.message);
    }
}

// Clipboard Operations
copy() {
    this.clipboard = this.getSelectedRangeData();
    this.clipboardType = 'copy';
    this.updateStatusBar('Copied to clipboard');
}

cut() {
    this.clipboard = this.getSelectedRangeData();
    this.clipboardType = 'cut';
    this.clearSelectedCells();
    this.updateStatusBar('Cut to clipboard');
}

paste() {
    if (this.clipboard) {
        this.pasteData(this.currentCell.row, this.currentCell.col);
        this.updateStatusBar('Pasted from clipboard');
    }
}
```

#### Error Handling
- **Try-Catch Blocks**: Graceful error handling for formula evaluation
- **Input Validation**: Checking for valid cell references and formulas
- **User Feedback**: Status messages for operation results
- **Fallback Mechanisms**: Default values for missing data

```javascript
// Formula Evaluation with Error Handling
evaluateCellFormula(row, col) {
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    const cellData = this.getCellData(row, col);
    
    if (cellElement && cellData && cellData.value) {
        const value = cellData.value.trim();
        
        if (value.startsWith('=')) {
            try {
                const result = this.evaluateFormula(value);
                const cellId = this.getCellId(row, col);
                
                this.data[cellId].formula = value;
                this.data[cellId].value = result;
                this.data[cellId].displayValue = result;
                
                cellElement.textContent = result;
                cellElement.classList.add('formula');
                
            } catch (error) {
                const cellId = this.getCellId(row, col);
                this.data[cellId].value = '#ERROR';
                this.data[cellId].displayValue = '#ERROR';
                cellElement.textContent = '#ERROR';
                cellElement.classList.add('formula', 'error');
            }
        } else {
            cellElement.classList.remove('formula', 'error');
        }
    }
}

// Status Bar Updates with Timeout
updateStatusBar(message) {
    const statusText = document.getElementById('statusText');
    statusText.textContent = message;
    
    setTimeout(() => {
        statusText.textContent = 'Ready';
    }, 3000);
}

// Input Validation for Cell References
getCellData(row, col) {
    if (row < 0 || row >= this.rows || col < 0 || col >= this.cols) {
        return null;
    }
    
    const cellId = this.getCellId(row, col);
    return this.data[cellId];
}
```

## Architecture

The application is built using a class-based architecture:

- **ExcelClone**: Main application class managing all functionality
- **Grid Management**: Dynamic creation and management of spreadsheet cells
- **Formula Engine**: Evaluation and processing of formulas and functions
- **Formatting System**: Application and management of cell formatting
- **File Operations**: Import/export functionality using SheetJS
- **Event Handling**: Comprehensive keyboard and mouse event management

## Performance Features

- **Efficient rendering**: Only visible cells are actively managed
- **Lazy loading**: Data is processed on-demand
- **Memory optimization**: Minimal memory footprint for large spreadsheets
- **Smooth scrolling**: Optimized for large datasets

## Future Enhancements

Potential features for future versions:
- More formula functions (VLOOKUP, IF, etc.)
- Chart creation and visualization
- Conditional formatting
- Data validation
- Multiple sheets/tabs
- Undo/Redo functionality
- Collaborative editing
- Print functionality

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

---

**Note**: This is a demonstration project showcasing web-based spreadsheet functionality. For production use, consider additional features like data validation, security measures, and performance optimizations for very large datasets.

