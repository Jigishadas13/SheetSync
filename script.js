// Excel Clone - Main JavaScript File
class ExcelClone {
    constructor() {
        this.rows = 1000;
        this.cols = 26;
        this.currentCell = { row: 0, col: 0 };
        this.selectedRange = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
        this.data = {};
        this.formats = {};
        this.isEditing = false;
        this.clipboard = null;
        this.clipboardType = null; // 'copy' or 'cut'
        
        this.init();
    }

    init() {
        this.createGrid();
        this.createHeaders();
        this.bindEvents();
        this.updateCellDisplay();
        this.setupKeyboardShortcuts();
        this.createSampleData();
    }

    createGrid() {
        const dataGrid = document.getElementById('dataGrid');
        const rowHeaders = document.getElementById('rowHeaders');
        const columnHeaders = document.getElementById('columnHeaders');

        // Clear existing content
        dataGrid.innerHTML = '';
        rowHeaders.innerHTML = '';
        columnHeaders.innerHTML = '';

        // Set CSS variables for grid
        dataGrid.style.setProperty('--rows', this.rows);
        dataGrid.style.setProperty('--cols', this.cols);

        // Create column headers (A, B, C, ...)
        for (let col = 0; col < this.cols; col++) {
            const headerCell = document.createElement('div');
            headerCell.className = 'column-header-cell';
            headerCell.textContent = this.getColumnName(col);
            headerCell.onclick = () => this.selectColumn(col);
            columnHeaders.appendChild(headerCell);
        }

        // Create row headers (1, 2, 3, ...)
        for (let row = 0; row < this.rows; row++) {
            const headerCell = document.createElement('div');
            headerCell.className = 'row-header-cell';
            headerCell.textContent = row + 1;
            headerCell.onclick = () => this.selectRow(row);
            rowHeaders.appendChild(headerCell);
        }

        // Create data cells
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.dataset.row = row;
                cell.dataset.col = col;
                cell.tabIndex = 0;
                cell.contentEditable = true;
                
                // Event listeners for cell
                cell.onclick = (e) => this.selectCell(row, col, e);
                cell.onkeydown = (e) => this.handleCellKeydown(e);
                cell.oninput = (e) => this.handleCellInput(e);
                cell.onfocus = () => this.focusCell(row, col);
                cell.onblur = () => this.blurCell(row, col);
                cell.onpaste = (e) => this.handleCellPaste(e);
                
                dataGrid.appendChild(cell);
            }
        }
    }

    createHeaders() {
        // Headers are created in createGrid method
    }

    bindEvents() {
        // Global keyboard events
        document.addEventListener('keydown', (e) => this.handleGlobalKeydown(e));
        
        // Context menu
        document.addEventListener('contextmenu', (e) => this.handleContextMenu(e));
        document.addEventListener('click', () => this.hideContextMenu());

        // Formula input events
        const formulaInput = document.getElementById('formulaInput');
        formulaInput.addEventListener('keydown', (e) => this.handleFormulaKeydown(e));

        // File input
        const fileInput = document.getElementById('fileInput');
        fileInput.addEventListener('change', (e) => this.handleFileOpen(e));
    }

    setupKeyboardShortcuts() {
        // This will be handled in handleGlobalKeydown
    }

    createSampleData() {
        // Add some sample data for testing formulas
        this.setCellData(0, 0, 'Name');
        this.setCellData(0, 1, 'Age');
        this.setCellData(0, 2, 'Salary');
        
        this.setCellData(1, 0, 'John');
        this.setCellData(1, 1, '25');
        this.setCellData(1, 2, '50000');
        
        this.setCellData(2, 0, 'Jane');
        this.setCellData(2, 1, '30');
        this.setCellData(2, 2, '60000');
        
        this.setCellData(3, 0, 'Bob');
        this.setCellData(3, 1, '35');
        this.setCellData(3, 2, '70000');
        
        // Add some numbers for testing SUM
        this.setCellData(5, 0, '10');
        this.setCellData(5, 1, '20');
        this.setCellData(5, 2, '30');
        this.setCellData(6, 0, '40');
        this.setCellData(6, 1, '50');
        this.setCellData(6, 2, '60');
        
        // Update all cell displays
        this.updateAllCellDisplays();
    }

    updateAllCellDisplays() {
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                this.updateCellValue(row, col);
            }
        }
    }

    getColumnName(col) {
        let result = '';
        while (col >= 0) {
            result = String.fromCharCode(65 + (col % 26)) + result;
            col = Math.floor(col / 26) - 1;
        }
        return result;
    }

    getCellId(row, col) {
        return `${this.getColumnName(col)}${row + 1}`;
    }

    selectCell(row, col, event = null) {
        if (event && event.shiftKey) {
            this.selectRange(row, col);
        } else {
            this.currentCell = { row, col };
            this.selectedRange = { start: { row, col }, end: { row, col } };
            this.isEditing = false;
        }
        
        this.updateCellDisplay();
        this.updateSelectionIndicator();
        this.updateFormulaBar();
        this.updateStatusBar();
        
        // Focus the cell for direct typing
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (cellElement) {
            cellElement.focus();
        }
    }

    selectRange(row, col) {
        this.selectedRange.end = { row, col };
        this.updateSelectionIndicator();
    }

    selectRow(row) {
        this.selectedRange = { 
            start: { row, col: 0 }, 
            end: { row, col: this.cols - 1 } 
        };
        this.currentCell = { row, col: 0 };
        this.updateCellDisplay();
        this.updateSelectionIndicator();
        this.updateFormulaBar();
    }

    selectColumn(col) {
        this.selectedRange = { 
            start: { row: 0, col }, 
            end: { row: this.rows - 1, col } 
        };
        this.currentCell = { row: 0, col };
        this.updateCellDisplay();
        this.updateSelectionIndicator();
        this.updateFormulaBar();
    }

    focusCell(row, col) {
        this.currentCell = { row, col };
        this.updateFormulaBar();
    }

    blurCell(row, col) {
        // Evaluate formula when cell loses focus
        this.evaluateCellFormula(row, col);
    }

    updateCellDisplay() {
        // Remove selected class from all cells
        document.querySelectorAll('.cell').forEach(cell => {
            cell.classList.remove('selected');
        });

        // Add selected class to current cell
        const currentCellElement = document.querySelector(`[data-row="${this.currentCell.row}"][data-col="${this.currentCell.col}"]`);
        if (currentCellElement) {
            currentCellElement.classList.add('selected');
        }

        // Add selected class to range
        const minRow = Math.min(this.selectedRange.start.row, this.selectedRange.end.row);
        const maxRow = Math.max(this.selectedRange.start.row, this.selectedRange.end.row);
        const minCol = Math.min(this.selectedRange.start.col, this.selectedRange.end.col);
        const maxCol = Math.max(this.selectedRange.start.col, this.selectedRange.end.col);

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
                if (cellElement) {
                    cellElement.classList.add('selected');
                }
            }
        }
    }

    updateSelectionIndicator() {
        const indicator = document.getElementById('selectionIndicator');
        const container = document.querySelector('.spreadsheet-container');
        
        const minRow = Math.min(this.selectedRange.start.row, this.selectedRange.end.row);
        const maxRow = Math.max(this.selectedRange.start.row, this.selectedRange.end.row);
        const minCol = Math.min(this.selectedRange.start.col, this.selectedRange.end.col);
        const maxCol = Math.max(this.selectedRange.start.col, this.selectedRange.end.col);

        const left = minCol * 100 + 40;
        const top = minRow * 25 + 25;
        const width = (maxCol - minCol + 1) * 100;
        const height = (maxRow - minRow + 1) * 25;

        indicator.style.left = left + 'px';
        indicator.style.top = top + 'px';
        indicator.style.width = width + 'px';
        indicator.style.height = height + 'px';
    }

    updateFormulaBar() {
        const cellAddress = document.getElementById('cellAddress');
        const formulaInput = document.getElementById('formulaInput');
        
        cellAddress.textContent = this.getCellId(this.currentCell.row, this.currentCell.col);
        
        const cellData = this.getCellData(this.currentCell.row, this.currentCell.col);
        if (cellData && cellData.formula) {
            formulaInput.value = cellData.formula;
        } else {
            formulaInput.value = cellData ? cellData.value : '';
        }
        
        // Focus the formula input when cell is selected
        formulaInput.focus();
        
        // Update toolbar button states
        this.updateToolbarButtons();
    }

    updateToolbarButtons() {
        const cellId = this.getCellId(this.currentCell.row, this.currentCell.col);
        const format = this.formats[cellId];
        
        // Update bold button
        const boldBtn = document.getElementById('boldBtn');
        if (format && format.bold) {
            boldBtn.classList.add('active');
        } else {
            boldBtn.classList.remove('active');
        }
        
        // Update italic button
        const italicBtn = document.getElementById('italicBtn');
        if (format && format.italic) {
            italicBtn.classList.add('active');
        } else {
            italicBtn.classList.remove('active');
        }
        
        // Update underline button
        const underlineBtn = document.getElementById('underlineBtn');
        if (format && format.underline) {
            underlineBtn.classList.add('active');
        } else {
            underlineBtn.classList.remove('active');
        }
        
        // Update font family dropdown
        const fontFamilySelect = document.getElementById('fontFamily');
        if (format && format.fontFamily) {
            fontFamilySelect.value = format.fontFamily;
        } else {
            fontFamilySelect.value = 'Arial';
        }
        
        // Update font size dropdown
        const fontSizeSelect = document.getElementById('fontSize');
        if (format && format.fontSize) {
            fontSizeSelect.value = format.fontSize;
        } else {
            fontSizeSelect.value = '12';
        }
    }

    updateStatusBar() {
        const statusText = document.getElementById('statusText');
        const selectionInfo = document.getElementById('selectionInfo');
        
        const minRow = Math.min(this.selectedRange.start.row, this.selectedRange.end.row);
        const maxRow = Math.max(this.selectedRange.start.row, this.selectedRange.end.row);
        const minCol = Math.min(this.selectedRange.start.col, this.selectedRange.end.col);
        const maxCol = Math.max(this.selectedRange.start.col, this.selectedRange.end.col);

        if (minRow === maxRow && minCol === maxCol) {
            selectionInfo.textContent = `Cell: ${this.getCellId(this.currentCell.row, this.currentCell.col)}`;
        } else {
            selectionInfo.textContent = `Range: ${this.getCellId(minRow, minCol)}:${this.getCellId(maxRow, maxCol)}`;
        }
    }

    getCellData(row, col) {
        const cellId = this.getCellId(row, col);
        return this.data[cellId];
    }

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
        
        this.updateCellValue(row, col);
    }

    updateCellValue(row, col) {
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        const cellData = this.getCellData(row, col);
        
        if (cellElement) {
            if (cellData) {
                cellElement.textContent = cellData.displayValue || cellData.value || '';
            } else {
                cellElement.textContent = '';
            }
            
            // Apply formatting
            this.applyCellFormatting(cellElement, row, col);
        }
    }

    applyCellFormatting(cellElement, row, col) {
        const cellId = this.getCellId(row, col);
        const format = this.formats[cellId];

        // Remove existing formatting classes (class churn kept minimal)
        cellElement.classList.remove('bold', 'italic', 'underline', 'align-left', 'align-center', 'align-right', 'align-top', 'align-middle', 'align-bottom');

        // Reset inline styles in one go
        cellElement.style.cssText += ';color:;background-color:;font-family:;font-size:';

        if (!format) return;

        // Apply classes
        if (format.bold) cellElement.classList.add('bold');
        if (format.italic) cellElement.classList.add('italic');
        if (format.underline) cellElement.classList.add('underline');
        if (format.alignment) cellElement.classList.add(`align-${format.alignment}`);

        // Apply inline styles (fast and avoids layout thrash)
        if (format.textColor) cellElement.style.color = format.textColor;
        if (format.backgroundColor) cellElement.style.backgroundColor = format.backgroundColor;
        if (format.fontFamily) cellElement.style.fontFamily = format.fontFamily;
        if (format.fontSize) cellElement.style.fontSize = format.fontSize + 'px';
    }

    handleCellKeydown(event) {
        // Allow typing characters and numbers
        if (event.key.length === 1 || event.key === 'Backspace' || event.key === 'Delete') {
            // Allow normal typing - don't prevent default
            this.isEditing = true;
            return;
        }

        switch (event.key) {
            case 'Enter':
                event.preventDefault();
                this.evaluateCellFormula(this.currentCell.row, this.currentCell.col);
                this.isEditing = false;
                this.moveToCell(this.currentCell.row + 1, this.currentCell.col);
                break;
            case 'Tab':
                event.preventDefault();
                this.evaluateCellFormula(this.currentCell.row, this.currentCell.col);
                this.isEditing = false;
                this.moveToCell(this.currentCell.row, this.currentCell.col + 1);
                break;
            case 'ArrowUp':
                if (!this.isEditing) {
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row - 1, this.currentCell.col);
                }
                break;
            case 'ArrowDown':
                if (!this.isEditing) {
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row + 1, this.currentCell.col);
                }
                break;
            case 'ArrowLeft':
                if (!this.isEditing) {
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row, this.currentCell.col - 1);
                }
                break;
            case 'ArrowRight':
                if (!this.isEditing) {
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row, this.currentCell.col + 1);
                }
                break;
            case 'F2':
                event.preventDefault();
                this.startEditing();
                break;
            case 'Escape':
                event.preventDefault();
                this.cancelEditing();
                break;
        }
    }

    handleCellInput(event) {
        // Don't process formulas during typing - only store the raw input
        const cellElement = event.target;
        const row = parseInt(cellElement.dataset.row);
        const col = parseInt(cellElement.dataset.col);
        
        // Just store the raw text content without processing formulas
        let value = cellElement.textContent.trim();
        this.setCellData(row, col, value);
    }

    handleCellPaste(event) {
        event.preventDefault();
        const text = (event.clipboardData || window.clipboardData).getData('text');
        const cellElement = event.target;
        const row = parseInt(cellElement.dataset.row);
        const col = parseInt(cellElement.dataset.col);
        
        // Insert the pasted text
        cellElement.textContent = text;
        this.setCellData(row, col, text);
    }

    evaluateCellFormula(row, col) {
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        const cellData = this.getCellData(row, col);
        
        if (cellElement && cellData && cellData.value) {
            const value = cellData.value.trim();
            
            // Check if it's a formula (starts with =)
            if (value.startsWith('=')) {
                try {
                    const result = this.evaluateFormula(value);
                    const cellId = this.getCellId(row, col);
                    
                    // Store both the formula and the result
                    this.data[cellId].formula = value;
                    this.data[cellId].value = result;
                    this.data[cellId].displayValue = result;
                    
                    // Update the display
                    cellElement.textContent = result;
                    
                    // Mark as formula cell for styling
                    cellElement.classList.add('formula');
                    
                } catch (error) {
                    // Show error in cell
                    const cellId = this.getCellId(row, col);
                    this.data[cellId].value = '#ERROR';
                    this.data[cellId].displayValue = '#ERROR';
                    cellElement.textContent = '#ERROR';
                    cellElement.classList.add('formula', 'error');
                }
            } else {
                // Remove formula styling for regular text
                cellElement.classList.remove('formula', 'error');
            }
        }
    }

    handleGlobalKeydown(event) {
        // Handle keyboard shortcuts
        if (event.ctrlKey || event.metaKey) {
            switch (event.key.toLowerCase()) {
                case 'c':
                    event.preventDefault();
                    this.copy();
                    break;
                case 'x':
                    event.preventDefault();
                    this.cut();
                    break;
                case 'v':
                    event.preventDefault();
                    this.paste();
                    break;
                case 'z':
                    event.preventDefault();
                    // Undo functionality (not implemented yet)
                    break;
                case 'y':
                    event.preventDefault();
                    // Redo functionality (not implemented yet)
                    break;
                case 'n':
                    event.preventDefault();
                    this.newFile();
                    break;
                case 'o':
                    event.preventDefault();
                    this.openFile();
                    break;
                case 's':
                    event.preventDefault();
                    this.saveFile();
                    break;
            }
        }

        // Handle arrow keys for navigation
        if (!event.ctrlKey && !event.metaKey && !this.isEditing) {
            switch (event.key) {
                case 'ArrowUp':
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row - 1, this.currentCell.col);
                    break;
                case 'ArrowDown':
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row + 1, this.currentCell.col);
                    break;
                case 'ArrowLeft':
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row, this.currentCell.col - 1);
                    break;
                case 'ArrowRight':
                    event.preventDefault();
                    this.moveToCell(this.currentCell.row, this.currentCell.col + 1);
                    break;
            }
        }
    }

    moveToCell(row, col) {
        // Clamp to grid boundaries
        row = Math.max(0, Math.min(row, this.rows - 1));
        col = Math.max(0, Math.min(col, this.cols - 1));
        
        this.selectCell(row, col);
        
        // Scroll into view
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (cellElement) {
            cellElement.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
        }
    }

    startEditing() {
        const cellElement = document.querySelector(`[data-row="${this.currentCell.row}"][data-col="${this.currentCell.col}"]`);
        if (cellElement) {
            cellElement.classList.add('editing');
            cellElement.focus();
            this.isEditing = true;
            
            // Select all text
            const range = document.createRange();
            range.selectNodeContents(cellElement);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
        }
    }

    cancelEditing() {
        this.isEditing = false;
        document.querySelectorAll('.cell').forEach(cell => {
            cell.classList.remove('editing');
        });
    }

    handleFormulaInput(event) {
        if (event.key === 'Enter') {
            const formulaInput = document.getElementById('formulaInput');
            this.setCellData(this.currentCell.row, this.currentCell.col, formulaInput.value);
            this.updateFormulaBar();
            this.moveToCell(this.currentCell.row + 1, this.currentCell.col);
        }
    }

    handleFormulaKeydown(event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const formulaInput = document.getElementById('formulaInput');
            this.setCellData(this.currentCell.row, this.currentCell.col, formulaInput.value);
            this.updateFormulaBar();
        }
    }

    clearSelectedCells() {
        const minRow = Math.min(this.selectedRange.start.row, this.selectedRange.end.row);
        const maxRow = Math.max(this.selectedRange.start.row, this.selectedRange.end.row);
        const minCol = Math.min(this.selectedRange.start.col, this.selectedRange.end.col);
        const maxCol = Math.max(this.selectedRange.start.col, this.selectedRange.end.col);

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cellId = this.getCellId(row, col);
                delete this.data[cellId];
                delete this.formats[cellId];
                
                const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
                if (cellElement) {
                    cellElement.textContent = '';
                    cellElement.className = 'cell';
                    cellElement.style.color = '';
                    cellElement.style.backgroundColor = '';
                }
            }
        }
    }

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

    // Formula evaluation
    evaluateFormula(formula) {
        if (!formula || !formula.startsWith('=')) {
            return formula;
        }

        try {
            // Remove the = sign
            const expression = formula.substring(1);
            
            // Handle basic formulas
            if (expression.toUpperCase().startsWith('SUM(')) {
                return this.evaluateSum(expression);
            } else if (expression.toUpperCase().startsWith('AVERAGE(')) {
                return this.evaluateAverage(expression);
            } else if (expression.toUpperCase().startsWith('COUNT(')) {
                return this.evaluateCount(expression);
            } else if (expression.toUpperCase().startsWith('MAX(')) {
                return this.evaluateMax(expression);
            } else if (expression.toUpperCase().startsWith('MIN(')) {
                return this.evaluateMin(expression);
            } else {
                // Handle basic arithmetic
                return this.evaluateArithmetic(expression);
            }
        } catch (error) {
            return '#ERROR';
        }
    }

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

    evaluateAverage(expression) {
        const range = this.parseRange(expression.match(/AVERAGE\(([^)]+)\)/i)[1]);
        let sum = 0;
        let count = 0;
        
        for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
                const cellData = this.getCellData(row, col);
                const value = parseFloat(cellData ? cellData.value : 0);
                if (!isNaN(value)) {
                    sum += value;
                    count++;
                }
            }
        }
        
        return count > 0 ? sum / count : 0;
    }

    evaluateCount(expression) {
        const range = this.parseRange(expression.match(/COUNT\(([^)]+)\)/i)[1]);
        let count = 0;
        
        for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
                const cellData = this.getCellData(row, col);
                if (cellData && cellData.value !== '') {
                    count++;
                }
            }
        }
        
        return count;
    }

    evaluateMax(expression) {
        const range = this.parseRange(expression.match(/MAX\(([^)]+)\)/i)[1]);
        let max = -Infinity;
        
        for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
                const cellData = this.getCellData(row, col);
                const value = parseFloat(cellData ? cellData.value : -Infinity);
                if (!isNaN(value)) {
                    max = Math.max(max, value);
                }
            }
        }
        
        return max === -Infinity ? 0 : max;
    }

    evaluateMin(expression) {
        const range = this.parseRange(expression.match(/MIN\(([^)]+)\)/i)[1]);
        let min = Infinity;
        
        for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
                const cellData = this.getCellData(row, col);
                const value = parseFloat(cellData ? cellData.value : Infinity);
                if (!isNaN(value)) {
                    min = Math.min(min, value);
                }
            }
        }
        
        return min === Infinity ? 0 : min;
    }

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

    getColumnIndex(colName) {
        let index = 0;
        for (let i = 0; i < colName.length; i++) {
            index = index * 26 + (colName.charCodeAt(i) - 65 + 1);
        }
        return index - 1;
    }

    handleContextMenu(event) {
        event.preventDefault();
        
        // Create context menu
        const contextMenu = document.createElement('div');
        contextMenu.className = 'context-menu';
        contextMenu.style.display = 'block';
        contextMenu.style.left = event.pageX + 'px';
        contextMenu.style.top = event.pageY + 'px';
        
        const menuItems = [
            { text: 'Copy', action: () => this.copy() },
            { text: 'Cut', action: () => this.cut() },
            { text: 'Paste', action: () => this.paste() },
            { text: 'Clear', action: () => this.clearSelectedCells() },
            { text: 'Insert Row', action: () => this.insertRow() },
            { text: 'Insert Column', action: () => this.insertColumn() },
            { text: 'Delete Row', action: () => this.deleteRow() },
            { text: 'Delete Column', action: () => this.deleteColumn() }
        ];
        
        menuItems.forEach(item => {
            const menuItem = document.createElement('div');
            menuItem.className = 'context-menu-item';
            menuItem.textContent = item.text;
            menuItem.onclick = () => {
                item.action();
                this.hideContextMenu();
            };
            contextMenu.appendChild(menuItem);
        });
        
        document.body.appendChild(contextMenu);
        
        // Store reference for cleanup
        this.currentContextMenu = contextMenu;
    }

    hideContextMenu() {
        if (this.currentContextMenu) {
            document.body.removeChild(this.currentContextMenu);
            this.currentContextMenu = null;
        }
    }

    handleFileOpen(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // Clear existing data
                    this.data = {};
                    this.formats = {};
                    
                    // Parse worksheet data
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

    updateStatusBar(message) {
        const statusText = document.getElementById('statusText');
        statusText.textContent = message;
        
        // Clear message after 3 seconds
        setTimeout(() => {
            statusText.textContent = 'Ready';
        }, 3000);
    }
}

// Global functions for HTML event handlers
let excelApp;

function newFile() {
    excelApp = new ExcelClone();
    excelApp.updateStatusBar('New file created');
}

function openFile() {
    document.getElementById('fileInput').click();
}

function saveFile() {
    excelApp.updateStatusBar('Save functionality - use Export to XLSX');
}

function exportToXLSX() {
    if (!excelApp) return;
    
    try {
        // Create workbook
        const wb = XLSX.utils.book_new();
        
        // Convert data to worksheet format
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
        
        // Generate and download file
        XLSX.writeFile(wb, 'excel-clone.xlsx');
        excelApp.updateStatusBar('File exported successfully');
    } catch (error) {
        excelApp.updateStatusBar('Export failed: ' + error.message);
    }
}

function handleFileOpen(event) {
    if (excelApp) {
        excelApp.handleFileOpen(event);
    }
}

// Formatting functions
function toggleBold() {
    if (!excelApp) return;
    
    // Get current cell for testing
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    
    
    // Initialize format if it doesn't exist
    if (!excelApp.formats[cellId]) {
        excelApp.formats[cellId] = {};
    }
    
    // Toggle bold
    excelApp.formats[cellId].bold = !excelApp.formats[cellId].bold;
    
    
    // Apply formatting
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (cellElement) {
        excelApp.applyCellFormatting(cellElement, row, col);
    }
    
    // Update button state
    const boldBtn = document.getElementById('boldBtn');
    if (excelApp.formats[cellId].bold) {
        boldBtn.classList.add('active');
    } else {
        boldBtn.classList.remove('active');
    }
    
}

function toggleItalic() {
    if (!excelApp) return;
    
    // Get current cell for testing
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    
    // Initialize format if it doesn't exist
    if (!excelApp.formats[cellId]) {
        excelApp.formats[cellId] = {};
    }
    
    // Toggle italic
    excelApp.formats[cellId].italic = !excelApp.formats[cellId].italic;
    
    // Apply formatting
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (cellElement) {
        excelApp.applyCellFormatting(cellElement, row, col);
    }
    
    // Update button state
    const italicBtn = document.getElementById('italicBtn');
    if (excelApp.formats[cellId].italic) {
        italicBtn.classList.add('active');
    } else {
        italicBtn.classList.remove('active');
    }
}

function toggleUnderline() {
    if (!excelApp) return;
    
    // Get current cell for testing
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    
    // Initialize format if it doesn't exist
    if (!excelApp.formats[cellId]) {
        excelApp.formats[cellId] = {};
    }
    
    // Toggle underline
    excelApp.formats[cellId].underline = !excelApp.formats[cellId].underline;
    
    // Apply formatting
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (cellElement) {
        excelApp.applyCellFormatting(cellElement, row, col);
    }
    
    // Update button state
    const underlineBtn = document.getElementById('underlineBtn');
    if (excelApp.formats[cellId].underline) {
        underlineBtn.classList.add('active');
    } else {
        underlineBtn.classList.remove('active');
    }
}

function setAlignment(alignment) {
    if (!excelApp) return;
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    if (!excelApp.formats[cellId]) excelApp.formats[cellId] = {};
    excelApp.formats[cellId].alignment = alignment;
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (cellElement) excelApp.applyCellFormatting(cellElement, row, col);
}

function changeFontFamily(fontFamily) {
    if (!excelApp) return;
    
    const minRow = Math.min(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const maxRow = Math.max(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const minCol = Math.min(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);
    const maxCol = Math.max(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);

    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cellId = excelApp.getCellId(row, col);
            if (!excelApp.formats[cellId]) {
                excelApp.formats[cellId] = {};
            }
            excelApp.formats[cellId].fontFamily = fontFamily;
            
            const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            if (cellElement) {
                excelApp.applyCellFormatting(cellElement, row, col);
            }
        }
    }
}

function changeFontSize(fontSize) {
    if (!excelApp) return;
    
    const minRow = Math.min(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const maxRow = Math.max(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const minCol = Math.min(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);
    const maxCol = Math.max(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);

    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cellId = excelApp.getCellId(row, col);
            if (!excelApp.formats[cellId]) {
                excelApp.formats[cellId] = {};
            }
            excelApp.formats[cellId].fontSize = fontSize;
            
            const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            if (cellElement) {
                excelApp.applyCellFormatting(cellElement, row, col);
            }
        }
    }
}

function changeTextColor(color) {
    if (!excelApp) return;
    
    const minRow = Math.min(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const maxRow = Math.max(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const minCol = Math.min(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);
    const maxCol = Math.max(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);

    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cellId = excelApp.getCellId(row, col);
            if (!excelApp.formats[cellId]) {
                excelApp.formats[cellId] = {};
            }
            excelApp.formats[cellId].textColor = color;
            
            const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            if (cellElement) {
                excelApp.applyCellFormatting(cellElement, row, col);
            }
        }
    }
}

function changeBackgroundColor(color) {
    if (!excelApp) return;
    
    const minRow = Math.min(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const maxRow = Math.max(excelApp.selectedRange.start.row, excelApp.selectedRange.end.row);
    const minCol = Math.min(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);
    const maxCol = Math.max(excelApp.selectedRange.start.col, excelApp.selectedRange.end.col);

    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cellId = excelApp.getCellId(row, col);
            if (!excelApp.formats[cellId]) {
                excelApp.formats[cellId] = {};
            }
            excelApp.formats[cellId].backgroundColor = color;
            
            const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            if (cellElement) {
                excelApp.applyCellFormatting(cellElement, row, col);
            }
        }
    }
}

function toggleBorder(type) {
    if (!excelApp) return;
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    if (!excelApp.formats[cellId]) excelApp.formats[cellId] = {};
    excelApp.formats[cellId].borderType = type;
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (!cellElement) return;
    // Apply border directly (minimal work)
    switch (type) {
        case 'all':
            cellElement.style.border = '1px solid #000';
            break;
        case 'outline':
            cellElement.style.border = '2px solid #000';
            break;
        case 'inside':
            cellElement.style.border = '3px solid #ccc';
            break;
    }
}

function clearBorders() {
    if (!excelApp) return;
    const row = excelApp.currentCell.row;
    const col = excelApp.currentCell.col;
    const cellId = excelApp.getCellId(row, col);
    if (excelApp.formats[cellId]) delete excelApp.formats[cellId].borderType;
    const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
    if (cellElement) cellElement.style.border = '1px solid #e9ecef';
}

function insertRow() {
    if (!excelApp) return;
    
    
    const insertAt = excelApp.currentCell.row;
    
    // Shift data down
    for (let row = excelApp.rows - 2; row >= insertAt; row--) {
        for (let col = 0; col < excelApp.cols; col++) {
            const currentCellId = excelApp.getCellId(row, col);
            const newCellId = excelApp.getCellId(row + 1, col);
            
            if (excelApp.data[currentCellId]) {
                excelApp.data[newCellId] = { ...excelApp.data[currentCellId] };
                delete excelApp.data[currentCellId];
            }
            
            if (excelApp.formats[currentCellId]) {
                excelApp.formats[newCellId] = { ...excelApp.formats[currentCellId] };
                delete excelApp.formats[currentCellId];
            }
        }
    }
    
    // Update display
    excelApp.updateAllCellDisplays();
    excelApp.updateStatusBar('Row inserted');
}

function deleteRow() {
    if (!excelApp) return;
    
    
    const deleteAt = excelApp.currentCell.row;
    
    // Shift data up
    for (let row = deleteAt; row < excelApp.rows - 1; row++) {
        for (let col = 0; col < excelApp.cols; col++) {
            const currentCellId = excelApp.getCellId(row + 1, col);
            const newCellId = excelApp.getCellId(row, col);
            
            if (excelApp.data[currentCellId]) {
                excelApp.data[newCellId] = { ...excelApp.data[currentCellId] };
                delete excelApp.data[currentCellId];
            } else {
                delete excelApp.data[newCellId];
            }
            
            if (excelApp.formats[currentCellId]) {
                excelApp.formats[newCellId] = { ...excelApp.formats[currentCellId] };
                delete excelApp.formats[currentCellId];
            } else {
                delete excelApp.formats[newCellId];
            }
        }
    }
    
    // Update display
    excelApp.updateAllCellDisplays();
    excelApp.updateStatusBar('Row deleted');
}

function insertColumn() {
    if (!excelApp) return;
    
    
    const insertAt = excelApp.currentCell.col;
    
    // Shift data right
    for (let row = 0; row < excelApp.rows; row++) {
        for (let col = excelApp.cols - 2; col >= insertAt; col--) {
            const currentCellId = excelApp.getCellId(row, col);
            const newCellId = excelApp.getCellId(row, col + 1);
            
            if (excelApp.data[currentCellId]) {
                excelApp.data[newCellId] = { ...excelApp.data[currentCellId] };
                delete excelApp.data[currentCellId];
            }
            
            if (excelApp.formats[currentCellId]) {
                excelApp.formats[newCellId] = { ...excelApp.formats[currentCellId] };
                delete excelApp.formats[currentCellId];
            }
        }
    }
    
    // Update display
    excelApp.updateAllCellDisplays();
    excelApp.updateStatusBar('Column inserted');
}

function deleteColumn() {
    if (!excelApp) return;
    
    
    const deleteAt = excelApp.currentCell.col;
    
    // Shift data left
    for (let row = 0; row < excelApp.rows; row++) {
        for (let col = deleteAt; col < excelApp.cols - 1; col++) {
            const currentCellId = excelApp.getCellId(row, col + 1);
            const newCellId = excelApp.getCellId(row, col);
            
            if (excelApp.data[currentCellId]) {
                excelApp.data[newCellId] = { ...excelApp.data[currentCellId] };
                delete excelApp.data[currentCellId];
            } else {
                delete excelApp.data[newCellId];
            }
            
            if (excelApp.formats[currentCellId]) {
                excelApp.formats[newCellId] = { ...excelApp.formats[currentCellId] };
                delete excelApp.formats[currentCellId];
            } else {
                delete excelApp.formats[newCellId];
            }
        }
    }
    
    // Update display
    excelApp.updateAllCellDisplays();
    excelApp.updateStatusBar('Column deleted');
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', () => {
    excelApp = new ExcelClone();
});
