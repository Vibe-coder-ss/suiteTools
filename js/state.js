/**
 * Global State Management
 * Stores all application state variables
 */

const AppState = {
    // Current mode: 'spreadsheet', 'document', or 'none'
    currentMode: 'none',
    
    // Spreadsheet state
    currentData: [],
    headers: [],
    allSheets: [],
    currentSheetIndex: 0,
    currentFile: null,
    fileHandle: null,
    
    // SQL state
    sqlTablesRegistered: false,
    hasSqlResults: false,
    sqlResultData: null,
    
    // Edit state
    hasUnsavedChanges: false,
    originalSheetData: [],
    
    // Auto-save state
    autoSaveEnabled: false,
    autoSaveTimeout: null,
    lastSaveTime: null,
    
    // Sort state
    sortColumn: -1,
    sortDirection: 'none',
    
    // Filter state
    activeFilters: [],
    
    // Export state
    selectedFormat: 'xlsx',
    
    // Document state
    docHasUnsavedChanges: false,
    originalDocContent: '',
    docFileHandle: null,
    
    // Slides state
    slidesHasUnsavedChanges: false,
    slidesData: [],           // Array of slide objects
    currentSlideIndex: 0,     // Currently displayed slide
    slidesFileHandle: null,
    slidesFile: null,
    
    // Reset spreadsheet state
    resetSpreadsheet() {
        this.currentData = [];
        this.headers = [];
        this.allSheets = [];
        this.currentSheetIndex = 0;
        this.currentFile = null;
        this.fileHandle = null;
        this.sqlTablesRegistered = false;
        this.hasSqlResults = false;
        this.sqlResultData = null;
        this.hasUnsavedChanges = false;
        this.originalSheetData = [];
        this.autoSaveEnabled = false;
        this.autoSaveTimeout = null;
        this.sortColumn = -1;
        this.sortDirection = 'none';
        this.activeFilters = [];
    },
    
    // Reset document state
    resetDocument() {
        this.docHasUnsavedChanges = false;
        this.originalDocContent = '';
        this.docFileHandle = null;
    },
    
    // Reset slides state
    resetSlides() {
        this.slidesHasUnsavedChanges = false;
        this.slidesData = [];
        this.currentSlideIndex = 0;
        this.slidesFileHandle = null;
        this.slidesFile = null;
    },
    
    // Reset all state
    resetAll() {
        this.currentMode = 'none';
        this.resetSpreadsheet();
        this.resetDocument();
        this.resetSlides();
    }
};

// Export for use in other modules
window.AppState = AppState;