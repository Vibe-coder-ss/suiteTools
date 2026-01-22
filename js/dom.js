/**
 * DOM Elements
 * Centralized DOM element references
 */

const DOM = {
    // Initialize all DOM references
    init() {
        // Theme
        this.themeToggle = document.getElementById('themeToggle');
        
        // Upload/Drop zone
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.uploadSection = document.getElementById('uploadSection');
        
        // File info bar
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileSize = document.getElementById('fileSize');
        this.rowCount = document.getElementById('rowCount');
        this.clearBtn = document.getElementById('clearBtn');
        
        // Spreadsheet
        this.tableContainer = document.getElementById('tableContainer');
        this.sheetTabs = document.getElementById('sheetTabs');
        this.searchInput = document.getElementById('searchInput');
        
        // Edit controls
        this.addRowBtn = document.getElementById('addRowBtn');
        this.addColBtn = document.getElementById('addColBtn');
        this.saveBtn = document.getElementById('saveBtn');
        this.undoAllBtn = document.getElementById('undoAllBtn');
        this.editIndicator = document.getElementById('editIndicator');
        this.saveStatus = document.getElementById('saveStatus');
        this.autosaveCheckbox = document.getElementById('autosaveCheckbox');
        this.autosaveToggle = document.getElementById('autosaveToggle');
        
        // Export modal
        this.exportBtn = document.getElementById('exportBtn');
        this.exportModal = document.getElementById('exportModal');
        this.modalClose = document.getElementById('modalClose');
        this.exportCancel = document.getElementById('exportCancel');
        this.exportConfirm = document.getElementById('exportConfirm');
        this.formatGrid = document.getElementById('formatGrid');
        this.formatNote = document.getElementById('formatNote');
        this.sheetCheckboxes = document.getElementById('sheetCheckboxes');
        this.currentSheetNameBadge = document.getElementById('currentSheetName');
        this.sheetCountBadge = document.getElementById('sheetCountBadge');
        this.allSheetsOption = document.getElementById('allSheetsOption');
        this.sheetsChecklist = document.getElementById('sheetsChecklist');
        
        // SQL elements
        this.sqlSection = document.getElementById('sqlSection');
        this.sqlInput = document.getElementById('sqlInput');
        this.sqlRunBtn = document.getElementById('sqlRunBtn');
        this.sqlExamplesBtn = document.getElementById('sqlExamplesBtn');
        this.sqlExamplesPanel = document.getElementById('sqlExamplesPanel');
        this.sqlStatus = document.getElementById('sqlStatus');
        this.sqlResultBar = document.getElementById('sqlResultBar');
        this.sqlResetBtn = document.getElementById('sqlResetBtn');
        this.sqlExpandBtn = document.getElementById('sqlExpandBtn');
        this.sqlEditorWrapper = document.getElementById('sqlEditorWrapper');
        this.sqlLineCount = document.getElementById('sqlLineCount');
        this.sqlTablesHint = document.getElementById('sqlTablesHint');
        this.sqlExportResultsBtn = document.getElementById('sqlExportResultsBtn');
        
        // Quick filter
        this.filterToggleBtn = document.getElementById('filterToggleBtn');
        this.quickFilterPanel = document.getElementById('quickFilterPanel');
        this.filterColumn = document.getElementById('filterColumn');
        this.filterOperator = document.getElementById('filterOperator');
        this.filterValue = document.getElementById('filterValue');
        this.filterApplyBtn = document.getElementById('filterApplyBtn');
        this.filterClearBtn = document.getElementById('filterClearBtn');
        this.activeFiltersDiv = document.getElementById('activeFilters');
        this.filterTags = document.getElementById('filterTags');
        
        // Create new dropdown
        this.createDropdownBtn = document.getElementById('createDropdownBtn');
        this.createDropdownMenu = document.getElementById('createDropdownMenu');
        this.createSheetBtn = document.getElementById('createSheetBtn');
        this.createDocBtn = document.getElementById('createDocBtn');
        this.createNewLargeBtn = document.getElementById('createNewLargeBtn');
        this.createDocLargeBtn = document.getElementById('createDocLargeBtn');
        
        // Document editor
        this.documentContainer = document.getElementById('documentContainer');
        this.docEditor = document.getElementById('docEditor');
        this.docInfo = document.getElementById('docInfo');
        this.docToolbar = document.getElementById('docToolbar');
        this.docFileName = document.getElementById('docFileName');
        this.docFileSize = document.getElementById('docFileSize');
        this.docEditIndicator = document.getElementById('docEditIndicator');
        this.docSaveBtn = document.getElementById('docSaveBtn');
        this.docExportBtn = document.getElementById('docExportBtn');
        this.docWordCount = document.getElementById('docWordCount');
        this.docCharCount = document.getElementById('docCharCount');
        this.fontSelect = document.getElementById('fontSelect');
        this.sizeSelect = document.getElementById('sizeSelect');
        this.textColorPicker = document.getElementById('textColorPicker');
        this.bgColorPicker = document.getElementById('bgColorPicker');
        this.insertLinkBtn = document.getElementById('insertLinkBtn');
        this.insertImageBtn = document.getElementById('insertImageBtn');
        
        // Slides editor
        this.slidesContainer = document.getElementById('slidesContainer');
        this.slidesInfo = document.getElementById('slidesInfo');
        this.slidesFileName = document.getElementById('slidesFileName');
        this.slidesFileSize = document.getElementById('slidesFileSize');
        this.slidesEditIndicator = document.getElementById('slidesEditIndicator');
        this.slidesSaveBtn = document.getElementById('slidesSaveBtn');
        this.slidesExportBtn = document.getElementById('slidesExportBtn');
        this.slidesViewer = document.getElementById('slidesViewer');
        this.slideCounter = document.getElementById('slideCounter');
        this.prevSlideBtn = document.getElementById('prevSlideBtn');
        this.nextSlideBtn = document.getElementById('nextSlideBtn');
        this.slidesThumbnails = document.getElementById('slidesThumbnails');
        this.createSlidesBtn = document.getElementById('createSlidesBtn');
        this.createSlidesLargeBtn = document.getElementById('createSlidesLargeBtn');
    }
};

// Export for use in other modules
window.DOM = DOM;