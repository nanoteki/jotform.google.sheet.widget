// Google Sheets Dropdown Widget Script
console.log('Google Sheets Dropdown Widget script loaded');

// Debounce utility function
const debounce = (fn, delay) => {
    let timeoutId = null;
    return (...args) => {
        if (timeoutId) {
            clearTimeout(timeoutId);
        }
        timeoutId = setTimeout(() => {
            fn(...args);
            timeoutId = null;
        }, delay);
    };
};

class GoogleSheetsDropdownWidget {
    constructor() {
        console.log('GoogleSheetsDropdownWidget constructor called');
        this.settings = {};
        this.currentValue = '';
        this.isInitialized = false;
        this.cachedData = null;
        this.minSearchLength = 2;
        this.debounceDelay = 300;
        
        // DOM elements
        this.searchInput = document.getElementById('search-input');
        this.dropdownList = document.getElementById('dropdown-list');
        this.inputContainer = document.querySelector('.input-container');
        this.selectedValueInput = document.getElementById('selected-value');
        this.noResults = document.getElementById('no-results');
        this.loadingOptions = document.getElementById('loading-options');
        this.typeToSearch = document.getElementById('type-to-search');
        this.inputSpinner = document.getElementById('input-spinner');
        this.loading = document.getElementById('loading');
        this.errorMessage = document.getElementById('error-message');
        this.errorText = document.getElementById('error-text');
        this.retryBtn = document.getElementById('retry-btn');
        this.configError = document.getElementById('config-error');
        this.missingParams = document.getElementById('missing-params');
        this.questionLabel = document.getElementById('question-label');

        // Store all options for filtering
        this.allOptions = [];
        this.filteredOptions = [];
        this.selectedOption = null;
        this.isOpen = false;
        this.isLoading = false;

        // Create debounced search function
        this.debouncedSearch = debounce((searchTerm) => {
            this.performSearch(searchTerm);
        }, this.debounceDelay);

        this.bindEvents();
    }
    
    bindEvents() {
        // Search input events
        this.searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value;
            
            if (searchTerm.length < this.minSearchLength) {
                this.showTypeToSearch();
                return;
            }
            
            this.showSearching();
            this.debouncedSearch(searchTerm);
        });

        this.searchInput.addEventListener('focus', () => {
            this.openDropdown();
        });

        this.searchInput.addEventListener('keydown', (e) => {
            this.handleKeydown(e);
        });

        // Click outside to close
        document.addEventListener('click', (e) => {
            if (!this.inputContainer.contains(e.target) && !this.dropdownList.contains(e.target)) {
                this.closeDropdown();
            }
        });

        // Retry button
        this.retryBtn.addEventListener('click', () => {
            this.loadSheetData();
        });
    }
    
    init(formId, value) {
        console.log('Widget init called with:', { formId, value });
        this.currentValue = value || '';

        // Get settings from JotForm or URL parameters (for testing)
        this.settings = this.getSettings();
        console.log('Widget settings:', this.settings);

        // Set question label
        const questionLabel = this.settings.QuestionLabel || 'Select a property';
        this.questionLabel.textContent = questionLabel;
        
        // Set minimum search length if configured
        if (this.settings.MinSearchLength) {
            this.minSearchLength = parseInt(this.settings.MinSearchLength, 10) || 2;
        }

        // Set debounce delay if configured
        if (this.settings.DebounceDelay) {
            this.debounceDelay = parseInt(this.settings.DebounceDelay, 10) || 300;
            this.debouncedSearch = debounce((searchTerm) => {
                this.performSearch(searchTerm);
            }, this.debounceDelay);
        }
        
        // Validate configuration
        if (!this.validateConfiguration()) {
            return;
        }
        
        // Load sheet data
        this.loadSheetData();
        
        this.isInitialized = true;
    }

    getSettings() {
        // Try to get settings from JotForm first
        if (typeof JFCustomWidget !== 'undefined' && JFCustomWidget.getWidgetSettings) {
            try {
                const jfSettings = JFCustomWidget.getWidgetSettings();
                if (jfSettings && Object.keys(jfSettings).length > 0) {
                    return jfSettings;
                }
            } catch (e) {
                console.log('JotForm settings not available, checking URL parameters');
            }
        }

        // Fallback to URL parameters for testing
        const urlParams = new URLSearchParams(window.location.search);
        const settings = {};

        // Map URL parameters to settings
        const paramMap = {
            'SpreadsheetId': 'SpreadsheetId',
            'SheetName': 'SheetName',
            'ValueColumn': 'ValueColumn',
            'LabelColumn': 'LabelColumn',
            'QuestionLabel': 'QuestionLabel',
            'MinSearchLength': 'MinSearchLength',
            'DebounceDelay': 'DebounceDelay'
        };

        for (const [urlParam, settingKey] of Object.entries(paramMap)) {
            const value = urlParams.get(urlParam);
            if (value) {
                settings[settingKey] = value;
            }
        }

        return settings;
    }
    
    validateConfiguration() {
        console.log('Validating configuration...');
        const requiredParams = ['SpreadsheetId'];
        const missingParams = [];

        requiredParams.forEach(param => {
            const value = this.settings[param];
            console.log(`Checking param ${param}:`, value ? 'present' : 'missing');
            if (!value || value.trim() === '') {
                missingParams.push(param);
            }
        });

        if (missingParams.length > 0) {
            console.log('Missing required parameters:', missingParams);
            this.showConfigurationError(missingParams);
            return false;
        }

        console.log('Configuration validation passed');
        return true;
    }
    
    showConfigurationError(missingParams) {
        this.hideAllMessages();
        this.missingParams.innerHTML = '';
        
        missingParams.forEach(param => {
            const li = document.createElement('li');
            li.textContent = this.getParamDisplayName(param);
            this.missingParams.appendChild(li);
        });
        
        this.configError.classList.remove('hidden');
    }
    
    getParamDisplayName(param) {
        const displayNames = {
            'SpreadsheetId': 'Google Spreadsheet ID',
            'SheetName': 'Sheet Name (optional, defaults to first sheet)',
            'ValueColumn': 'Value Column (optional, defaults to A)',
            'LabelColumn': 'Label Column (optional, defaults to A)'
        };
        return displayNames[param] || param;
    }
    
    async loadSheetData() {
        this.showLoading();
        
        try {
            const data = await this.fetchSheetData();
            this.cachedData = data;
            this.allOptions = this.parseSheetData(data);
            console.log(`Loaded ${this.allOptions.length} options from Google Sheets`);
            this.hideAllMessages();
            
            // Restore current value if it exists
            if (this.currentValue) {
                const selectedOption = this.allOptions.find(opt => opt.value === this.currentValue);
                if (selectedOption) {
                    this.selectOption(selectedOption, false);
                }
            }
        } catch (error) {
            console.error('Error loading sheet data:', error);
            this.showError(error.message);
        }
    }
    
    async fetchSheetData() {
        const spreadsheetId = this.settings.SpreadsheetId;
        const sheetName = this.settings.SheetName || 'Sheet1';
        
        // Use Google Sheets API v4 with public access (no API key needed for public sheets)
        // Format: https://docs.google.com/spreadsheets/d/{spreadsheetId}/gviz/tq?tqx=out:json&sheet={sheetName}
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;

        console.log('Fetching Google Sheets data from:', url);

        const response = await fetch(url);

        if (!response.ok) {
            throw new Error(`Failed to fetch spreadsheet: ${response.status}`);
        }

        const text = await response.text();
        
        // Google returns JSONP-like response, need to extract JSON
        // Response format: google.visualization.Query.setResponse({...})
        const jsonMatch = text.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?$/);
        
        if (!jsonMatch) {
            throw new Error('Invalid response format from Google Sheets');
        }

        const data = JSON.parse(jsonMatch[1]);
        
        if (data.status === 'error') {
            throw new Error(data.errors?.[0]?.message || 'Error fetching spreadsheet data');
        }

        return data;
    }
    
    parseSheetData(data) {
        const valueColumnLetter = (this.settings.ValueColumn || 'A').toUpperCase();
        const labelColumnLetter = (this.settings.LabelColumn || 'A').toUpperCase();
        
        // Convert column letters to indices (A=0, B=1, etc.)
        const valueColumnIndex = valueColumnLetter.charCodeAt(0) - 65;
        const labelColumnIndex = labelColumnLetter.charCodeAt(0) - 65;

        console.log('Parsing sheet data with columns:', { 
            valueColumn: valueColumnLetter, 
            valueIndex: valueColumnIndex,
            labelColumn: labelColumnLetter, 
            labelIndex: labelColumnIndex 
        });

        const rows = data.table?.rows || [];
        const options = [];

        // Skip header row (first row)
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const cells = row.c || [];
            
            const valueCell = cells[valueColumnIndex];
            const labelCell = cells[labelColumnIndex];
            
            const value = valueCell?.v ?? valueCell?.f ?? '';
            const label = labelCell?.v ?? labelCell?.f ?? '';
            
            // Skip empty rows
            if (value === '' && label === '') continue;
            
            options.push({
                value: String(value),
                label: String(label || value)
            });
        }

        return options;
    }

    performSearch(searchTerm) {
        if (!searchTerm || searchTerm.length < this.minSearchLength) {
            this.showTypeToSearch();
            return;
        }

        const search = searchTerm.toLowerCase().trim();
        
        // Filter options based on search term
        this.filteredOptions = this.allOptions.filter(option => {
            const label = option.label ? option.label.toLowerCase() : '';
            const value = option.value ? option.value.toLowerCase() : '';
            return label.includes(search) || value.includes(search);
        });

        // Sort results alphabetically by label
        this.filteredOptions.sort((a, b) => {
            const labelA = (a.label || '').toLowerCase();
            const labelB = (b.label || '').toLowerCase();
            return labelA.localeCompare(labelB);
        });

        this.hideSearching();
        this.renderDropdownItems();
    }

    showTypeToSearch() {
        this.hideSearching();
        this.dropdownList.innerHTML = '';
        
        const typeToSearchItem = document.createElement('div');
        typeToSearchItem.className = 'type-to-search-item';
        typeToSearchItem.textContent = `Type at least ${this.minSearchLength} characters to search...`;
        this.dropdownList.appendChild(typeToSearchItem);
        
        this.openDropdown();
    }

    showSearching() {
        this.inputSpinner.classList.remove('hidden');
    }

    hideSearching() {
        this.inputSpinner.classList.add('hidden');
    }

    renderDropdownItems() {
        // Clear existing items
        this.dropdownList.innerHTML = '';

        if (this.filteredOptions.length === 0) {
            // Show no results
            const noResultsItem = document.createElement('div');
            noResultsItem.className = 'no-results-item';
            noResultsItem.textContent = 'No options match your search';
            this.dropdownList.appendChild(noResultsItem);
        } else {
            // Add filtered options
            this.filteredOptions.forEach(option => {
                const item = document.createElement('div');
                item.className = 'dropdown-item';
                item.textContent = option.label;
                item.dataset.value = option.value;

                // Highlight if selected
                if (this.selectedOption && this.selectedOption.value === option.value) {
                    item.classList.add('selected');
                }

                // Click handler
                item.addEventListener('click', () => {
                    this.selectOption(option);
                    this.closeDropdown();
                });

                this.dropdownList.appendChild(item);
            });
        }

        this.openDropdown();
    }

    selectOption(option, sendData = true) {
        this.selectedOption = option;
        this.currentValue = option.value;
        this.searchInput.value = option.label;
        this.selectedValueInput.value = option.value;
        
        if (sendData) {
            this.sendData();
        }
        
        console.log('Selected option:', option);
    }

    openDropdown() {
        this.isOpen = true;
        this.inputContainer.classList.add('open');
        this.dropdownList.classList.remove('hidden');
    }

    closeDropdown() {
        this.isOpen = false;
        this.inputContainer.classList.remove('open');
        this.dropdownList.classList.add('hidden');

        // If no option is selected and input has text, clear it
        if (!this.selectedOption && this.searchInput.value) {
            this.searchInput.value = '';
        }
    }

    handleKeydown(e) {
        if (!this.isOpen) {
            if (e.key === 'ArrowDown' || e.key === 'Enter') {
                this.openDropdown();
                e.preventDefault();
            }
            return;
        }

        const items = this.dropdownList.querySelectorAll('.dropdown-item');
        const currentSelected = this.dropdownList.querySelector('.dropdown-item.selected');
        let newIndex = -1;

        if (e.key === 'ArrowDown') {
            const currentIndex = currentSelected ? Array.from(items).indexOf(currentSelected) : -1;
            newIndex = Math.min(currentIndex + 1, items.length - 1);
            e.preventDefault();
        } else if (e.key === 'ArrowUp') {
            const currentIndex = currentSelected ? Array.from(items).indexOf(currentSelected) : items.length;
            newIndex = Math.max(currentIndex - 1, 0);
            e.preventDefault();
        } else if (e.key === 'Enter') {
            if (currentSelected) {
                const value = currentSelected.dataset.value;
                const option = this.filteredOptions.find(opt => opt.value === value);
                if (option) {
                    this.selectOption(option);
                    this.closeDropdown();
                }
            }
            e.preventDefault();
        } else if (e.key === 'Escape') {
            this.closeDropdown();
            e.preventDefault();
        }

        // Update selection highlight
        if (newIndex >= 0 && items[newIndex]) {
            items.forEach(item => item.classList.remove('selected'));
            items[newIndex].classList.add('selected');
            items[newIndex].scrollIntoView({ block: 'nearest' });
        }
    }
    
    showLoading() {
        this.hideAllMessages();
        this.loading.classList.remove('hidden');
        this.searchInput.disabled = true;
        this.searchInput.placeholder = 'Loading data...';
    }

    showError(message) {
        this.hideAllMessages();
        this.errorText.textContent = message;
        this.errorMessage.classList.remove('hidden');
        this.searchInput.disabled = false;
        this.searchInput.placeholder = 'Type to search...';
    }

    hideAllMessages() {
        this.loading.classList.add('hidden');
        this.errorMessage.classList.add('hidden');
        this.configError.classList.add('hidden');
        this.dropdownList.classList.add('hidden');
        this.searchInput.disabled = false;
        this.searchInput.placeholder = 'Type to search...';
    }
    
    sendData() {
        if (this.isInitialized) {
            if (typeof JFCustomWidget !== 'undefined') {
                JFCustomWidget.sendData({
                    value: this.currentValue
                });
            } else {
                // For testing - just log the value
                console.log('Widget value:', this.currentValue);
            }
        }
    }
    
    handleSubmit() {
        const isValid = this.currentValue !== '';

        if (typeof JFCustomWidget !== 'undefined') {
            JFCustomWidget.sendSubmit({
                valid: isValid,
                value: this.currentValue
            });
        } else {
            // For testing - just log the submission
            console.log('Widget submit:', { valid: isValid, value: this.currentValue });
        }
    }
}

// Initialize widget
console.log('Creating Google Sheets widget instance...');
const widget = new GoogleSheetsDropdownWidget();

// Subscribe to JotForm events or initialize directly for testing
if (typeof JFCustomWidget !== 'undefined') {
    console.log('JFCustomWidget available, subscribing to events');

    JFCustomWidget.subscribe("ready", function(formId, value) {
        console.log('JFCustomWidget ready event fired');
        widget.init(formId, value);
    });

    JFCustomWidget.subscribe("submit", function() {
        console.log('JFCustomWidget submit event fired');
        widget.handleSubmit();
    });

    // For direct testing - if we have URL parameters, initialize directly
    document.addEventListener('DOMContentLoaded', function() {
        console.log('DOMContentLoaded fired, checking for URL parameters');
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.has('SpreadsheetId')) {
            console.log('URL parameters detected, initializing widget directly');
            setTimeout(() => {
                widget.init('test-form', '');
            }, 100);
        }
    });
} else {
    console.log('JFCustomWidget not available, using DOMContentLoaded');
    document.addEventListener('DOMContentLoaded', function() {
        console.log('DOMContentLoaded fired, initializing widget');
        widget.init('test-form', '');
    });
}

