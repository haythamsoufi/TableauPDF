// Initialize Alpine.js data store
document.addEventListener('alpine:init', () => {
    Alpine.data('appState', () => ({
        // --- Constants ---
        STORAGE_KEY: 'tableauExporterConfigs', // Key for localStorage

        // --- Version ---
        current_version: 'v1.11-Web-LS-v2', // LS for LocalStorage

        // --- Main Configuration State (Mirrors form fields) ---
        config: {
            server_url: '',
            site_id: '',
            token_name: '',
            token_secret: '', // Will be loaded/set from modal/localStorage
            workbook_name: '',
            export_mode: 'automate',
            excel_filepath: '', // Server path to uploaded file
            sheet_name: '',
            tableau_filter_field: '',
            file_naming_option: 'By view',
            organize_by_1: 'None',
            organize_by_2: 'None',
            export_format: 'PDF',
            numbering_enabled: true,
            excluded_views: [],
            filters: [],
            conditions: [],
            parameters: []
        },

        // --- Server Config Modal State ---
        modalConfig: { // Holds temporary values for the modal inputs
            server_url: '',
            site_id: '',
            token_name: '',
            token_secret: '',
            workbook_name: ''
        },
        isTestingConnection: false, // Loading state for modal test button
        modalTestStatus: '',      // Feedback message inside modal
        modalTestError: false,    // Flag for styling modal feedback

        // --- Configuration Management State ---
        savedConfigs: {}, // In-memory store of saved configs {name: configObj}
        savedConfigNames: [], // Array of names for the dropdown
        configNameInput: '', // Bound to the input for naming configs
        selectedConfigToLoad: '', // Bound to the dropdown selection

        // --- UI & Loading State ---
        flashMessage: { text: '', type: 'info' },
        globalError: '',
        isLoadingViews: false, // For the main "Load Views" button
        viewLoadStatus: 'Configure server and workbook, then click Load Views.',
        viewLoadError: false,
        availableViews: [],
        excelUploadStatus: 'No file selected.',
        excelUploadError: false,
        availableSheets: [],
        availableColumns: [],
        isExporting: false,
        exportProgress: 0,
        exportStatusMessage: 'Idle',
        logMessages: [],
        pollingInterval: null,
        currentTaskId: null,

        // --- Methods ---

        // --- Initialization ---
        init() {
            console.log('Alpine app initializing...');
            this.loadSavedConfigNames();
            // Load 'Default' or last used config if desired
            // this.loadConfig('Default');
            this.syncModalWithCurrentConfig(); // Initialize modal with current (likely empty) config
            // Add listener to sync modal state when opened
            const modalElement = document.getElementById('serverConfigModal');
            if(modalElement) {
                modalElement.addEventListener('show.bs.modal', () => {
                    this.syncModalWithCurrentConfig();
                    // Clear previous test status when opening
                    this.modalTestStatus = '';
                    this.modalTestError = false;
                });
            }
        },

        // --- Server Config Modal ---
        syncModalWithCurrentConfig() {
            // Copy current main config state TO modal state
            this.modalConfig.server_url = this.config.server_url;
            this.modalConfig.site_id = this.config.site_id;
            this.modalConfig.token_name = this.config.token_name;
            this.modalConfig.workbook_name = this.config.workbook_name;
            // IMPORTANT: Also copy the current secret TO the modal when opening
            this.modalConfig.token_secret = this.config.token_secret;
        },

        saveServerConfigFromModal() {
            console.log('Applying server config from modal');
            // Apply changes from modal state back to main config
            this.config.server_url = this.modalConfig.server_url?.trim() || '';
            this.config.site_id = this.modalConfig.site_id?.trim() || '';
            this.config.token_name = this.modalConfig.token_name?.trim() || '';
            this.config.workbook_name = this.modalConfig.workbook_name?.trim() || '';
            // Apply the secret from the modal to the main config
            this.config.token_secret = this.modalConfig.token_secret || ''; // Don't trim secret

            // Reset view loading status as server details might have changed
            this.availableViews = [];
            this.config.excluded_views = []; // Also reset exclusions tied to views
            this.viewLoadStatus = 'Server settings applied. Click Load Views.';
            this.viewLoadError = false;

            this.showFlashMessage('Server settings applied.', 'success');
            // Note: We don't automatically save to localStorage here, user must use Save button
        },

        async testServerConnection() {
            console.log('Testing server connection from modal...');
            this.isTestingConnection = true;
            this.modalTestStatus = 'Testing...';
            this.modalTestError = false;

            // Basic client-side checks
            if (!this.modalConfig.server_url || !this.modalConfig.token_name || !this.modalConfig.token_secret) {
                this.modalTestStatus = 'Error: URL, PAT Name, and Secret required.';
                this.modalTestError = true;
                this.isTestingConnection = false;
                return;
            }

            try {
                // *** IMPORTANT: This requires a new backend endpoint '/test_connection' ***
                const response = await fetch('/test_connection', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                    body: JSON.stringify({ // Send details currently in the modal
                        server_url: this.modalConfig.server_url,
                        token_name: this.modalConfig.token_name,
                        token_secret: this.modalConfig.token_secret,
                        site_id: this.modalConfig.site_id
                    })
                });
                const data = await response.json();
                if (!response.ok || !data.success) {
                    throw new Error(data.error || `Connection test failed (HTTP ${response.status})`);
                }
                this.modalTestStatus = 'Connection Successful!';
                this.modalTestError = false;
            } catch (error) {
                console.error('Error testing connection:', error);
                this.modalTestStatus = `Connection Failed: ${error.message}`;
                this.modalTestError = true;
            } finally {
                this.isTestingConnection = false;
            }
        },

        // --- Configuration Management (LocalStorage) ---
        loadSavedConfigNames() {
            try {
                const stored = localStorage.getItem(this.STORAGE_KEY);
                this.savedConfigs = stored ? JSON.parse(stored) : {};
                this.savedConfigNames = Object.keys(this.savedConfigs).sort();
                console.log('Loaded saved config names:', this.savedConfigNames);
            } catch (e) {
                console.error("Error loading config names from localStorage:", e);
                this.showFlashMessage("Error loading saved configurations.", "error");
                this.savedConfigs = {};
                this.savedConfigNames = [];
                localStorage.removeItem(this.STORAGE_KEY);
            }
        },

        saveConfigToLocalStorage(name, configToSave) {
            if (!name) return;
            try {
                // Create a deep copy to avoid modifying the original object
                const configCopy = JSON.parse(JSON.stringify(configToSave));
                // *** We are NOW saving the secret ***
                // delete configCopy.token_secret; // NO LONGER DELETING
                delete configCopy.excel_filepath; // Still don't save server file path

                this.savedConfigs[name] = configCopy;
                localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.savedConfigs));
                this.loadSavedConfigNames(); // Refresh names list
                this.showFlashMessage(`Configuration '${this.escapeHtml(name)}' saved successfully (including PAT Secret).`, 'success');
            } catch (e) {
                console.error(`Error saving config '${name}' to localStorage:`, e);
                this.showFlashMessage(`Error saving configuration '${this.escapeHtml(name)}'. Check browser storage limits.`, 'error');
            }
        },

        loadConfigFromLocalStorage(name) {
            if (!name || !this.savedConfigs[name]) {
                this.showFlashMessage(`Configuration '${this.escapeHtml(name)}' not found.`, 'warning');
                return null;
            }
            try {
                const loadedConfig = JSON.parse(JSON.stringify(this.savedConfigs[name]));
                 // Ensure all expected fields exist, providing defaults if missing from save
                 loadedConfig.server_url = loadedConfig.server_url || '';
                 loadedConfig.site_id = loadedConfig.site_id || '';
                 loadedConfig.token_name = loadedConfig.token_name || '';
                 loadedConfig.token_secret = loadedConfig.token_secret || ''; // *** Load the saved secret ***
                 loadedConfig.workbook_name = loadedConfig.workbook_name || '';
                 loadedConfig.export_mode = loadedConfig.export_mode || 'automate';
                 loadedConfig.excel_filepath = ''; // Clear file path
                 loadedConfig.sheet_name = loadedConfig.sheet_name || '';
                 loadedConfig.tableau_filter_field = loadedConfig.tableau_filter_field || '';
                 loadedConfig.file_naming_option = loadedConfig.file_naming_option || 'By view';
                 loadedConfig.organize_by_1 = loadedConfig.organize_by_1 || 'None';
                 loadedConfig.organize_by_2 = loadedConfig.organize_by_2 || 'None';
                 loadedConfig.export_format = loadedConfig.export_format || 'PDF';
                 loadedConfig.numbering_enabled = loadedConfig.numbering_enabled ?? true; // Default to true if missing
                 loadedConfig.excluded_views = loadedConfig.excluded_views || [];
                 loadedConfig.filters = loadedConfig.filters || [];
                 loadedConfig.conditions = loadedConfig.conditions || [];
                 loadedConfig.parameters = loadedConfig.parameters || [];

                 return loadedConfig;
            } catch (e) {
                 console.error(`Error loading/parsing config '${name}' from memory:`, e);
                 this.showFlashMessage(`Error loading configuration '${this.escapeHtml(name)}'. It might be corrupted.`, 'error');
                 return null;
            }
        },

        deleteConfigFromLocalStorage(name) {
            if (!name || !this.savedConfigs[name]) return;
            try {
                delete this.savedConfigs[name];
                localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.savedConfigs));
                this.loadSavedConfigNames();
                this.showFlashMessage(`Configuration '${this.escapeHtml(name)}' deleted.`, 'success');
                if (this.configNameInput === name) this.configNameInput = '';
                if (this.selectedConfigToLoad === name) this.selectedConfigToLoad = '';
            } catch (e) {
                console.error(`Error deleting config '${name}' from localStorage:`, e);
                this.showFlashMessage(`Error deleting configuration '${this.escapeHtml(name)}'.`, 'error');
            }
        },

        // --- Button Actions for Config Management ---
        saveCurrentConfig() {
            const name = this.configNameInput.trim();
            if (!name) {
                this.showFlashMessage('Please enter a name for the configuration.', 'warning');
                return;
            }
            if (name.length > 50) {
                this.showFlashMessage('Configuration name is too long (max 50 chars).', 'warning');
                return;
            }
            if (this.savedConfigNames.includes(name)) {
                if (!confirm(`Configuration '${this.escapeHtml(name)}' already exists. Overwrite?`)) {
                    return;
                }
            }
            // Ensure the current live secret is included in the config object being saved
            this.config.token_secret = this.modalConfig.token_secret || this.config.token_secret || '';
            this.saveConfigToLocalStorage(name, this.config);
            this.selectedConfigToLoad = name; // Select the newly saved config in dropdown
        },

        loadConfig(name) {
            const nameToLoad = name || this.selectedConfigToLoad;
            if (!nameToLoad) {
                this.showFlashMessage('Please select or enter a configuration name to load.', 'warning');
                return;
            }
            const loaded = this.loadConfigFromLocalStorage(nameToLoad);
            if (loaded) {
                this.config = loaded; // Replace main config
                this.resetTransientStateAfterLoad();
                this.configNameInput = nameToLoad;
                this.selectedConfigToLoad = nameToLoad;
                this.syncModalWithCurrentConfig(); // Update modal fields with loaded config (including secret)
                this.showFlashMessage(`Configuration '${this.escapeHtml(nameToLoad)}' loaded.`, 'success');
                this.globalError = '';
            }
        },

        deleteConfig(name) {
            const nameToDelete = name || this.selectedConfigToLoad;
            if (!nameToDelete) {
                this.showFlashMessage('Please select a configuration to delete.', 'warning');
                return;
            }
            if (confirm(`Are you sure you want to delete configuration '${this.escapeHtml(nameToDelete)}'? This cannot be undone.`)) {
                this.deleteConfigFromLocalStorage(nameToDelete);
            }
        },

        resetTransientStateAfterLoad() {
            // Reset things that depend on external data or actions
            this.availableViews = [];
            this.viewLoadStatus = 'Config loaded. Click Load Views if needed.';
            this.viewLoadError = false;
            this.excelUploadStatus = 'No file selected (load file if needed).';
            this.excelUploadError = false;
            this.availableSheets = [];
            this.availableColumns = [];
            // config.excel_filepath is already cleared in loadConfigFromLocalStorage
            this.isExporting = false;
            this.exportProgress = 0;
            this.exportStatusMessage = 'Idle';
            this.logMessages = [];
            this.clearPolling();
            this.currentTaskId = null;
            this.globalError = '';
            this.flashMessage = { text: '', type: 'info' };
            const fileInput = document.getElementById('excel_file_input');
            if(fileInput) fileInput.value = '';
        },

        // --- Tableau Interaction (Main Page Load Views) ---
        async loadViews() {
            this.isLoadingViews = true;
            this.viewLoadStatus = 'Connecting to Tableau...';
            this.viewLoadError = false;
            this.availableViews = [];
            this.globalError = '';

            // Validate using the main config object now
            if (!this.config.server_url || !this.config.token_name || !this.config.workbook_name) {
                this.viewLoadStatus = 'Error: Server URL, PAT Name, and Workbook Name required (Configure Server).';
                this.viewLoadError = true;
                this.isLoadingViews = false;
                return;
            }
             if (!this.config.token_secret) { // Check the main config secret
                this.viewLoadStatus = 'Error: PAT Secret is missing. Please set it via Configure Server.';
                this.viewLoadError = true;
                this.isLoadingViews = false;
                 // Optionally open the modal automatically
                 // bootstrap.Modal.getOrCreateInstance(document.getElementById('serverConfigModal')).show();
                return;
            }

            try {
                const response = await fetch('/load_views', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                    body: JSON.stringify({ // Send details from main config
                        server_url: this.config.server_url,
                        token_name: this.config.token_name,
                        token_secret: this.config.token_secret, // Send the live secret
                        site_id: this.config.site_id,
                        workbook_name: this.config.workbook_name
                    })
                });
                const data = await response.json();
                if (!response.ok) {
                    throw new Error(data.error || `HTTP error! status: ${response.status}`);
                }
                this.availableViews = data.views || [];
                this.viewLoadStatus = this.availableViews.length > 0 ? `Successfully loaded ${this.availableViews.length} views.` : 'Workbook found, but it contains no views.';
                this.viewLoadError = false;
                 // Clear status message after success
                 setTimeout(() => { if(!this.viewLoadError) this.viewLoadStatus = ''; }, 5000);
            } catch (error) {
                console.error('Error loading views:', error);
                this.viewLoadStatus = `Error loading views: ${error.message}. Check details & server status.`;
                this.viewLoadError = true;
            } finally {
                this.isLoadingViews = false;
            }
        },

        // --- Excel Interaction (Unchanged from previous) ---
        async handleExcelUpload(event) { /* ... keep previous implementation ... */
            const file = event.target.files[0];
            if (!file) {
                this.excelUploadStatus = 'No file selected.';
                this.excelUploadError = false;
                return;
            }
            this.excelUploadStatus = 'Uploading & Processing...';
            this.excelUploadError = false;
            this.globalError = '';
            this.availableSheets = [];
            this.availableColumns = [];
            this.config.excel_filepath = '';
            this.config.sheet_name = '';
            this.config.tableau_filter_field = '';
            this.config.file_naming_option = 'By view';
            this.config.organize_by_1 = 'None';
            this.config.organize_by_2 = 'None';
            this.config.filters = [];
            this.config.conditions = [];
            const formData = new FormData();
            formData.append('excel_file', file);
            try {
                const response = await fetch('/upload_excel', {
                    method: 'POST',
                    body: formData,
                    headers: { 'Accept': 'application/json' },
                });
                const data = await response.json();
                 if (!response.ok) {
                    throw new Error(data.error || `HTTP error! status: ${response.status}`);
                }
                this.excelUploadStatus = `File '${this.escapeHtml(file.name)}' processed.`;
                this.excelUploadError = false;
                this.availableSheets = data.sheets || [];
                this.availableColumns = data.columns || [];
                this.config.excel_filepath = data.filepath;
                if (this.availableSheets.length > 0) {
                    this.config.sheet_name = this.availableSheets[0];
                } else {
                     this.excelUploadStatus += ' Warning: No sheets found.';
                }
                 if (this.availableColumns.length === 0 && this.availableSheets.length > 0) {
                     this.excelUploadStatus += ` Warning: Sheet '${this.config.sheet_name}' has no columns/header.`;
                }
                 if (this.availableColumns.length > 0){
                     setTimeout(() => { if(this.excelUploadStatus.includes('processed')) this.excelUploadStatus = ''; }, 5000);
                 }
            } catch (error) {
                console.error('Error uploading Excel:', error);
                this.excelUploadStatus = `Upload Error: ${error.message}`;
                this.excelUploadError = true;
                this.resetAutomateFields();
                event.target.value = '';
            }
         },
        async loadColumnsForSheet() { /* ... keep previous implementation ... */
             if (!this.config.excel_filepath || !this.config.sheet_name) {
                 this.availableColumns = [];
                 return;
             }
             console.log(`Requesting columns for sheet: ${this.config.sheet_name}`);
             this.availableColumns = [];
             this.config.tableau_filter_field = '';
             this.config.file_naming_option = 'By view';
             this.config.organize_by_1 = 'None';
             this.config.organize_by_2 = 'None';
             this.config.filters = [];
             this.config.conditions = [];
             try {
                 const response = await fetch('/get_columns', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                    body: JSON.stringify({
                        filepath: this.config.excel_filepath,
                        sheet_name: this.config.sheet_name
                    })
                 });
                 const data = await response.json();
                  if (!response.ok) {
                    throw new Error(data.error || `HTTP error! status: ${response.status}`);
                 }
                 this.availableColumns = data.columns || [];
                 console.log(`Loaded columns: ${this.availableColumns.join(', ')}`);
                  if (this.availableColumns.length === 0){
                     this.showFlashMessage(`Warning: Sheet '${this.escapeHtml(this.config.sheet_name)}' appears empty or has no header row. Dependent options disabled.`, 'warning');
                 }
             } catch (error) {
                 console.error('Error loading columns:', error);
                 this.showFlashMessage(`Error loading columns for sheet '${this.escapeHtml(this.config.sheet_name)}': ${error.message}`, 'danger');
             }
        },
        handleModeChange() { /* ... keep previous implementation ... */
            console.log("Mode changed to:", this.config.export_mode);
            if (this.config.export_mode === 'all_once') {
                this.resetAutomateFields();
                const fileInput = document.getElementById('excel_file_input');
                if(fileInput) fileInput.value = '';
            }
        },
        resetAutomateFields() { /* ... keep previous implementation ... */
            console.log('Resetting automate fields');
            this.availableSheets = [];
            this.availableColumns = [];
            this.config.excel_filepath = '';
            this.config.sheet_name = '';
            this.config.tableau_filter_field = '';
            this.config.file_naming_option = 'By view';
            this.config.organize_by_1 = 'None';
            this.config.organize_by_2 = 'None';
            this.config.filters = [];
            this.config.conditions = [];
        },

        // --- Dynamic Row Methods (Unchanged) ---
        addFilter() { this.config.filters.push({ field: '', values_str: '' }); },
        removeFilter(index) { this.config.filters.splice(index, 1); },
        addCondition() { this.config.conditions.push({ field: '', type: 'Equals', value: '', excluded_views_str: '' }); },
        removeCondition(index) { this.config.conditions.splice(index, 1); },
        addParameter() { this.config.parameters.push({ name: '', value: '' }); },
        removeParameter(index) { this.config.parameters.splice(index, 1); },

        // --- Export Process (Modified Validation) ---
        validateConfig() { /* ... keep previous implementation ... */
            this.globalError = '';
            const errors = [];
            if (!this.config.server_url) errors.push("Tableau Server URL is required (Configure Server).");
            if (!this.config.token_name) errors.push("Tableau PAT Name is required (Configure Server).");
            if (!this.config.token_secret) errors.push("Tableau PAT Secret is required (Configure Server modal).");
            if (!this.config.workbook_name) errors.push("Tableau Workbook Name is required (Configure Server).");
             if (this.config.export_mode === 'automate') {
                 if (!this.config.excel_filepath) errors.push("Source Excel File must be uploaded for automate mode.");
                 if (!this.config.sheet_name) errors.push("Sheet Name must be selected for automate mode.");
                 if (this.config.excel_filepath && this.availableColumns.length === 0 && this.availableSheets.length > 0) errors.push("Selected sheet appears to have no columns or header row.");
                 if(this.availableColumns.length > 0) {
                     if (this.config.tableau_filter_field && !this.availableColumns.includes(this.config.tableau_filter_field)) errors.push("Selected Key Field is not a valid column in the current sheet.");
                     if (this.config.file_naming_option !== 'By view' && !this.availableColumns.includes(this.config.file_naming_option)) errors.push("Selected File Naming column is not valid.");
                     if (this.config.organize_by_1 !== 'None' && !this.availableColumns.includes(this.config.organize_by_1)) errors.push("Selected Organize By 1 column is not valid.");
                     if (this.config.organize_by_2 !== 'None' && !this.availableColumns.includes(this.config.organize_by_2)) errors.push("Selected Organize By 2 column is not valid.");
                 }
             }
             if (this.config.export_mode === 'automate' && this.availableColumns.length > 0) {
                this.config.filters.forEach((f, i) => {
                    if (!f.field) errors.push(`Filter #${i + 1}: Field cannot be empty.`);
                    else if (!this.availableColumns.includes(f.field)) errors.push(`Filter #${i + 1}: Selected field '${f.field}' is not valid for the current sheet.`);
                    if (!f.values_str) errors.push(`Filter #${i + 1}: Values cannot be empty (UI needed).`);
                });
                this.config.conditions.forEach((c, i) => {
                    if (!c.field) errors.push(`Condition #${i + 1}: Field cannot be empty.`);
                    else if (!this.availableColumns.includes(c.field)) errors.push(`Condition #${i + 1}: Selected field '${c.field}' is not valid for the current sheet.`);
                    if (!c.type.includes('Blank') && c.value === '') errors.push(`Condition #${i + 1}: Value cannot be empty for type '${c.type}'.`);
                    if (!c.excluded_views_str) errors.push(`Condition #${i + 1}: Excluded views cannot be empty (UI needed).`);
                });
                 this.config.parameters.forEach((p, i) => {
                    if (!p.name) errors.push(`Parameter #${i + 1}: Parameter Name cannot be empty.`);
                    if (p.value === '') errors.push(`Parameter #${i + 1}: Value cannot be empty.`);
                });
            }
            if (errors.length > 0) {
                this.globalError = "<strong>Please fix the following errors:</strong><br>- " + errors.join("<br>- ");
                window.scrollTo(0, 0);
                return false;
            }
            return true;
         },
        async submitExportForm() { /* ... keep previous implementation ... */
            if (!this.validateConfig()) return;
            this.isExporting = true;
            this.exportProgress = 0;
            this.exportStatusMessage = 'Initializing export...';
            this.logMessages = ['[Client] Validated config. Sending export request...'];
            this.clearPolling();
            this.currentTaskId = null;
            this.globalError = '';
            this.flashMessage = { text: '', type: 'info' };
            const configToSend = JSON.parse(JSON.stringify(this.config));
            // Secret is already in configToSend from main config object
            console.log("Submitting export request with config:", configToSend); // Secret will be logged here! Be careful in production.
            try {
                 const response = await fetch('/start_export', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                    body: JSON.stringify(configToSend)
                 });
                 const data = await response.json();
                 if (!response.ok) {
                    throw new Error(data.error || `Export initiation failed (HTTP ${response.status})`);
                 }
                 this.logMessages.push(`[Server] ${data.message || 'Export request received.'}`);
                 this.exportStatusMessage = "Export initiated, monitoring progress...";
                 if (data.task_id) {
                     this.currentTaskId = data.task_id;
                     this.startPolling();
                 } else {
                     this.exportStatusMessage = "Export finished (simulation/no task ID).";
                     this.logMessages.push("[Client] No task ID returned from server.");
                     this.isExporting = false;
                 }
            } catch (error) {
                 console.error('Error starting export:', error);
                 this.exportStatusMessage = `Error: ${error.message}`;
                 this.logMessages.push(`[Client] Error starting export: ${error.message}`);
                 this.isExporting = false;
                 this.globalError = `Failed to start export: ${error.message}`;
                 window.scrollTo(0, 0);
            }
        },

        // --- Polling and Stop Methods (Unchanged from previous version) ---
        startPolling() { /* ... keep previous implementation ... */
            this.clearPolling();
            console.log(`Starting polling for task: ${this.currentTaskId}`);
            this.pollingInterval = setInterval(async () => {
                if (!this.currentTaskId || !this.isExporting) {
                    console.log('Stopping polling condition met.');
                    this.clearPolling();
                    return;
                }
                try {
                    const timestamp = new Date().getTime();
                    const response = await fetch(`/export_status/${this.currentTaskId}?t=${timestamp}`,{
                         headers: { 'Accept': 'application/json' },
                    });
                    if (!response.ok) {
                         console.warn(`Polling error: ${response.status}`);
                         if (response.status === 404) {
                             this.exportStatusMessage = `Error: Task ID ${this.currentTaskId} not found. Stopping polling.`;
                             this.isExporting = false;
                             this.clearPolling();
                         }
                         return;
                    }
                    const data = await response.json();
                    console.log('Poll status:', data);
                    this.exportProgress = data.progress ?? this.exportProgress;
                    if(data.log && Array.isArray(data.log)) {
                        let addedMessage = false;
                        data.log.forEach(msg => {
                            const serverMsg = `[Server] ${msg}`;
                            if (!this.logMessages.slice(-5).includes(serverMsg)) {
                                 this.logMessages.push(serverMsg);
                                 addedMessage = true;
                            }
                        });
                        const maxClientLogLines = 150;
                         if (this.logMessages.length > maxClientLogLines) {
                             this.logMessages = this.logMessages.slice(-maxClientLogLines);
                         }
                        if (addedMessage) {
                             this.$nextTick(() => {
                                 const logDiv = document.getElementById('log-output');
                                 if (logDiv) logDiv.scrollTop = logDiv.scrollHeight;
                             });
                        }
                    }
                    if (data.status === 'SUCCESS' || data.status === 'FAILURE') {
                        this.exportStatusMessage = data.status === 'SUCCESS' ? "Export Completed Successfully" : `Export Failed: ${data.error || 'Unknown error'}`;
                        if(data.status === 'FAILURE') {
                             this.logMessages.push(`[Server] ERROR: ${data.error || 'Unknown error'}`);
                        }
                        this.isExporting = false;
                        this.clearPolling();
                        this.exportProgress = (data.status === 'SUCCESS') ? 100 : (data.progress || 0);
                    } else if (data.status === 'PROGRESS') {
                         this.exportStatusMessage = `Exporting... (${this.exportProgress}%)`;
                    } else {
                         this.exportStatusMessage = data.status ? `Status: ${data.status}` : 'Waiting...';
                    }
                } catch (error) {
                    console.error("Polling failed:", error);
                    this.exportStatusMessage = "Status check failed (Network Error?).";
                }
            }, 4000);
         },
        clearPolling() { /* ... keep previous implementation ... */
            if (this.pollingInterval) {
                console.log('Clearing polling interval.');
                clearInterval(this.pollingInterval);
                this.pollingInterval = null;
            }
         },
        stopExport() { /* ... keep previous implementation ... */
            console.warn("Stop functionality requires backend implementation.");
            this.showFlashMessage("Stop functionality is not fully implemented yet. Stopping client-side monitoring.", "warning");
            this.isExporting = false;
            this.clearPolling();
            this.exportStatusMessage = "Export stopped by user (client-side only).";
            this.logMessages.push("[Client] Stop requested. Client monitoring halted.");
         },

        // --- Utility Methods ---
        showFlashMessage(message, type = 'info', duration = 5000) { /* ... keep previous implementation ... */
             this.flashMessage = { text: message, type: type };
             console.log(`Flash (${type}): ${message}`);
             setTimeout(() => {
                 if (this.flashMessage.text === message) {
                     this.flashMessage = { text: '', type: 'info' };
                 }
             }, duration);
         },
        escapeHtml(unsafe) { /* ... keep previous implementation ... */
            if (typeof unsafe !== 'string') return '';
            return unsafe
                 .replace(/&/g, "&amp;")
                 .replace(/</g, "&lt;")
                 .replace(/>/g, "&gt;")
                 .replace(/"/g, "&quot;")
                 .replace(/'/g, "&#039;");
         }

    }));
});