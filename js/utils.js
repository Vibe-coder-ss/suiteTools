/**
 * Utility Functions
 * Common helper functions used across the application
 */

const Utils = {
    // Format file size for display
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },

    // Escape HTML special characters
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    // Escape regex special characters
    escapeRegex(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    },

    // Show toast notification
    showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast-notification ${type}`;
        toast.innerHTML = message;
        document.body.appendChild(toast);
        
        setTimeout(() => toast.classList.add('show'), 10);
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => document.body.removeChild(toast), 300);
        }, 3000);
    },

    // Show export success notification
    showExportSuccess(filename) {
        const toast = document.createElement('div');
        toast.className = 'toast-notification';
        toast.innerHTML = `✅ Exported: <strong>${filename}</strong>`;
        document.body.appendChild(toast);
        
        setTimeout(() => toast.classList.add('show'), 10);
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => document.body.removeChild(toast), 300);
        }, 3000);
    },

    // Download content as blob
    downloadBlob(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        this.showExportSuccess(filename);
    },

    // Download blob file
    downloadBlobFile(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        this.showExportSuccess(filename);
    },

    // Check if arrays are equal
    arraysEqual(a, b) {
        if (!a || !b) return false;
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (a[i] !== b[i]) return false;
        }
        return true;
    },

    // Detect CSV delimiter
    detectDelimiter(line) {
        const delimiters = [',', ';', '\t', '|'];
        let maxCount = 0;
        let bestDelimiter = ',';

        delimiters.forEach(d => {
            const count = (line.match(new RegExp(this.escapeRegex(d), 'g')) || []).length;
            if (count > maxCount) {
                maxCount = count;
                bestDelimiter = d;
            }
        });

        return bestDelimiter;
    },

    // Parse CSV line handling quoted values
    parseCSVLine(line, delimiter) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === delimiter && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current.trim());

        return result;
    },

    // Custom confirm dialog (returns Promise)
    // Returns: true (OK/confirm), false (cancel/discard), 'third' (third button), null (overlay/escape)
    confirm(message, options = {}) {
        return new Promise((resolve) => {
            const modal = document.getElementById('confirmModal');
            const titleEl = document.getElementById('confirmModalTitle');
            const messageEl = document.getElementById('confirmModalMessage');
            const iconEl = document.getElementById('confirmModalIcon');
            const okBtn = document.getElementById('confirmModalOk');
            const cancelBtn = document.getElementById('confirmModalCancel');
            const thirdBtn = document.getElementById('confirmModalThird');
            const modalContent = modal.querySelector('.confirm-modal');
            
            // Set content
            titleEl.textContent = options.title || 'Confirm';
            messageEl.textContent = message;
            iconEl.textContent = options.icon || '⚠️';
            okBtn.textContent = options.confirmText || options.okText || 'OK';
            cancelBtn.textContent = options.cancelText || 'Cancel';
            
            // Set button style
            okBtn.className = 'btn-confirm';
            if (options.danger) {
                okBtn.classList.add('danger');
            } else if (options.success) {
                okBtn.classList.add('success');
            }
            
            // Handle cancel button style for three-option dialogs
            cancelBtn.className = 'btn-cancel';
            if (options.showThirdOption) {
                cancelBtn.classList.add('warning');
            }
            
            // Handle third button (e.g., "Cancel" for a Save/Discard/Cancel dialog)
            if (options.showThirdOption && options.thirdOptionText) {
                thirdBtn.style.display = 'block';
                thirdBtn.textContent = options.thirdOptionText;
            } else {
                thirdBtn.style.display = 'none';
            }
            
            // Remove alert-only class
            modalContent.classList.remove('alert-only');
            
            // Show modal
            modal.classList.add('show');
            okBtn.focus();
            
            // Cleanup function
            const cleanup = () => {
                modal.classList.remove('show');
                thirdBtn.style.display = 'none';
                cancelBtn.className = 'btn-cancel';
                okBtn.removeEventListener('click', handleOk);
                cancelBtn.removeEventListener('click', handleCancel);
                thirdBtn.removeEventListener('click', handleThird);
                document.removeEventListener('keydown', handleKeydown);
            };
            
            // Handlers
            const handleOk = () => {
                cleanup();
                resolve(true);
            };
            
            const handleCancel = () => {
                cleanup();
                resolve(false);
            };
            
            const handleThird = () => {
                cleanup();
                resolve('third');
            };
            
            const handleKeydown = (e) => {
                if (e.key === 'Escape') {
                    cleanup();
                    resolve(options.showThirdOption ? 'third' : null);
                } else if (e.key === 'Enter') {
                    handleOk();
                }
            };
            
            // Attach listeners
            okBtn.addEventListener('click', handleOk);
            cancelBtn.addEventListener('click', handleCancel);
            thirdBtn.addEventListener('click', handleThird);
            document.addEventListener('keydown', handleKeydown);
            
            // Close on overlay click
            const handleOverlayClick = (e) => {
                if (e.target === modal) {
                    cleanup();
                    resolve(options.showThirdOption ? 'third' : null);
                }
            };
            modal.addEventListener('click', handleOverlayClick, { once: true });
        });
    },

    // Custom alert dialog (returns Promise)
    alert(message, options = {}) {
        return new Promise((resolve) => {
            const modal = document.getElementById('confirmModal');
            const titleEl = document.getElementById('confirmModalTitle');
            const messageEl = document.getElementById('confirmModalMessage');
            const iconEl = document.getElementById('confirmModalIcon');
            const okBtn = document.getElementById('confirmModalOk');
            const modalContent = modal.querySelector('.confirm-modal');
            
            // Set content
            titleEl.textContent = options.title || 'Alert';
            messageEl.textContent = message;
            iconEl.textContent = options.icon || 'ℹ️';
            okBtn.textContent = options.okText || 'OK';
            
            // Set button style
            okBtn.className = 'btn-confirm';
            if (options.danger) {
                okBtn.classList.add('danger');
            } else if (options.success) {
                okBtn.classList.add('success');
            }
            
            // Add alert-only class (hides cancel button)
            modalContent.classList.add('alert-only');
            
            // Show modal
            modal.classList.add('show');
            okBtn.focus();
            
            // Cleanup function
            const cleanup = () => {
                modal.classList.remove('show');
                okBtn.removeEventListener('click', handleOk);
                document.removeEventListener('keydown', handleKeydown);
            };
            
            // Handlers
            const handleOk = () => {
                cleanup();
                resolve();
            };
            
            const handleKeydown = (e) => {
                if (e.key === 'Escape' || e.key === 'Enter') {
                    handleOk();
                }
            };
            
            // Attach listeners
            okBtn.addEventListener('click', handleOk);
            document.addEventListener('keydown', handleKeydown);
            
            // Close on overlay click
            modal.addEventListener('click', (e) => {
                if (e.target === modal) handleOk();
            }, { once: true });
        });
    },

    // Custom prompt dialog (returns Promise)
    prompt(message, defaultValue = '', options = {}) {
        return new Promise((resolve) => {
            const modal = document.getElementById('promptModal');
            const titleEl = document.getElementById('promptModalTitle');
            const messageEl = document.getElementById('promptModalMessage');
            const iconEl = document.getElementById('promptModalIcon');
            const inputEl = document.getElementById('promptModalInput');
            const okBtn = document.getElementById('promptModalOk');
            const cancelBtn = document.getElementById('promptModalCancel');
            
            // Set content
            titleEl.textContent = options.title || 'Input';
            messageEl.textContent = message;
            iconEl.textContent = options.icon || '✏️';
            okBtn.textContent = options.okText || 'OK';
            cancelBtn.textContent = options.cancelText || 'Cancel';
            inputEl.value = defaultValue;
            inputEl.placeholder = options.placeholder || '';
            
            // Show modal
            modal.classList.add('show');
            
            // Focus and select input
            setTimeout(() => {
                inputEl.focus();
                inputEl.select();
            }, 100);
            
            // Cleanup function
            const cleanup = () => {
                modal.classList.remove('show');
                okBtn.removeEventListener('click', handleOk);
                cancelBtn.removeEventListener('click', handleCancel);
                inputEl.removeEventListener('keydown', handleInputKeydown);
                document.removeEventListener('keydown', handleKeydown);
            };
            
            // Handlers
            const handleOk = () => {
                const value = inputEl.value;
                cleanup();
                resolve(value);
            };
            
            const handleCancel = () => {
                cleanup();
                resolve(null);
            };
            
            const handleInputKeydown = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    handleOk();
                }
            };
            
            const handleKeydown = (e) => {
                if (e.key === 'Escape') {
                    handleCancel();
                }
            };
            
            // Attach listeners
            okBtn.addEventListener('click', handleOk);
            cancelBtn.addEventListener('click', handleCancel);
            inputEl.addEventListener('keydown', handleInputKeydown);
            document.addEventListener('keydown', handleKeydown);
            
            // Close on overlay click
            modal.addEventListener('click', (e) => {
                if (e.target === modal) handleCancel();
            }, { once: true });
        });
    }
};

// Export for use in other modules
window.Utils = Utils;