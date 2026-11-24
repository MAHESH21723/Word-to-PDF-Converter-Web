document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const themeIcon = themeToggleBtn.querySelector('i');
    const themeText = themeToggleBtn.querySelector('span');
    
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileSelectedNode = document.getElementById('fileSelectedNode');
    const selectedFileName = document.getElementById('selectedFileName');
    const selectedFileSize = document.getElementById('selectedFileSize');
    const btnRemoveFile = document.getElementById('btnRemoveFile');
    
    const btnConvert = document.getElementById('btnConvert');
    const progressContainer = document.getElementById('progressContainer');
    const progressBarFill = document.getElementById('progressBarFill');
    const progressStatusText = document.getElementById('progressStatusText');
    const progressPercentText = document.getElementById('progressPercentText');
    
    const resultBox = document.getElementById('resultBox');
    const successResult = document.getElementById('successResult');
    const errorResult = document.getElementById('errorResult');
    const successMessage = document.getElementById('successMessage');
    const errorMessage = document.getElementById('errorMessage');
    const btnDownloadPDF = document.getElementById('btnDownloadPDF');
    
    const historyList = document.getElementById('historyList');
    const searchInput = document.getElementById('searchInput');
    const btnClearHistory = document.getElementById('btnClearHistory');
    
    // Stats Elements
    const statTotal = document.getElementById('statTotal');
    const statSuccessRate = document.getElementById('statSuccessRate');
    const statAvgSpeed = document.getElementById('statAvgSpeed');
    
    // Sidebar toggle (for mobile screens)
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.getElementById('sidebar');

    let selectedFile = null;

    // --- Theme Management ---
    const getSavedTheme = () => localStorage.getItem('theme') || 'light';
    const saveTheme = (theme) => localStorage.setItem('theme', theme);
    
    const applyTheme = (theme) => {
        document.documentElement.setAttribute('data-theme', theme);
        if (theme === 'dark') {
            themeIcon.className = 'bi bi-sun-fill';
            themeText.textContent = 'Light Mode';
        } else {
            themeIcon.className = 'bi bi-moon-stars-fill';
            themeText.textContent = 'Dark Mode';
        }
    };
    
    // Initialize Theme
    applyTheme(getSavedTheme());
    
    themeToggleBtn.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        applyTheme(newTheme);
        saveTheme(newTheme);
    });

    // --- Sidebar Mobile Toggle ---
    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('show');
        });
        
        // Close sidebar if clicking outside of it on mobile
        document.addEventListener('click', (e) => {
            if (sidebar.classList.contains('show') && 
                !sidebar.contains(e.target) && 
                !sidebarToggle.contains(e.target)) {
                sidebar.classList.remove('show');
            }
        });
    }

    // --- Drag and Drop File Handlers ---
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('drag-over');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('drag-over');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFileSelection(files[0]);
        }
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelection(e.target.files[0]);
        }
    });

    btnRemoveFile.addEventListener('click', (e) => {
        e.stopPropagation();
        resetFileSelection();
    });

    // --- Helper formatting ---
    const formatBytes = (bytes, decimals = 2) => {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    };

    const handleFileSelection = (file) => {
        // Validation
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext !== 'docx') {
            showAlert('Invalid file type! Only Microsoft Word (.docx) documents are supported.', 'danger');
            resetFileSelection();
            return;
        }

        const maxSize = 15 * 1024 * 1024; // 15MB
        if (file.size > maxSize) {
            showAlert('File is too large! Maximum allowed file size is 15MB.', 'danger');
            resetFileSelection();
            return;
        }

        selectedFile = file;
        selectedFileName.textContent = file.name;
        selectedFileSize.textContent = formatBytes(file.size);
        fileSelectedNode.style.display = 'flex';
        btnConvert.disabled = false;
        
        // Hide previous results
        resultBox.style.display = 'none';
        successResult.style.display = 'none';
        errorResult.style.display = 'none';
    };

    const resetFileSelection = () => {
        selectedFile = null;
        fileInput.value = '';
        fileSelectedNode.style.display = 'none';
        btnConvert.disabled = true;
    };

    // --- Alert Helper ---
    const showAlert = (message, type = 'success') => {
        // We will inject a temporary floating Toast Alert into the main layout
        const alertNode = document.createElement('div');
        alertNode.className = `alert alert-${type} alert-dismissible fade show glass-card`;
        alertNode.style.position = 'fixed';
        alertNode.style.top = '20px';
        alertNode.style.right = '20px';
        alertNode.style.zIndex = '9999';
        alertNode.style.boxShadow = 'var(--shadow-lg)';
        alertNode.style.borderLeft = `5px solid var(--${type})`;
        alertNode.style.background = 'var(--card-bg)';
        alertNode.style.backdropFilter = 'var(--glass-blur)';
        alertNode.style.color = 'var(--text-primary)';
        
        const icon = type === 'success' ? 'bi-check-circle-fill' : type === 'danger' ? 'bi-exclamation-triangle-fill' : 'bi-info-circle-fill';
        
        alertNode.innerHTML = `
            <div class="d-flex align-items-center gap-2">
                <i class="bi ${icon}" style="color: var(--${type}); font-size: 1.25rem;"></i>
                <div>${message}</div>
            </div>
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close" style="filter: var(--text-primary)"></button>
        `;
        document.body.appendChild(alertNode);
        
        setTimeout(() => {
            alertNode.classList.remove('show');
            setTimeout(() => alertNode.remove(), 300);
        }, 4000);
    };

    // --- Conversion Handler ---
    btnConvert.addEventListener('click', () => {
        if (!selectedFile) return;
        
        // UI Preparation
        btnConvert.disabled = true;
        progressContainer.style.display = 'block';
        resultBox.style.display = 'none';
        successResult.style.display = 'none';
        errorResult.style.display = 'none';
        
        // Smoothly simulate progress up to 90%
        let progress = 0;
        progressBarFill.style.width = '0%';
        progressPercentText.textContent = '0%';
        progressStatusText.textContent = 'Uploading file...';
        
        const progressInterval = setInterval(() => {
            if (progress < 40) {
                progress += Math.floor(Math.random() * 8) + 4;
            } else if (progress < 75) {
                progressStatusText.textContent = 'Parsing Word content...';
                progress += Math.floor(Math.random() * 4) + 2;
            } else if (progress < 90) {
                progressStatusText.textContent = 'Generating PDF pages...';
                progress += Math.floor(Math.random() * 2) + 1;
            }
            
            if (progress > 90) progress = 90;
            
            progressBarFill.style.width = `${progress}%`;
            progressPercentText.textContent = `${progress}%`;
        }, 150);

        // Build FormData
        const formData = new FormData();
        formData.append('word_file', selectedFile);

        fetch('/convert', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            // Always read the JSON body so server-side validation messages reach the user.
            return response.json().catch(() => ({})).then(body => {
                if (!response.ok && body.status !== 'error') {
                    body = { status: 'error', message: `Server error (HTTP ${response.status})` };
                }
                return body;
            });
        })
        .then(data => {
            clearInterval(progressInterval);
            
            if (data.status === 'success') {
                // Complete progress bar
                progressBarFill.style.width = '100%';
                progressPercentText.textContent = '100%';
                progressStatusText.textContent = 'Finished!';
                
                setTimeout(() => {
                    progressContainer.style.display = 'none';
                    resultBox.style.display = 'block';
                    successResult.style.display = 'block';
                    
                    successMessage.innerHTML = `
                        <strong>${htmlEscape(data.filename)}</strong> converted successfully in <strong>${data.conversion_time}s</strong>.
                    `;
                    
                    btnDownloadPDF.href = `/download/${data.pdf_filename}`;
                    showAlert('File converted successfully!', 'success');
                    
                    // Update lists and counts
                    loadHistory();
                    loadMetrics();
                    resetFileSelection();
                }, 400);
            } else {
                // Handle logic failures
                throw new Error(data.message || 'Conversion failed');
            }
        })
        .catch(err => {
            clearInterval(progressInterval);
            progressContainer.style.display = 'none';
            resultBox.style.display = 'block';
            errorResult.style.display = 'block';
            errorMessage.textContent = err.message || 'An unexpected error occurred during conversion.';
            showAlert('Conversion failed!', 'danger');
            btnConvert.disabled = false;
            
            // Reload history to capture failed entry if it is logged
            loadHistory();
            loadMetrics();
        });
    });

    // Simple HTML Escape function
    const htmlEscape = (str) => {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    };

    // --- Load metrics ---
    const loadMetrics = () => {
        fetch('/metrics')
            .then(res => res.json())
            .then(metrics => {
                statTotal.textContent = metrics.total;
                statSuccessRate.textContent = `${metrics.success_rate}%`;
                statAvgSpeed.textContent = `${metrics.avg_speed}s`;
            })
            .catch(err => console.error('Error fetching metrics:', err));
    };

    // --- Load history ---
    const loadHistory = () => {
        fetch('/history')
            .then(res => res.json())
            .then(data => {
                renderHistory(data);
            })
            .catch(err => {
                console.error('Error fetching history:', err);
                historyList.innerHTML = `
                    <div class="empty-history text-danger">
                        <i class="bi bi-exclamation-triangle-fill"></i>
                        <p>Failed to load conversion history.</p>
                    </div>
                `;
            });
    };

    const renderHistory = (items) => {
        if (!items || items.length === 0) {
            historyList.innerHTML = `
                <div class="empty-history">
                    <i class="bi bi-clock-history"></i>
                    <p>No conversions found yet. Upload a document to get started!</p>
                </div>
            `;
            return;
        }

        const query = searchInput.value.toLowerCase().trim();
        const filteredItems = items.filter(item => 
            item.filename.toLowerCase().includes(query) || 
            (item.converted_filename && item.converted_filename.toLowerCase().includes(query))
        );

        if (filteredItems.length === 0) {
            historyList.innerHTML = `
                <div class="empty-history">
                    <i class="bi bi-search"></i>
                    <p>No results matching "${htmlEscape(query)}" found.</p>
                </div>
            `;
            return;
        }

        historyList.innerHTML = filteredItems.map(item => {
            const isSuccess = item.status === 'success';
            const sizeStr = formatBytes(item.original_size);
            const dateStr = formatDate(item.created_at);
            
            let statusBadge = '';
            let actionBtn = '';
            
            if (isSuccess) {
                statusBadge = `<span class="status-badge success"><i class="bi bi-check-circle"></i> Success</span>`;
                actionBtn = `
                    <a href="/download/${item.converted_filename}" class="btn-action-icon download" title="Download PDF" download>
                        <i class="bi bi-download"></i>
                    </a>
                `;
            } else {
                statusBadge = `<span class="status-badge danger" title="${htmlEscape(item.error_message || '')}"><i class="bi bi-x-circle"></i> Failed</span>`;
                actionBtn = `
                    <button class="btn-action-icon" disabled style="opacity: 0.3;">
                        <i class="bi bi-download"></i>
                    </button>
                `;
            }

            return `
                <div class="history-item">
                    <div class="history-item-left">
                        <div class="history-item-icon docx">
                            <i class="bi bi-file-earmark-word-fill"></i>
                        </div>
                        <div class="history-item-details">
                            <div class="history-item-name" title="${htmlEscape(item.filename)}">${htmlEscape(item.filename)}</div>
                            <div class="history-item-meta">
                                <span>${sizeStr}</span>
                                <span>&bull;</span>
                                <span>${dateStr}</span>
                                <span>&bull;</span>
                                <span>${item.conversion_time ? item.conversion_time + 's' : 'N/A'}</span>
                            </div>
                        </div>
                    </div>
                    <div class="history-item-actions">
                        ${statusBadge}
                        ${actionBtn}
                    </div>
                </div>
            `;
        }).join('');
    };

    const formatDate = (dateString) => {
        // Format: "YYYY-MM-DD HH:MM:SS" SQLite string
        try {
            // Check if string contains T, space, or else
            const parts = dateString.split(/[- :]/);
            if (parts.length < 5) return dateString;
            
            // Re-render local date formats nicely, e.g. "Jun 9, 2026 10:45 PM"
            const year = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const day = parseInt(parts[2]);
            const hour = parseInt(parts[3]);
            const minute = parseInt(parts[4]);
            
            const date = new Date(year, month, day, hour, minute);
            
            return date.toLocaleDateString(undefined, { 
                month: 'short', 
                day: 'numeric', 
                year: 'numeric' 
            }) + ' ' + date.toLocaleTimeString(undefined, { 
                hour: '2-digit', 
                minute: '2-digit' 
            });
        } catch (e) {
            return dateString;
        }
    };

    // --- Search functionality ---
    searchInput.addEventListener('input', () => {
        // Instead of hitting server, filter locally (since we have the list loaded)
        loadHistory();
    });

    // --- Clear history button ---
    btnClearHistory.addEventListener('click', () => {
        if (confirm('Are you sure you want to clear your conversion history? This will delete all logged items.')) {
            fetch('/history/clear', {
                method: 'POST'
            })
            .then(res => res.json())
            .then(data => {
                if (data.status === 'success') {
                    showAlert('Conversion history cleared.', 'success');
                    loadHistory();
                    loadMetrics();
                } else {
                    showAlert('Failed to clear history.', 'danger');
                }
            })
            .catch(err => {
                console.error('Error clearing history:', err);
                showAlert('Failed to clear history.', 'danger');
            });
        }
    });

    // --- Initial Load ---
    loadMetrics();
    loadHistory();
});
