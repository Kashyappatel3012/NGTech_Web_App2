function toggleCctvHistoryInput(show) {
  const input = document.getElementById('cctvHistoryDuration');
  if (!input) return;
  input.style.display = show ? '' : 'none';
  if (!show) {
    input.value = '';
  }
}

function toggleLockerCctvHistoryInput(show) {
  const input = document.getElementById('lockerCctvHistoryDuration');
  if (!input) return;
  input.style.display = show ? '' : 'none';
  if (!show) {
    input.value = '';
  }
}
document.addEventListener('DOMContentLoaded', function() {
    // --- Chart.js Bar Chart with Color Logic ---
    const performanceData = window.performanceData || {
        punctuality: 0,
        clientSatisfaction: 0,
        behaviour: 0,
        communicationSkills: 0,
        technicalSkills: 0,
        teamCoordination: 0
    };

    function getBarColor(score) {
        if (score <= 6) return 'rgba(255, 99, 132, 0.8)';      // Red
        if (score <= 8) return 'rgba(255, 206, 86, 0.8)';       // Yellow
        return 'rgba(75, 192, 192, 0.8)';                       // Green
    }
    function getBarBorderColor(score) {
        if (score <= 6) return 'rgba(255, 99, 132, 1)';
        if (score <= 8) return 'rgba(255, 206, 86, 1)';
        return 'rgba(75, 192, 192, 1)';
    }

    const perfScores = [
        performanceData.punctuality,
        performanceData.clientSatisfaction,
        performanceData.behaviour,
        performanceData.communicationSkills,
        performanceData.technicalSkills,
        performanceData.teamCoordination
    ];
    const barColors = perfScores.map(getBarColor);
    const barBorderColors = perfScores.map(getBarBorderColor);

    // Wait for Chart.js to load
    let chartInitRetries = 0;
    const MAX_CHART_INIT_RETRIES = 50; // Maximum 5 seconds (50 * 100ms)
    
    function initializeCharts() {
        if (typeof Chart === 'undefined') {
            chartInitRetries++;
            if (chartInitRetries >= MAX_CHART_INIT_RETRIES) {
                console.error('Chart.js failed to load after maximum retries. Please check your internet connection or CDN availability.');
                return;
            }
            console.warn('Chart.js not loaded yet, retrying... (' + chartInitRetries + '/' + MAX_CHART_INIT_RETRIES + ')');
            setTimeout(initializeCharts, 100);
            return;
        }
        
        // Reset retry counter on success
        chartInitRetries = 0;

        // Bar Chart - Previous Month Performance
        const progressCtx = document.getElementById('progressChart');
        if (progressCtx) {
            console.log('Initializing progressChart...');
            console.log('Performance data:', performanceData);
            console.log('Performance scores:', perfScores);
            
            // Destroy existing chart if any
            if (progressCtx.chartInstance) {
                progressCtx.chartInstance.destroy();
                progressCtx.chartInstance = null;
            }
            
            try {
                // Ensure canvas is visible and has dimensions
                const canvasContainer = progressCtx.parentElement;
                if (canvasContainer) {
                    canvasContainer.style.display = 'block';
                    canvasContainer.style.width = '100%';
                    canvasContainer.style.height = '260px';
                    canvasContainer.style.position = 'relative';
                }
                
                progressCtx.style.display = 'block';
                progressCtx.width = progressCtx.offsetWidth || 500;
                progressCtx.height = progressCtx.offsetHeight || 260;
                
                progressCtx.chartInstance = new Chart(progressCtx.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: [
                            'Punctuality',
                            'Client\nSatisfaction',
                            'Behaviour',
                            'Communication\nSkills',
                            'Technical\nSkills',
                            'Team\nCoordination'
                        ],
                        datasets: [{
                            label: 'Performance Metrics',
                            data: perfScores,
                            backgroundColor: barColors,
                            borderColor: barBorderColors,
                            borderWidth: 2,
                            barPercentage: 0.6,
                            categoryPercentage: 0.8
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        animation: {
                            duration: 1000,
                            easing: 'easeInOutQuart'
                        },
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                enabled: true,
                                callbacks: {
                                    label: function(context) {
                                        return `Score: ${context.parsed.y}/10`;
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                ticks: {
                                    font: { size: 12 },
                                    callback: function(value) {
                                        const label = this.getLabelForValue(value);
                                        return label.split('\n');
                                    },
                                    maxRotation: 0,
                                    minRotation: 0,
                                    autoSkip: false
                                },
                                grid: {
                                    display: true
                                }
                            },
                            y: {
                                beginAtZero: true,
                                max: 10,
                                ticks: {
                                    stepSize: 1,
                                    font: { size: 11 }
                                },
                                grid: {
                                    display: true
                                }
                            }
                        }
                    }
                });
                console.log('Progress chart initialized successfully');
                console.log('Chart instance:', progressCtx.chartInstance);
            } catch (error) {
                console.error('Error initializing progress chart:', error);
                console.error('Error stack:', error.stack);
            }
        } else {
            console.error('progressChart canvas element not found!');
        }

        // Line Chart - Last Year Performance
        const lastYearCtx = document.getElementById('lastYearChart');
        if (lastYearCtx) {
            console.log('Initializing lastYearChart...');
            console.log('Performance history:', window.performanceHistory);
            
            // Destroy existing chart if any
            if (lastYearCtx.chartInstance) {
                lastYearCtx.chartInstance.destroy();
                lastYearCtx.chartInstance = null;
            }
            
            try {
                // Ensure canvas is visible and has dimensions
                const canvasContainer = lastYearCtx.parentElement;
                if (canvasContainer) {
                    canvasContainer.style.display = 'block';
                    canvasContainer.style.width = '100%';
                    canvasContainer.style.height = '260px';
                    canvasContainer.style.position = 'relative';
                }
                
                lastYearCtx.style.display = 'block';
                lastYearCtx.width = lastYearCtx.offsetWidth || 500;
                lastYearCtx.height = lastYearCtx.offsetHeight || 260;
                
                const historyData = window.performanceHistory && Array.isArray(window.performanceHistory) 
                    ? window.performanceHistory 
                    : [];
                
                if (historyData.length > 0) {
                    // Extract labels and data safely
                    const labels = [];
                    const data = [];
                    
                    for (let i = 0; i < historyData.length; i++) {
                        const item = historyData[i];
                        if (item && item.month_name && item.year_short) {
                            labels.push(`${item.month_name}\n'${item.year_short}`);
                            data.push(item.average !== undefined && item.average !== null ? item.average : 0);
                        }
                    }
                    
                    console.log('Last year chart labels:', labels);
                    console.log('Last year chart data:', data);
                    
                    if (labels.length > 0 && data.length > 0) {
                        // Calculate min safely
                        const validData = data.filter(d => !isNaN(d) && d !== null && d !== undefined && d >= 0);
                        const dataMin = validData.length > 0 ? Math.min(...validData) : 0;
                        const yMin = Math.max(0, dataMin > 0 ? (dataMin - 1) : 0);
                        
                        lastYearCtx.chartInstance = new Chart(lastYearCtx.getContext('2d'), {
                            type: 'line',
                            data: {
                                labels: labels,
                                datasets: [{
                                    label: 'Performance Trend',
                                    data: data,
                                    borderColor: 'rgba(75, 192, 192, 1)',
                                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                                    tension: 0.1,
                                    fill: true,
                                    pointBackgroundColor: 'rgba(75, 192, 192, 1)',
                                    pointBorderColor: '#fff',
                                    pointRadius: 4,
                                    pointHoverRadius: 6,
                                    pointHoverBackgroundColor: '#fff',
                                    pointHoverBorderColor: 'rgba(75, 192, 192, 1)'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: {
                                    legend: { 
                                        display: false,
                                        position: 'top'
                                    },
                                    tooltip: {
                                        callbacks: {
                                            title: function(context) {
                                                if (historyData.length > 0 && context.length > 0) {
                                                    const index = context[0].dataIndex;
                                                    const item = historyData[index];
                                                    if (item && item.year && item.month) {
                                                        const fullMonth = new Date(item.year, item.month - 1).toLocaleString('default', { month: 'long' });
                                                        return `${fullMonth} ${item.year}`;
                                                    }
                                                }
                                                return '';
                                            },
                                            label: function(context) {
                                                return `Score: ${context.raw}/10`;
                                            }
                                        }
                                    }
                                },
                                scales: {
                                    y: {
                                        beginAtZero: validData.length === 0 || dataMin === 0,
                                        min: validData.length > 0 ? yMin : 0,
                                        max: 10,
                                        title: {
                                            display: true,
                                            text: 'Performance Score'
                                        },
                                        ticks: {
                                            stepSize: 1
                                        }
                                    },
                                    x: {
                                        title: {
                                            display: true,
                                            text: 'Timeline'
                                        },
                                        ticks: {
                                            autoSkip: false,
                                            maxRotation: 90,
                                            minRotation: 90,
                                            font: {
                                                size: 12
                                            }
                                        }
                                    }
                                }
                            }
                        });
                        console.log('Last year chart initialized successfully');
                    } else {
                        console.warn('No valid data for last year chart');
                        // Show placeholder
                        lastYearCtx.chartInstance = new Chart(lastYearCtx.getContext('2d'), {
                            type: 'line',
                            data: {
                                labels: ['No Data Available'],
                                datasets: [{ label: 'Performance Trend', data: [0], borderColor: 'rgba(200, 200, 200, 0.5)' }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: { y: { beginAtZero: true, max: 10 } }
                            }
                        });
                    }
                } else {
                    console.warn('No performance history data available');
                    // Show placeholder
                    lastYearCtx.chartInstance = new Chart(lastYearCtx.getContext('2d'), {
                        type: 'line',
                        data: {
                            labels: ['No Data Available'],
                            datasets: [{ label: 'Performance Trend', data: [0], borderColor: 'rgba(200, 200, 200, 0.5)' }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: { legend: { display: false } },
                            scales: { y: { beginAtZero: true, max: 10 } }
                        }
                    });
                }
            } catch (error) {
                console.error('Error initializing last year chart:', error);
                console.error('Error stack:', error.stack);
            }
        } else {
            console.error('lastYearChart canvas element not found!');
        }
    }
    
    // Start chart initialization with delay to ensure DOM and Chart.js are ready
    setTimeout(function() {
        initializeCharts();
    }, 200);

// Disaster Recovery Evidence Form Submission Handler
function handleDisasterRecoveryEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_disaster_recovery_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const filename = response.headers.get('Content-Disposition');
            let downloadName = 'Disaster Recovery Review.xlsx';
            if (filename) {
                const filenameMatch = filename.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('disasterRecoveryEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            // Try to get error message from response
            return response.text().then(text => {
                let errorMessage = 'An error occurred during report generation. Please try again.';
                try {
                    // Try to parse as JSON if possible
                    const errorData = JSON.parse(text);
                    if (errorData.error || errorData.message) {
                        errorMessage = errorData.error || errorData.message;
                    }
                } catch (e) {
                    // If not JSON, check if it's an HTML redirect (common for Flask flash messages)
                    if (text.includes('error') || response.status >= 400) {
                        console.error('Server response status:', response.status);
                        console.error('Server response:', text.substring(0, 200));
                    }
                }
                throw new Error(errorMessage);
            });
        }
    })
    .catch(error => {
        console.error('Error:', error);
        const errorMessage = error.message || 'An error occurred during report generation. Please try again.';
        alert(errorMessage);
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}
    // --- UI Functionality (unchanged) ---
    // Tab Functionality
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(`${tabId}-tab`).classList.add('active');
        });
    });

    // Smooth scrolling for sidebar links
    document.querySelectorAll('.sidebar nav a').forEach(anchor => {
        anchor.addEventListener('click', function(e) {
            if (this.getAttribute('href').startsWith('#')) {
                e.preventDefault();
                document.querySelectorAll('.sidebar nav a').forEach(a => a.classList.remove('active'));
                this.classList.add('active');
                const targetId = this.getAttribute('href');
                document.querySelector(targetId).scrollIntoView({ behavior: 'smooth' });
            }
        });
    });

    // Chat prompt selection
    const promptSelect = document.getElementById('prompt-select');
    if (promptSelect) {
        promptSelect.addEventListener('change', function() {
            if (this.value) {
                const promptText = this.options[this.selectedIndex].text;
                document.querySelector('.chat-input input').value = promptText;
            }
        });
    }

    // Accordion functionality
    document.querySelectorAll('.accordion-btn').forEach(btn => {
        btn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            
            this.classList.toggle('active');
            const content = this.nextElementSibling;
            const chevron = this.querySelector('.fa-chevron-down');
            
            // Check if content element exists and has the correct class
            if (!content || !content.classList.contains('accordion-content')) {
                console.error('Accordion content not found or incorrect class');
                return;
            }
            
            // Get the current max-height value (inline style or computed)
            const currentMaxHeight = content.style.maxHeight;
            const computedStyle = window.getComputedStyle(content);
            const computedMaxHeight = computedStyle.maxHeight;
            
            // Determine if accordion is currently open
            // It's open if maxHeight is set and not 0px or 0
            const isOpen = (currentMaxHeight && 
                           currentMaxHeight !== '0px' && 
                           currentMaxHeight !== '0' &&
                           currentMaxHeight !== '') ||
                           (computedMaxHeight && 
                            computedMaxHeight !== '0px' && 
                            computedMaxHeight !== 'none');
            
            if (isOpen) {
                // Close accordion
                content.style.maxHeight = '0px';
                if (chevron) {
                    chevron.style.transform = 'rotate(0deg)';
                }
            } else {
                // Open accordion
                content.style.maxHeight = content.scrollHeight + "px";
                if (chevron) {
                    chevron.style.transform = 'rotate(180deg)';
                }
            }
        });
    });

    // Report button animations
    document.querySelectorAll('.report-option').forEach(btn => {
        btn.addEventListener('mouseenter', function() {
            this.style.transform = 'translateX(8px) scale(1.03)';
        });
        btn.addEventListener('mouseleave', function() {
            this.style.transform = 'translateX(0) scale(1)';
        });
    });

    // Work card animations
    document.querySelectorAll('.work-card').forEach(card => {
        card.addEventListener('mouseenter', function() {
            this.style.transform = 'translateY(-5px) scale(1.03)';
            this.querySelector('i').style.transform = 'scale(1.15)';
        });
        card.addEventListener('mouseleave', function() {
            this.style.transform = 'translateY(0) scale(1)';
            this.querySelector('i').style.transform = 'scale(1)';
        });
    });
});


// Add this JavaScript to your HTML template or a separate JS file
function showReportGenerationAlert(unmatchedCount) {
    // Create the dialog container
    const dialog = document.createElement('div');
    dialog.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.7);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 10000;
        font-family: Arial, sans-serif;
    `;

    // Create the dialog content
    const dialogContent = document.createElement('div');
    dialogContent.style.cssText = `
        background-color: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 5px 15px rgba(0, 0, 0, 0.3);
        max-width: 500px;
        width: 80%;
        text-align: center;
    `;

    // Create the title
    const title = document.createElement('h2');
    title.textContent = 'Report Generation Alert';
    title.style.cssText = `
        color: #3553E8;
        margin-top: 0;
        margin-bottom: 20px;
    `;

    // Create the message
    const message = document.createElement('p');
    message.textContent = `The report generation process has completed.`;
    message.style.cssText = `
        margin-bottom: 15px;
        line-height: 1.5;
    `;

    // Create the unmatched vulnerabilities warning if applicable
    let warningMessage = null;
    if (unmatchedCount > 0) {
        warningMessage = document.createElement('p');
        warningMessage.textContent = `Warning: ${unmatchedCount} vulnerability/vulnerabilities could not be matched with the catalog and will not appear in the Infra_VAPT worksheet.`;
        warningMessage.style.cssText = `
            color: #FF0000;
            font-weight: bold;
            margin-bottom: 20px;
            line-height: 1.5;
        `;
    }

    // Create the download button
    const downloadButton = document.createElement('button');
    downloadButton.textContent = 'Download Report';
    downloadButton.style.cssText = `
        background-color: #3553E8;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
        font-weight: bold;
        margin-top: 10px;
    `;
    downloadButton.onclick = function() {
        dialog.remove();
    };

    // Create the close button
    const closeButton = document.createElement('button');
    closeButton.textContent = 'Close';
    closeButton.style.cssText = `
        background-color: #f0f0f0;
        color: #333;
        border: 1px solid #ccc;
        padding: 12px 24px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
        margin-left: 10px;
    `;
    closeButton.onclick = function() {
        dialog.remove();
    };

    // Assemble the dialog
    dialogContent.appendChild(title);
    dialogContent.appendChild(message);
    if (warningMessage) {
        dialogContent.appendChild(warningMessage);
    }
    
    const buttonContainer = document.createElement('div');
    buttonContainer.style.marginTop = '20px';
    buttonContainer.appendChild(downloadButton);
    buttonContainer.appendChild(closeButton);
    dialogContent.appendChild(buttonContainer);
    
    dialog.appendChild(dialogContent);
    document.body.appendChild(dialog);

    // Make dialog closable by clicking outside
    dialog.addEventListener('click', function(e) {
        if (e.target === dialog) {
            dialog.remove();
        }
    });
}

// Disaster Recovery Evidence Form Submission Handler
function handleDisasterRecoveryEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_disaster_recovery_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const filename = response.headers.get('Content-Disposition');
            let downloadName = 'Disaster Recovery Review.xlsx';
            if (filename) {
                const filenameMatch = filename.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('disasterRecoveryEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            // Try to get error message from response
            return response.text().then(text => {
                let errorMessage = 'An error occurred during report generation. Please try again.';
                try {
                    // Try to parse as JSON if possible
                    const errorData = JSON.parse(text);
                    if (errorData.error || errorData.message) {
                        errorMessage = errorData.error || errorData.message;
                    }
                } catch (e) {
                    // If not JSON, check if it's an HTML redirect (common for Flask flash messages)
                    if (text.includes('error') || response.status >= 400) {
                        console.error('Server response status:', response.status);
                        console.error('Server response:', text.substring(0, 200));
                    }
                }
                throw new Error(errorMessage);
            });
        }
    })
    .catch(error => {
        console.error('Error:', error);
        const errorMessage = error.message || 'An error occurred during report generation. Please try again.';
        alert(errorMessage);
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Example using Fetch API (modify according to your current implementation)
document.getElementById('uploadForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const formData = new FormData(this);
    
    try {
        const response = await fetch('/process_first_audit_report', {
            method: 'POST',
            body: formData
        });
        
        // Get the unmatched count and filename from the response headers
        const unmatchedCount = parseInt(response.headers.get('X-Unmatched-Vulnerabilities') || '0');
        const filename = response.headers.get('X-Filename') || 'combined_scan_results.xlsx';
        
        // Create a blob from the response
        const blob = await response.blob();
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        
        // Show the alert dialog
        showReportGenerationAlert(unmatchedCount);
        
    } catch (error) {
        console.error('Error:', error);
        alert('An error occurred during report generation.');
    }
});

// Network Review Evidence Form Submission Handler
function handleNetworkReviewEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_network_review_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Network Review.xlsx';
            if (contentDisposition) {
                // Extract filename from Content-Disposition header
                // Pattern: filename="Network Review.xlsx" or filename=Network Review.xlsx
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim();
                    // Remove surrounding quotes if present
                    extracted = extracted.replace(/^["']|["']$/g, '');
                    // Clean up any extra whitespace
                    extracted = extracted.trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('networkReviewEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Disaster Recovery Evidence Form Submission Handler
function handleDisasterRecoveryEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_disaster_recovery_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const filename = response.headers.get('Content-Disposition');
            let downloadName = 'Disaster Recovery Review.xlsx';
            if (filename) {
                const filenameMatch = filename.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('disasterRecoveryEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            // Try to get error message from response
            return response.text().then(text => {
                let errorMessage = 'An error occurred during report generation. Please try again.';
                try {
                    // Try to parse as JSON if possible
                    const errorData = JSON.parse(text);
                    if (errorData.error || errorData.message) {
                        errorMessage = errorData.error || errorData.message;
                    }
                } catch (e) {
                    // If not JSON, check if it's an HTML redirect (common for Flask flash messages)
                    if (text.includes('error') || response.status >= 400) {
                        console.error('Server response status:', response.status);
                        console.error('Server response:', text.substring(0, 200));
                    }
                }
                throw new Error(errorMessage);
            });
        }
    })
    .catch(error => {
        console.error('Error:', error);
        const errorMessage = error.message || 'An error occurred during report generation. Please try again.';
        alert(errorMessage);
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Data Centre Evidence Form Submission Handler
function handleDataCentreEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_data_centre_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Data Centre Review.xlsx';
            if (contentDisposition) {
                // Extract filename from Content-Disposition header
                // Pattern: filename="Data Centre.xlsx" or filename=Data Centre.xlsx
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim();
                    // Remove surrounding quotes if present
                    extracted = extracted.replace(/^["']|["']$/g, '');
                    // Clean up any extra whitespace
                    extracted = extracted.trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('dataCentreEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Disaster Recovery Evidence Form Submission Handler
function handleDisasterRecoveryEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_disaster_recovery_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const filename = response.headers.get('Content-Disposition');
            let downloadName = 'Disaster Recovery Review.xlsx';
            if (filename) {
                const filenameMatch = filename.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('disasterRecoveryEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            // Try to get error message from response
            return response.text().then(text => {
                let errorMessage = 'An error occurred during report generation. Please try again.';
                try {
                    // Try to parse as JSON if possible
                    const errorData = JSON.parse(text);
                    if (errorData.error || errorData.message) {
                        errorMessage = errorData.error || errorData.message;
                    }
                } catch (e) {
                    // If not JSON, check if it's an HTML redirect (common for Flask flash messages)
                    if (text.includes('error') || response.status >= 400) {
                        console.error('Server response status:', response.status);
                        console.error('Server response:', text.substring(0, 200));
                    }
                }
                throw new Error(errorMessage);
            });
        }
    })
    .catch(error => {
        console.error('Error:', error);
        const errorMessage = error.message || 'An error occurred during report generation. Please try again.';
        alert(errorMessage);
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Firewall Evidence Form Submission Handler
function handleFirewallEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_firewall_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Firewall Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('firewallEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Core Switch Evidence Form Submission Handler
function handleCoreSwitchEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_core_switch_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Core Switch Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('coreSwitchEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Router Evidence Form Submission Handler
function handleRouterEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_router_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Router Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('routerEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Domain Controller Evidence Form Submission Handler
function handleDomainControllerEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_domain_controller_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Domain Controller AD Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('domainControllerEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// H2H Audit Evidence Form Submission Handler
function handleH2HEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_h2h_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'H2H Audit Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('h2hEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Antivirus Evidence Form Submission Handler
function handleAntivirusEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_antivirus_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Antivirus Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('antivirusEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// ATM Evidence Form Submission Handler
function handleATMEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_atm_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'ATM Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('atmEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Mail and Messaging Evidence Form Submission Handler
function handleMailMessagingEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_mail_messaging_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Mail and Messaging Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('mailMessagingEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// HO Win_Server Evidence Form Submission Handler
function handleHoWinServerEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_ho_win_server_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'HO Win Server Logical Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('hoWinServerEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Linux_Server Evidence Form Submission Handler
function handleLinuxServerEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_linux_server_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Linux Server Logical Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('linuxServerEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// ESXi Server Evidence Form Submission Handler
function handleEsxiServerEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_esxi_server_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'ESXi Server Logical Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('esxiServerEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Access Control – OS Level Evidence Form Submission Handler
function handleAccessControlOSEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_access_control_os_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Access Control OS Level Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('accessControlOSEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Access Control Application Evidence Form Submission Handler
function handleAccessControlApplicationEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_access_control_application_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Access Control Application Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('accessControlApplicationEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Application Evidence Form Submission Handler
function handleApplicationEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_application_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Application Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('applicationEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Internet Banking Evidence Form Submission Handler
function handleInternetBankingEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_internet_banking_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Internet Banking Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('internetBankingEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Internal Control Evaluation Evidence Form Submission Handler
function handleInternalControlEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_internal_control_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Internal Control Evaluation Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('internalControlEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Fire Protection Evidence Form Submission Handler
function handleFireProtectionEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_fire_protection_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Fire Protection Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('fireProtectionEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// AMC Evidence Form Submission Handler
function handleAmcEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_amc_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'AMC Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('amcEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Data Input Controls Evidence Form Submission Handler
function handleDataInputControlEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_data_input_control_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Data Input Control Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('dataInputControlEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Purging of Data Files Evidence Form Submission Handler
function handlePurgingDataFilesEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_purging_data_files_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Purging of Data Files Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('purgingDataFilesEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Business Continuity Planning Evidence Form Submission Handler
function handleBusinessContinuityPlanningEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_business_continuity_planning_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Business Continuity Planning Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('businessContinuityPlanningEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// In-house and Out-sourced Evidence Form Submission Handler
function handleInhouseOutsourcedEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_inhouse_outsourced_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'In-house and Out-sourced Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('inhouseOutsourcedEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Audit Trail Evidence Form Submission Handler
function handleAuditTrailEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_audit_trail_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Audit Trail Review.xlsx';
            if (contentDisposition) {
                // Extract filename from Content-Disposition header
                // Pattern: filename="Audit Trail.xlsx" or filename=Audit Trail.xlsx
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim();
                    // Remove surrounding quotes if present
                    extracted = extracted.replace(/^["']|["']$/g, '');
                    // Clean up any extra whitespace
                    extracted = extracted.trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('auditTrailEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Packaged Software Evidence Form Submission Handler
function handlePackagedSoftwareEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_packaged_software_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Packaged Software Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('packagedSoftwareEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// User Account Maintenance Evidence Form Submission Handler
function handleUserAccountMaintenanceEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_user_account_maintenance_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'User Account Maintenance Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('userAccountMaintenanceEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Logical Access Controls Evidence Form Submission Handler
function handleLogicalAccessControlsEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_logical_access_controls_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Logical Access Controls Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('logicalAccessControlsEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Database Controls Evidence Form Submission Handler
function handleDatabaseControlsEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_database_controls_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Database Controls Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('databaseControlsEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Penetration Testing Evidence Form Submission Handler
function handlePenetrationTestingEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_penetration_testing_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Penetration Testing Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('penetrationTestingEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Training Evidence Form Submission Handler
function handleTrainingEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_training_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Training Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('trainingEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Remote Access Evidence Form Submission Handler
function handleRemoteAccessEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_remote_access_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Remote Access Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('remoteAccessEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Power Supply Evidence Form Submission Handler
function handlePowerSupplyEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_power_supply_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Power Supply UPS Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('powerSupplyEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Backup and Restoration Evidence Form Submission Handler
function handleBackupRestorationEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_backup_restoration_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Backup and Restoration Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('backupRestorationEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Maintenance & App Patches Evidence Form Submission Handler
function handleMaintenancePatchesEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_maintenance_patches_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Maintenance and App Patches Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('maintenancePatchesEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Network Monitor Tool Evidence Form Submission Handler
function handleNetworkMonitorToolEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_network_monitoring_tool_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Network Monitor Tool Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('networkMonitorToolEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// SAN Switch CISCO Evidence Form Submission Handler
function handleSanSwitchCiscoEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_san_switch_cisco_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'SAN Switch CISCO Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('sanSwitchCiscoEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// SAN Storage Evidence Form Submission Handler
function handleSanStorageEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_san_storage_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'SAN Storage Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('sanStorageEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// NAS Evidence Form Submission Handler
function handleNasEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_nas_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'NAS Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('nasEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Load Balancer Array Evidence Form Submission Handler
function handleLoadBalancerArrayEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_load_balancer_array_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Load Balancer Array Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('loadBalancerArrayEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// PAM Evidence Form Submission Handler
function handlePamEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_pam_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'PAM Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('pamEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// SOC Evidence Form Submission Handler
function handleSocEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_soc_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'SOC Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('socEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Change Management Evidence Form Submission Handler
function handleChangeManagementEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_change_management_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Change Management Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('changeManagementEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Asset Management Evidence Form Submission Handler
function handleAssetManagementEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_asset_management_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Asset Management Review.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim().replace(/^["']|["']$/g, '').trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('assetManagementEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}

// Others Evidence Form Submission Handler
function handleOthersEvidenceSubmit(event) {
    event.preventDefault();
    
    const form = event.target;
    const formData = new FormData(form);
    
    // Show loading state
    const submitButton = form.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.textContent = 'Processing...';
    submitButton.disabled = true;
    
    fetch('/process_others_evidence', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            // Get the filename from the response headers
            const contentDisposition = response.headers.get('Content-Disposition');
            let downloadName = 'Others Review.xlsx';
            if (contentDisposition) {
                // Extract filename from Content-Disposition header
                // Pattern: filename="Others.xlsx" or filename=Others.xlsx
                const filenameMatch = contentDisposition.match(/filename[^;]*=([^;]*)/);
                if (filenameMatch) {
                    let extracted = filenameMatch[1].trim();
                    // Remove surrounding quotes if present
                    extracted = extracted.replace(/^["']|["']$/g, '');
                    // Clean up any extra whitespace
                    extracted = extracted.trim();
                    if (extracted && extracted !== '') {
                        downloadName = extracted;
                    }
                }
            }
            
            // Create a blob from the response
            return response.blob().then(blob => {
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = downloadName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                // Close the modal
                closeModal('othersEvidenceModal');
                
                // Refresh the page
                window.location.reload();
            });
        } else {
            throw new Error('Network response was not ok');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred during report generation. Please try again.');
    })
    .finally(() => {
        // Reset button state
        submitButton.textContent = originalText;
        submitButton.disabled = false;
    });
}
// Handle Combine Asset Review Excels form submission with page refresh
function handleCombineAssetReviewExcelsSubmit(event) {
    console.log('Combine Asset Review Excels form submission started');
    event.preventDefault(); // Prevent default form submission
    
    // Show loading state
    const submitBtn = document.getElementById('combineAssetReviewExcelsSubmitBtn');
    const originalText = submitBtn.innerHTML;
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    // Get form data
    const form = document.getElementById('combineAssetReviewExcelsForm');
    const formData = new FormData(form);
    
    console.log('Submitting form with fetch...');
    
    // Submit form using fetch
    fetch('/combine_assets_excels', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        console.log('Response received:', response);
        
        if (response.ok) {
            // Check if response is a file download
            const contentType = response.headers.get('content-type');
            const contentDisposition = response.headers.get('content-disposition');
            
            console.log('Content-Type:', contentType);
            console.log('Content-Disposition:', contentDisposition);
            
            if (contentType && contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
                // Extract filename from content-disposition header
                let filename = 'Asset_Review_Excel_Files_Combined.xlsx';
                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename="([^"]+)"/);
                    if (filenameMatch) {
                        filename = filenameMatch[1];
                    }
                }
                
                console.log('Downloading file:', filename);
                
                // Convert response to blob and create download link
                return response.blob().then(blob => {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    
                    console.log('File downloaded successfully');
                    
                    // Close modal and refresh page
                    closeModal('combineAssetReviewExcelsModal');
                    setTimeout(() => {
                        window.location.reload();
                    }, 1000);
                });
            } else {
                // Handle non-file response
                return response.text().then(text => {
                    console.log('Response text:', text);
                    alert('Unexpected response format');
                });
            }
        } else {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred: ' + error.message);
    })
    .finally(() => {
        // Reset button state
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}
