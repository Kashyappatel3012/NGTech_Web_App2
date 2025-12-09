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

    // Bar Chart
    const progressCtx = document.getElementById('progressChart');
    if (progressCtx) {
        new Chart(progressCtx.getContext('2d'), {
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
                    barPercentage: 1.5,
                    categoryPercentage: 0.4
                }]
            },
            options: {
                plugins: {
                    legend: { display: false }
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
                        }
                    },
                    y: {
                        beginAtZero: true,
                        max: 10
                    }
                }
            }
        });
    }

    // --- Chart.js Line Chart for Last Year ---
// Replace the line chart code with this updated version
const lastYearCtx = document.getElementById('lastYearChart');
if (lastYearCtx && window.performanceHistory) {
    const historyData = window.performanceHistory;
    
    // Extract labels and data from performance history
    const labels = historyData.map(item => `${item.month_name}\n'${item.year_short}`);
    const data = historyData.map(item => item.average);
    
    new Chart(lastYearCtx.getContext('2d'), {
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
            plugins: {
                legend: { 
                    display: false,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            const index = context[0].dataIndex;
                            const item = historyData[index];
                            const fullMonth = new Date(item.year, item.month - 1).toLocaleString('default', { month: 'long' });
                            return `${fullMonth} ${item.year}`;
                        },
                        label: function(context) {
                            return `Score: ${context.raw}/10`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: false,
                    min: Math.max(0, Math.min(...data) - 1), // Dynamic min based on data
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
            const downloadName = filename ? filename.split('filename=')[1] : 'Disaster_Recovery_Evidence_With_POC.xlsx';
            
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
        btn.addEventListener('click', function() {
            this.classList.toggle('active');
            const content = this.nextElementSibling;
            if (content.style.maxHeight) {
                content.style.maxHeight = null;
                this.querySelector('.fa-chevron-down').style.transform = 'rotate(0deg)';
            } else {
                content.style.maxHeight = content.scrollHeight + "px";
                this.querySelector('.fa-chevron-down').style.transform = 'rotate(180deg)';
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
            const downloadName = filename ? filename.split('filename=')[1] : 'Disaster_Recovery_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Network_Review_Evidence_With_POC.xlsx';
            
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
            const downloadName = filename ? filename.split('filename=')[1] : 'Disaster_Recovery_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Data_Centre_Evidence_With_POC.xlsx';
            
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
            const downloadName = filename ? filename.split('filename=')[1] : 'Disaster_Recovery_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Firewall_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Core_Switch_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Router_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Domain_Controller_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'H2H_Audit_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Antivirus_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'ATM_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Mail_Messaging_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'HO_Win_Server_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Linux_Server_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'ESXi_Server_Evidence_With_POC.xlsx';
            
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

// Access Control â€“ OS Level Evidence Form Submission Handler
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Access_Control_OS_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Access_Control_Application_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Application_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Internet_Banking_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Internal_Control_Evaluation_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Fire_Protection_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'AMC_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Data_Input_Control_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Purging_of_Data_Files_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Business_Continuity_Planning_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'In_house_Out_Sou_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Audit_Trail_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Packaged_Software_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'User_Account_Maintenance_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Logical_Access_Controls_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Database_Controls_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Penetration_Testing_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Training_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Remote_Access_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Power_Supply_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Backup_Restoration_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Maintenance_App_Patches_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Network_Monitor_Tool_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'SAN_Switch_CISCO_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'SAN_Storage_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'NAS_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Load_Balancer_Array_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'PAM_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'SOC_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Change_Management_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Asset_Management_Evidence_With_POC.xlsx';
            
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
            const filename = response.headers.get('Content-Disposition');
            const downloadName = filename ? filename.split('filename=')[1] : 'Others_Evidence_With_POC.xlsx';
            
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
