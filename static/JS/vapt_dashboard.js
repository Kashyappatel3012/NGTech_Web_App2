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
                    backgroundColor: 'rgba(52, 152, 219, 0.1)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    pointBackgroundColor: 'rgba(52, 152, 219, 1)',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointHoverBackgroundColor: '#fff',
                    pointHoverBorderColor: 'rgba(52, 152, 219, 1)',
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0, 0, 0, 0.8)',
                        titleFont: { size: 14 },
                        bodyFont: { size: 13 },
                        padding: 12,
                        cornerRadius: 8,
                        displayColors: false,
                        callbacks: {
                            label: function(context) {
                                return 'Score: ' + context.parsed.y.toFixed(2);
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: { size: 11 },
                            maxRotation: 0,
                            minRotation: 0
                        }
                    },
                    y: {
                        beginAtZero: true,
                        max: 10,
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        },
                        ticks: {
                            font: { size: 11 },
                            stepSize: 2
                        }
                    }
                },
                interaction: {
                    intersect: false,
                    mode: 'index'
                }
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

    // Navigation functionality with smooth scrolling
    const navLinks = document.querySelectorAll('nav a');

    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            const href = this.getAttribute('href');
            
            // Skip logout and other non-section links
            if (href === '#' || !href.startsWith('#')) {
                return;
            }
            
            e.preventDefault();
            
            // Update active state
            navLinks.forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            
            // Smooth scroll to section
            const targetId = href.substring(1);
            const targetSection = document.getElementById(targetId);
            
            if (targetSection) {
                targetSection.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });
});

// Modal functions
function openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.style.display = 'block';
    }
}

function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.style.display = 'none';
    }
}

// Close modal when clicking outside
window.onclick = function(event) {
    if (event.target.classList.contains('modal')) {
        event.target.style.display = 'none';
    }
}

// Toggle Other Field for dropdowns
function toggleGrcOtherField(selectId, otherId) {
    const select = document.getElementById(selectId);
    const otherField = document.getElementById(otherId);
    
    if (select.value === 'Other') {
        otherField.style.display = 'block';
        otherField.required = true;
    } else {
        otherField.style.display = 'none';
        otherField.required = false;
        otherField.value = '';
    }
}

// GRC IS Audit Compliance Worksheet Functions
function toggleGrcInfraVAPTOtherOrganization() {
    const select = document.getElementById('grcInfraVAPTOrganizationName');
    const otherGroup = document.getElementById('grcInfraVAPTOtherOrganizationGroup');
    const otherInput = document.getElementById('grcInfraVAPTOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcIsAuditComplianceSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcIsAuditComplianceSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcIsAuditComplianceForm'));
    
    fetch('/grc_process_is_audit_compliance', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Download the file
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            // Close modal
            closeModal('grcIsAuditComplianceModal');
            
            // Reset form
            document.getElementById('grcIsAuditComplianceForm').reset();
            toggleGrcOtherField('grcIsAuditOrgName', 'grcIsAuditOrgNameOther'); // Reset the other org field visibility
            
            // Delete the file from server and reload page
            setTimeout(() => {
                fetch('/grc_cleanup_is_audit_compliance', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => {
                    // Reload the page after cleanup
                    location.reload();
                })
                .catch(err => {
                    console.error('Cleanup error:', err);
                    // Reload even if cleanup fails
                    location.reload();
                });
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC IS Audit Compliance Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Infrastructure VAPT Compliance Worksheet Functions
function handleGrcInfraVAPTComplianceSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcInfraVAPTComplianceSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcInfraVAPTComplianceForm'));
    
    fetch('/grc_process_infra_vapt_compliance', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcInfraVAPTComplianceModal');
            document.getElementById('grcInfraVAPTComplianceForm').reset();
            toggleGrcInfraVAPTOtherOrganization();
            
            setTimeout(() => {
                fetch('/grc_cleanup_infra_vapt_compliance', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Infrastructure VAPT Compliance Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Website VAPT Compliance Worksheet Functions
function toggleGrcWebsiteVAPTOtherOrganization() {
    const select = document.getElementById('grcWebsiteVAPTOrganizationName');
    const otherGroup = document.getElementById('grcWebsiteVAPTOtherOrganizationGroup');
    const otherInput = document.getElementById('grcWebsiteVAPTOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcWebsiteVAPTComplianceSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcWebsiteVAPTComplianceSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcWebsiteVAPTComplianceForm'));
    
    fetch('/grc_process_website_vapt_compliance', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcWebsiteVAPTComplianceModal');
            document.getElementById('grcWebsiteVAPTComplianceForm').reset();
            toggleGrcWebsiteVAPTOtherOrganization();
            
            setTimeout(() => {
                fetch('/grc_cleanup_website_vapt_compliance', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Website VAPT Compliance Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Public IP VAPT Compliance Worksheet Functions
function toggleGrcPublicIPVAPTOtherOrganization() {
    const select = document.getElementById('grcPublicIPVAPTOrganizationName');
    const otherGroup = document.getElementById('grcPublicIPVAPTOtherOrganizationGroup');
    const otherInput = document.getElementById('grcPublicIPVAPTOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcPublicIPVAPTComplianceSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcPublicIPVAPTComplianceSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcPublicIPVAPTComplianceForm'));
    
    fetch('/grc_process_public_ip_vapt_compliance', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcPublicIPVAPTComplianceModal');
            document.getElementById('grcPublicIPVAPTComplianceForm').reset();
            toggleGrcPublicIPVAPTOtherOrganization();
            
            setTimeout(() => {
                fetch('/grc_cleanup_public_ip_vapt_compliance', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Public IP VAPT Compliance Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Compliance Certificate Functions
function toggleGrcCertOtherOrganization() {
    const select = document.getElementById('grcCertOrganizationName');
    const otherGroup = document.getElementById('grcCertOtherOrganizationGroup');
    const otherInput = document.getElementById('grcCertOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function toggleGrcCertOtherFinancialYear() {
    const select = document.getElementById('grcCertFinancialYear');
    const otherGroup = document.getElementById('grcCertOtherFinancialYearGroup');
    const otherInput = document.getElementById('grcCertFinancialYearOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcIsAuditComplianceCertSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcIsAuditComplianceCertSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcIsAuditComplianceCertForm'));
    
    fetch('/grc_process_is_audit_compliance_certificate', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcIsAuditComplianceCertModal');
            document.getElementById('grcIsAuditComplianceCertForm').reset();
            toggleGrcCertOtherOrganization();
            toggleGrcCertOtherFinancialYear();
            
            setTimeout(() => {
                fetch('/grc_cleanup_is_audit_compliance_certificate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC IS Audit Compliance Certificate Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Infrastructure VAPT Compliance Certificate Functions
function toggleGrcInfraCertOtherOrganization() {
    const select = document.getElementById('grcInfraCertOrganizationName');
    const otherGroup = document.getElementById('grcInfraCertOtherOrganizationGroup');
    const otherInput = document.getElementById('grcInfraCertOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function toggleGrcInfraCertOtherFinancialYear() {
    const select = document.getElementById('grcInfraCertFinancialYear');
    const otherGroup = document.getElementById('grcInfraCertOtherFinancialYearGroup');
    const otherInput = document.getElementById('grcInfraCertFinancialYearOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcInfraVAPTComplianceCertSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcInfraVAPTComplianceCertSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcInfraVAPTComplianceCertForm'));
    
    fetch('/grc_process_infrastructure_vapt_compliance_certificate', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcInfraVAPTComplianceCertModal');
            document.getElementById('grcInfraVAPTComplianceCertForm').reset();
            toggleGrcInfraCertOtherOrganization();
            toggleGrcInfraCertOtherFinancialYear();
            
            setTimeout(() => {
                fetch('/grc_cleanup_infrastructure_vapt_compliance_certificate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Infrastructure VAPT Compliance Certificate Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Website VAPT Compliance Certificate Functions
function toggleGrcWebCertOtherOrganization() {
    const select = document.getElementById('grcWebCertOrganizationName');
    const otherGroup = document.getElementById('grcWebCertOtherOrganizationGroup');
    const otherInput = document.getElementById('grcWebCertOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function toggleGrcWebCertOtherFinancialYear() {
    const select = document.getElementById('grcWebCertFinancialYear');
    const otherGroup = document.getElementById('grcWebCertOtherFinancialYearGroup');
    const otherInput = document.getElementById('grcWebCertFinancialYearOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcWebsiteVAPTComplianceCertSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcWebsiteVAPTComplianceCertSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcWebsiteVAPTComplianceCertForm'));
    
    fetch('/grc_process_website_vapt_compliance_certificate', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcWebsiteVAPTComplianceCertModal');
            document.getElementById('grcWebsiteVAPTComplianceCertForm').reset();
            toggleGrcWebCertOtherOrganization();
            toggleGrcWebCertOtherFinancialYear();
            
            setTimeout(() => {
                fetch('/grc_cleanup_website_vapt_compliance_certificate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Website VAPT Compliance Certificate Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

// GRC Public IP VAPT Compliance Certificate Functions
function toggleGrcPubCertOtherOrganization() {
    const select = document.getElementById('grcPubCertOrganizationName');
    const otherGroup = document.getElementById('grcPubCertOtherOrganizationGroup');
    const otherInput = document.getElementById('grcPubCertOrganizationNameOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function toggleGrcPubCertOtherFinancialYear() {
    const select = document.getElementById('grcPubCertFinancialYear');
    const otherGroup = document.getElementById('grcPubCertOtherFinancialYearGroup');
    const otherInput = document.getElementById('grcPubCertFinancialYearOther');
    
    if (select.value === 'Other') {
        otherGroup.style.display = 'block';
        otherInput.required = true;
    } else {
        otherGroup.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
    }
}

function handleGrcPublicIPVAPTComplianceCertSubmit(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('grcPublicIPVAPTComplianceCertSubmitBtn');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
    submitBtn.disabled = true;
    
    const formData = new FormData(document.getElementById('grcPublicIPVAPTComplianceCertForm'));
    
    fetch('/grc_process_public_ip_vapt_compliance_certificate', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const link = document.createElement('a');
            link.href = data.download_url;
            link.download = data.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            closeModal('grcPublicIPVAPTComplianceCertModal');
            document.getElementById('grcPublicIPVAPTComplianceCertForm').reset();
            toggleGrcPubCertOtherOrganization();
            toggleGrcPubCertOtherFinancialYear();
            
            setTimeout(() => {
                fetch('/grc_cleanup_public_ip_vapt_compliance_certificate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: data.filename })
                })
                .then(() => location.reload())
                .catch(err => location.reload());
            }, 1000);
        } else {
            alert('Error: ' + (data.error || 'Unknown error'));
        }
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    })
    .catch(error => {
        console.error('GRC Public IP VAPT Compliance Certificate Error:', error);
        alert('An error occurred: ' + error.message);
        submitBtn.innerHTML = originalText;
        submitBtn.disabled = false;
    });
}

