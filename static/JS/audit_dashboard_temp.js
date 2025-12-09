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
