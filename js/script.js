// Main script for the DOCX Text Replacer application

document.addEventListener('DOMContentLoaded', function() {
    // Tutorial elements
    const tutorialTrigger = document.getElementById('tutorialTrigger');
    const tutorialOverlay = document.getElementById('tutorialOverlay');
    const tutorialClose = document.getElementById('tutorialClose');
    const tutorialSteps = document.getElementById('tutorialSteps');
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const tutorialProgress = document.getElementById('tutorialProgress');
    const tutorialComplete = document.getElementById('tutorialComplete');
    const tutorialFinishBtn = document.getElementById('tutorialFinishBtn');
    
    // Initialize tutorial if all elements are present
    if (tutorialTrigger && tutorialOverlay) {
        initializeTutorial();
    }
    
    // Function to initialize the tutorial
    function initializeTutorial() {
        // Open tutorial on trigger click
        tutorialTrigger.addEventListener('click', function() {
            openTutorial();
        });
        
        // Close tutorial on close button click
        if (tutorialClose) {
            tutorialClose.addEventListener('click', function() {
                closeTutorial();
            });
        }
        
        // Close tutorial when clicking outside the modal
        tutorialOverlay.addEventListener('click', function(e) {
            if (e.target === tutorialOverlay) {
                closeTutorial();
            }
        });
        
        // Listen for escape key to close tutorial
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && tutorialOverlay.classList.contains('active')) {
                closeTutorial();
            }
        });
        
        // Previous button functionality
        if (prevBtn) {
            prevBtn.addEventListener('click', function() {
                navigateTutorial('prev');
            });
        }
        
        // Next button functionality
        if (nextBtn) {
            nextBtn.addEventListener('click', function() {
                navigateTutorial('next');
            });
        }
        
        // Progress dots functionality
        if (tutorialProgress) {
            const dots = tutorialProgress.querySelectorAll('.tutorial-dot');
            dots.forEach(dot => {
                dot.addEventListener('click', function() {
                    const stepNumber = parseInt(this.getAttribute('data-step'));
                    goToStep(stepNumber);
                });
            });
        }
        
        // Finish button functionality
        if (tutorialFinishBtn) {
            tutorialFinishBtn.addEventListener('click', function() {
                closeTutorial();
            });
        }
        
        // Show tutorial on first visit (only once)
        const hasSeenTutorial = localStorage.getItem('hasSeenTutorial');
        if (!hasSeenTutorial) {
            setTimeout(function() {
                openTutorial();
                localStorage.setItem('hasSeenTutorial', 'true');
            }, 1000);
        }
    }
    
    // Function to open the tutorial
    function openTutorial() {
        if (tutorialOverlay) {
            // Reset to first step
            goToStep(1);
            
            // Show overlay
            tutorialOverlay.classList.add('active');
            document.body.style.overflow = 'hidden';
        }
    }
    
    // Function to close the tutorial
    function closeTutorial() {
        if (tutorialOverlay) {
            tutorialOverlay.classList.remove('active');
            document.body.style.overflow = '';
            
            // Hide complete message if shown
            if (tutorialComplete && tutorialComplete.classList.contains('active')) {
                tutorialComplete.classList.remove('active');
                
                // Show steps again
                if (tutorialSteps) {
                    tutorialSteps.style.display = '';
                }
                
                // Show navigation
                if (document.getElementById('tutorialNavigation')) {
                    document.getElementById('tutorialNavigation').style.display = '';
                }
                
                // Show progress
                if (tutorialProgress) {
                    tutorialProgress.style.display = '';
                }
            }
        }
    }
    
    // Function to navigate between tutorial steps
    function navigateTutorial(direction) {
        if (!tutorialSteps) return;
        
        // Find current active step
        const activeStep = tutorialSteps.querySelector('.tutorial-step.active');
        if (!activeStep) return;
        
        const currentStep = parseInt(activeStep.getAttribute('data-step'));
        let newStep;
        
        if (direction === 'next') {
            newStep = currentStep + 1;
        } else {
            newStep = currentStep - 1;
        }
        
        // Get total steps
        const totalSteps = tutorialSteps.querySelectorAll('.tutorial-step').length;
        
        // Check if we're at the end
        if (newStep > totalSteps) {
            showCompletionMessage();
            return;
        }
        
        goToStep(newStep);
    }
    
    // Function to go to a specific step
    function goToStep(stepNumber) {
        if (!tutorialSteps) return;
        
        // Hide all steps
        const steps = tutorialSteps.querySelectorAll('.tutorial-step');
        steps.forEach(step => {
            step.classList.remove('active');
        });
        
        // Show the selected step
        const targetStep = tutorialSteps.querySelector(`.tutorial-step[data-step="${stepNumber}"]`);
        if (targetStep) {
            targetStep.classList.add('active');
        }
        
        // Update progress dots
        if (tutorialProgress) {
            const dots = tutorialProgress.querySelectorAll('.tutorial-dot');
            dots.forEach(dot => {
                const dotStep = parseInt(dot.getAttribute('data-step'));
                if (dotStep === stepNumber) {
                    dot.classList.add('active');
                } else {
                    dot.classList.remove('active');
                }
            });
        }
        
        // Update button states
        if (prevBtn) {
            if (stepNumber === 1) {
                prevBtn.disabled = true;
            } else {
                prevBtn.disabled = false;
            }
        }
        
        if (nextBtn) {
            const totalSteps = tutorialSteps.querySelectorAll('.tutorial-step').length;
            if (stepNumber === totalSteps) {
                nextBtn.textContent = 'Finish';
            } else {
                nextBtn.textContent = 'Next';
            }
        }
    }
    
    // Function to show completion message
    function showCompletionMessage() {
        if (!tutorialComplete) return;
        
        // Hide steps and navigation
        if (tutorialSteps) {
            tutorialSteps.style.display = 'none';
        }
        
        if (document.getElementById('tutorialNavigation')) {
            document.getElementById('tutorialNavigation').style.display = 'none';
        }
        
        if (tutorialProgress) {
            tutorialProgress.style.display = 'none';
        }
        
        // Show completion message
        tutorialComplete.classList.add('active');
    }
    // Get the form and button elements
    const uploadForm = document.getElementById('uploadForm');
    const processBtn = document.getElementById('processBtn');
    const documentInput = document.getElementById('document');
    const replacementsContainer = document.getElementById('replacements-container');
    const addReplacementBtn = document.getElementById('addReplacementBtn');
    
    // Drag and drop elements
    const documentDropzone = document.getElementById('documentDropzone');
    const browseBtn = document.getElementById('browseBtn');
    const filePreview = document.getElementById('filePreview');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const removeFile = document.getElementById('removeFile');
    
    let replacementCounter = 0;
    
    // Initialize drag and drop for the main document uploader
    if (documentDropzone && documentInput) {
        initDragAndDrop(documentDropzone, documentInput, filePreview, fileName, fileSize);
        
        if (browseBtn) {
            browseBtn.addEventListener('click', function() {
                documentInput.click();
            });
        }
        
        if (removeFile) {
            removeFile.addEventListener('click', function() {
                documentInput.value = '';
                filePreview.classList.add('d-none');
                documentDropzone.classList.remove('d-none');
            });
        }
        
        documentInput.addEventListener('change', function(e) {
            handleFileSelect(e, documentDropzone, filePreview, fileName, fileSize);
        });
    }
    
    // Initialize drag and drop for ELISA converter if on that page
    const outsideDocDropzone = document.getElementById('outsideDocDropzone');
    const outsideDocInput = document.getElementById('outside_document');
    const outsideFilePreview = document.getElementById('outsideFilePreview');
    const outsideFileName = document.getElementById('outsideFileName');
    const outsideFileSize = document.getElementById('outsideFileSize');
    const removeOutsideFile = document.getElementById('removeOutsideFile');
    
    const templateDocDropzone = document.getElementById('templateDocDropzone');
    const templateDocInput = document.getElementById('template_document');
    const templateFilePreview = document.getElementById('templateFilePreview');
    const templateFileName = document.getElementById('templateFileName');
    const templateFileSize = document.getElementById('templateFileSize');
    const removeTemplateFile = document.getElementById('removeTemplateFile');
    
    const browseBosterBtn = document.getElementById('browseBosterBtn');
    const browseTemplateBtn = document.getElementById('browseTemplateBtn');
    
    // Initialize ELISA converter drag and drop elements if they exist
    if (outsideDocDropzone && outsideDocInput) {
        initDragAndDrop(outsideDocDropzone, outsideDocInput, outsideFilePreview, outsideFileName, outsideFileSize);
        
        if (browseBosterBtn) {
            browseBosterBtn.addEventListener('click', function() {
                outsideDocInput.click();
            });
        }
        
        if (removeOutsideFile) {
            removeOutsideFile.addEventListener('click', function() {
                outsideDocInput.value = '';
                outsideFilePreview.classList.add('d-none');
                outsideDocDropzone.classList.remove('d-none');
            });
        }
        
        outsideDocInput.addEventListener('change', function(e) {
            handleFileSelect(e, outsideDocDropzone, outsideFilePreview, outsideFileName, outsideFileSize);
        });
    }
    
    if (templateDocDropzone && templateDocInput) {
        initDragAndDrop(templateDocDropzone, templateDocInput, templateFilePreview, templateFileName, templateFileSize);
        
        if (browseTemplateBtn) {
            browseTemplateBtn.addEventListener('click', function() {
                templateDocInput.click();
            });
        }
        
        if (removeTemplateFile) {
            removeTemplateFile.addEventListener('click', function() {
                templateDocInput.value = '';
                templateFilePreview.classList.add('d-none');
                templateDocDropzone.classList.remove('d-none');
            });
        }
        
        templateDocInput.addEventListener('change', function(e) {
            handleFileSelect(e, templateDocDropzone, templateFilePreview, templateFileName, templateFileSize);
        });
    }
    
    // Initialize drag and drop functionality
    function initDragAndDrop(dropzone, fileInput, previewElement, nameElement, sizeElement) {
        if (!dropzone || !fileInput) return;
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, highlight, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, unhighlight, false);
        });
        
        function highlight() {
            dropzone.classList.add('dragover');
        }
        
        function unhighlight() {
            dropzone.classList.remove('dragover');
        }
        
        dropzone.addEventListener('drop', handleDrop, false);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if (files.length > 0) {
                fileInput.files = files;
                // Trigger the change event
                const changeEvent = new Event('change', { bubbles: true });
                fileInput.dispatchEvent(changeEvent);
            }
        }
    }
    
    // Handle file selection (both from drag & drop and browse button)
    function handleFileSelect(e, dropzone, previewElement, nameElement, sizeElement) {
        const files = e.target.files;
        
        if (files.length > 0) {
            const file = files[0];
            
            // Update the UI
            if (nameElement) nameElement.textContent = file.name;
            if (sizeElement) sizeElement.textContent = formatFileSize(file.size);
            
            // Show the preview, hide the dropzone
            if (previewElement) previewElement.classList.remove('d-none');
            if (dropzone) dropzone.classList.add('d-none');
        }
    }
    
    // Format file size to human-readable format
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // Function to validate the form
    function validateForm() {
        // Check if file is selected
        if (!documentInput.files || documentInput.files.length === 0) {
            showError('Please select a DOCX file to upload');
            return false;
        }
        
        // Check file type
        const fileName = documentInput.files[0].name;
        const fileExt = fileName.split('.').pop().toLowerCase();
        if (fileExt !== 'docx') {
            showError('Only DOCX files are supported');
            return false;
        }
        
        // Check file size (max 16MB)
        const maxSize = 16 * 1024 * 1024; // 16MB in bytes
        if (documentInput.files[0].size > maxSize) {
            showError('File size exceeds the maximum limit of 16MB');
            return false;
        }
        
        // Check if at least one find text is provided
        const findTextInputs = document.querySelectorAll('.find-text');
        let hasValidFindText = false;
        
        findTextInputs.forEach(input => {
            if (input.value.trim()) {
                hasValidFindText = true;
            }
        });
        
        if (!hasValidFindText) {
            showError('Please enter at least one text to find');
            return false;
        }
        
        return true;
    }
    
    // Function to show error message
    function showError(message) {
        // Create a Bootstrap alert
        const alertDiv = document.createElement('div');
        alertDiv.className = 'alert alert-danger alert-dismissible fade show';
        alertDiv.setAttribute('role', 'alert');
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        `;
        
        if (uploadForm) {
            // Insert the alert before the form
            uploadForm.parentNode.insertBefore(alertDiv, uploadForm);
            
            // Scroll to the alert
            alertDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }
    
    // Function to add a new replacement row
    function addReplacementRow() {
        if (!replacementsContainer) return;
        
        replacementCounter++;
        
        const newRow = document.createElement('div');
        newRow.className = 'row replacement-row mb-3';
        newRow.innerHTML = `
            <div class="col-md-5">
                <label for="find_text_${replacementCounter}" class="form-label">Find Text</label>
                <input type="text" class="form-control find-text" id="find_text_${replacementCounter}" name="find_text[]" placeholder="Text to find" required>
            </div>
            <div class="col-md-5">
                <label for="replace_text_${replacementCounter}" class="form-label">Replace With</label>
                <input type="text" class="form-control replace-text" id="replace_text_${replacementCounter}" name="replace_text[]" placeholder="Text to replace with (leave empty to delete)">
            </div>
            <div class="col-md-2 d-flex align-items-end mb-2">
                <button type="button" class="btn btn-outline-danger remove-row">
                    <i class="fas fa-times"></i>
                </button>
            </div>
        `;
        
        replacementsContainer.appendChild(newRow);
        
        // Add event listener to the remove button
        const removeBtn = newRow.querySelector('.remove-row');
        removeBtn.addEventListener('click', function() {
            replacementsContainer.removeChild(newRow);
            updateRemoveButtons();
        });
        
        // Add event listener to input for validation
        const findTextInput = newRow.querySelector('.find-text');
        findTextInput.addEventListener('input', function() {
            if (this.value.trim()) {
                this.classList.add('is-valid');
                this.classList.remove('is-invalid');
            } else {
                this.classList.remove('is-valid');
                this.classList.add('is-invalid');
            }
        });
        
        updateRemoveButtons();
        return newRow;
    }
    
    // Function to update remove buttons visibility
    function updateRemoveButtons() {
        if (!replacementsContainer) return;
        
        const rows = replacementsContainer.querySelectorAll('.replacement-row');
        
        // Show remove buttons only if there's more than one row
        rows.forEach(row => {
            const removeBtn = row.querySelector('.remove-row');
            if (rows.length > 1) {
                removeBtn.style.display = 'block';
            } else {
                removeBtn.style.display = 'none';
            }
        });
    }
    
    // Add click event to the "Add Another Replacement" button
    if (addReplacementBtn) {
        addReplacementBtn.addEventListener('click', function() {
            const newRow = addReplacementRow();
            if (newRow) {
                newRow.querySelector('.find-text').focus();
            }
        });
    }
    
    // Show loading state when form is submitted
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(event) {
            if (!validateForm()) {
                event.preventDefault();
                return;
            }
            
            // Show loading state
            if (processBtn) {
                processBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Processing...';
                processBtn.disabled = true;
            }
            document.body.classList.add('processing');
        });
    }
    
    // Initialize the first row's validation
    const firstFindTextInput = document.querySelector('.find-text');
    if (firstFindTextInput) {
        firstFindTextInput.addEventListener('input', function() {
            if (this.value.trim()) {
                this.classList.add('is-valid');
                this.classList.remove('is-invalid');
            } else {
                this.classList.remove('is-valid');
                this.classList.add('is-invalid');
            }
        });
    }
    
    // Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert:not(.alert-info)');
    alerts.forEach(alert => {
        setTimeout(() => {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });
});
