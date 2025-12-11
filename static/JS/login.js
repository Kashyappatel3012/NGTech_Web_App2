        // Add animation to password field
        document.getElementById('password').addEventListener('focus', function() {
            this.parentNode.querySelector('.fa-lock').style.color = '#64ffda';
        });
        
        document.getElementById('password').addEventListener('blur', function() {
            this.parentNode.querySelector('.fa-lock').style.color = '#8892b0';
        });
        
        // Add animation to username field
        document.getElementById('username').addEventListener('focus', function() {
            this.parentNode.querySelector('i').style.color = '#64ffda';
        });
        
        document.getElementById('username').addEventListener('blur', function() {
            this.parentNode.querySelector('i').style.color = '#8892b0';
        });
        
        // Add cyber effect to login button
        const loginBtn = document.querySelector('.login-btn');
        loginBtn.addEventListener('mouseenter', function() {
            this.style.boxShadow = '0 5px 15px rgba(100, 255, 218, 0.4)';
        });
        
        loginBtn.addEventListener('mouseleave', function() {
            this.style.boxShadow = 'none';
        });
        
        // Add pulse animation to logo
        const logo = document.querySelector('.logo img');
        setInterval(() => {
            logo.style.transform = 'scale(1.05)';
            setTimeout(() => {
                logo.style.transform = 'scale(1)';
            }, 500);
        }, 3000);

        // Toggle password visibility
        const togglePassword = document.querySelector('#togglePassword');
        const password = document.querySelector('#password');
        
        togglePassword.addEventListener('click', function() {
            // Toggle the type attribute
            const type = password.getAttribute('type') === 'password' ? 'text' : 'password';
            password.setAttribute('type', type);
            
            // Toggle the eye icon
            this.querySelector('i').classList.toggle('fa-eye');
            this.querySelector('i').classList.toggle('fa-eye-slash');
            
            // Change color when active
            this.style.color = type === 'text' ? 'var(--accent)' : 'var(--text-secondary)';
        });

        // Browser Fingerprinting
        function getBrowserFingerprint() {
            const components = [];
            
            // User Agent
            components.push(navigator.userAgent || '');
            
            // Screen Resolution
            components.push(`${screen.width}x${screen.height}x${screen.colorDepth}`);
            
            // Timezone
            components.push(Intl.DateTimeFormat().resolvedOptions().timeZone || '');
            components.push(new Date().getTimezoneOffset().toString());
            
            // Language
            components.push(navigator.language || '');
            components.push((navigator.languages || []).join(','));
            
            // Platform
            components.push(navigator.platform || '');
            
            // Hardware Concurrency
            components.push(navigator.hardwareConcurrency?.toString() || '');
            
            // Device Memory (if available)
            components.push(navigator.deviceMemory?.toString() || '');
            
            // Canvas Fingerprint
            try {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = 200;
                canvas.height = 50;
                ctx.textBaseline = 'top';
                ctx.font = '14px Arial';
                ctx.textBaseline = 'alphabetic';
                ctx.fillStyle = '#f60';
                ctx.fillRect(125, 1, 62, 20);
                ctx.fillStyle = '#069';
                ctx.fillText('Browser fingerprint', 2, 15);
                ctx.fillStyle = 'rgba(102, 204, 0, 0.7)';
                ctx.fillText('Browser fingerprint', 4, 17);
                components.push(canvas.toDataURL());
            } catch (e) {
                components.push('canvas-error');
            }
            
            // WebGL Fingerprint
            try {
                const gl = document.createElement('canvas').getContext('webgl') || 
                          document.createElement('canvas').getContext('experimental-webgl');
                if (gl) {
                    const debugInfo = gl.getExtension('WEBGL_debug_renderer_info');
                    if (debugInfo) {
                        components.push(gl.getParameter(debugInfo.UNMASKED_VENDOR_WEBGL));
                        components.push(gl.getParameter(debugInfo.UNMASKED_RENDERER_WEBGL));
                    }
                }
            } catch (e) {
                components.push('webgl-error');
            }
            
            // Audio Context Fingerprint
            try {
                const audioContext = new (window.AudioContext || window.webkitAudioContext)();
                const oscillator = audioContext.createOscillator();
                const analyser = audioContext.createAnalyser();
                const gainNode = audioContext.createGain();
                const scriptProcessor = audioContext.createScriptProcessor(4096, 1, 1);
                
                gainNode.gain.value = 0;
                oscillator.connect(analyser);
                analyser.connect(scriptProcessor);
                scriptProcessor.connect(gainNode);
                gainNode.connect(audioContext.destination);
                oscillator.start(0);
                
                const fingerprint = analyser.frequencyData.length.toString();
                oscillator.stop();
                audioContext.close();
                components.push(fingerprint);
            } catch (e) {
                components.push('audio-error');
            }
            
            // Combine all components and create hash
            const fingerprintString = components.join('|');
            
            // Generate MD5 hash (same as main application for consistency)
            if (typeof CryptoJS !== 'undefined') {
                return CryptoJS.MD5(fingerprintString).toString();
            } else {
                // Fallback: simple hash if CryptoJS is not available
                let hash = 0;
                for (let i = 0; i < fingerprintString.length; i++) {
                    const char = fingerprintString.charCodeAt(i);
                    hash = ((hash << 5) - hash) + char;
                    hash = hash & hash; // Convert to 32bit integer
                }
                return Math.abs(hash).toString(16).padStart(32, '0');
            }
        }
        
        // Display browser fingerprint
        function displayBrowserFingerprint() {
            const fingerprintElement = document.getElementById('fingerprintValue');
            if (fingerprintElement) {
                try {
                    const fingerprint = getBrowserFingerprint();
                    fingerprintElement.textContent = fingerprint;
                    fingerprintElement.style.color = 'var(--accent)';
                } catch (error) {
                    fingerprintElement.textContent = 'Error generating fingerprint';
                    fingerprintElement.style.color = '#ff4757';
                }
            }
        }
        
        // Load CAPTCHA
        function loadCaptcha() {
            fetch('/captcha_image')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        const captchaImage = document.getElementById('captchaImage');
                        const captchaText = document.getElementById('captchaText');
                        
                        if (data.image) {
                            // Show image CAPTCHA
                            captchaImage.src = data.image;
                            captchaImage.style.display = 'block';
                            captchaText.style.display = 'none';
                        } else if (data.text) {
                            // Fallback: show text CAPTCHA
                            captchaText.textContent = data.text;
                            captchaText.style.display = 'block';
                            captchaImage.style.display = 'none';
                        }
                    }
                })
                .catch(error => {
                    console.error('Error loading CAPTCHA:', error);
                });
        }

        // Refresh CAPTCHA
        const refreshCaptchaBtn = document.getElementById('refreshCaptcha');
        if (refreshCaptchaBtn) {
            refreshCaptchaBtn.addEventListener('click', function() {
                loadCaptcha();
                // Clear CAPTCHA input
                const captchaInput = document.getElementById('captcha');
                if (captchaInput) {
                    captchaInput.value = '';
                }
            });
        }

        // Auto-uppercase CAPTCHA input
        const captchaInput = document.getElementById('captcha');
        if (captchaInput) {
            captchaInput.addEventListener('input', function() {
                this.value = this.value.toUpperCase();
            });
        }

        // Generate and display fingerprint when page loads
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function() {
                displayBrowserFingerprint();
                checkBrowserFingerprintAccess();
                loadCaptcha();
            });
        } else {
            displayBrowserFingerprint();
            checkBrowserFingerprintAccess();
            loadCaptcha();
        }
        
        // Check browser fingerprint access on page load via AJAX (not URL)
        function checkBrowserFingerprintAccess() {
            const fingerprint = getBrowserFingerprint();
            if (fingerprint) {
                // Hide page content immediately while validating
                const loginContainer = document.querySelector('.login-container');
                if (loginContainer) {
                    loginContainer.style.opacity = '0';
                }
                
                // Validate fingerprint via AJAX (not URL)
                fetch('/validate_fingerprint', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ browser_fingerprint: fingerprint })
                })
                .then(response => {
                    if (response.status === 404 || response.status === 403) {
                        // Invalid fingerprint - redirect to custom error page immediately
                        window.location.replace('/fingerprint_error');
                        return;
                    }
                    return response.json();
                })
                .then(data => {
                    if (!data) return; // Already handled redirect
                    
                    if (data.valid) {
                        // Fingerprint is valid, show page and allow it to function
                        if (loginContainer) {
                            loginContainer.style.opacity = '1';
                        }
                        
                        // Update hidden input for form submission
                        const fingerprintInput = document.getElementById('browserFingerprintInput');
                        if (fingerprintInput) {
                            fingerprintInput.value = fingerprint;
                        }
                        
                        // Store in sessionStorage for verify_otp page
                        sessionStorage.setItem('browser_fingerprint', fingerprint);
                    } else {
                        // Fingerprint is invalid - redirect to custom error page
                        window.location.replace('/fingerprint_error');
                    }
                })
                .catch(error => {
                    console.error('Error validating fingerprint:', error);
                    // On error, redirect to custom error page
                    window.location.replace('/fingerprint_error');
                });
            } else {
                // If fingerprint cannot be generated, redirect to custom error page
                window.location.replace('/fingerprint_error');
            }
        }
        
        
        // Add fingerprint to form on submit
        const loginForm = document.getElementById('loginForm');
        if (loginForm) {
            loginForm.addEventListener('submit', function(e) {
                const fingerprint = getBrowserFingerprint();
                const fingerprintInput = document.getElementById('browserFingerprintInput');
                if (fingerprintInput && fingerprint) {
                    fingerprintInput.value = fingerprint;
                } else if (fingerprint) {
                    // Create input if it doesn't exist
                    const hiddenInput = document.createElement('input');
                    hiddenInput.type = 'hidden';
                    hiddenInput.name = 'browser_fingerprint';
                    hiddenInput.value = fingerprint;
                    this.appendChild(hiddenInput);
                }
            });
        }

