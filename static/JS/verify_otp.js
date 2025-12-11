
        // Countdown timer
        function startTimer(duration, display) {
            let timer = duration, minutes, seconds;
            const interval = setInterval(function () {
                minutes = parseInt(timer / 60, 10);
                seconds = parseInt(timer % 60, 10);

                minutes = minutes < 10 ? "0" + minutes : minutes;
                seconds = seconds < 10 ? "0" + seconds : seconds;

                display.textContent = minutes + ":" + seconds;

                if (--timer < 0) {
                    clearInterval(interval);
                    display.textContent = "00:00";
                    // Use textContent instead of innerHTML to prevent XSS
                    const timerElement = document.querySelector('.timer');
                    if (timerElement) {
                        timerElement.textContent = '';
                        const span = document.createElement('span');
                        span.style.color = '#ff4757';
                        span.textContent = 'OTP has expired. Please request a new one.';
                        timerElement.appendChild(span);
                    }
                }
            }, 1000);
        }

        // Add animation to OTP field
        document.getElementById('otp').addEventListener('focus', function() {
            this.parentNode.querySelector('i').style.color = '#64ffda';
        });
        
        document.getElementById('otp').addEventListener('blur', function() {
            this.parentNode.querySelector('i').style.color = '#8892b0';
        });
        
        // Add cyber effect to verify button
        const verifyBtn = document.querySelector('.verify-btn');
        verifyBtn.addEventListener('mouseenter', function() {
            this.style.boxShadow = '0 5px 15px rgba(100, 255, 218, 0.4)';
        });
        
        verifyBtn.addEventListener('mouseleave', function() {
            this.style.boxShadow = 'none';
        });
        
        // Add pulse animation to timer
        const timerElement = document.querySelector('.timer');
        setInterval(() => {
            timerElement.style.animation = 'pulse 2s infinite';
        }, 4000);
        
        // Browser Fingerprinting for OTP verification
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
        
        // Get fingerprint from sessionStorage or generate it (not from URL for security)
        function getStoredFingerprint() {
            // Try sessionStorage first (set during login)
            let fingerprint = sessionStorage.getItem('browser_fingerprint');
            if (fingerprint) {
                return fingerprint;
            }
            
            // Last resort: generate new fingerprint
            fingerprint = getBrowserFingerprint();
            sessionStorage.setItem('browser_fingerprint', fingerprint);
            return fingerprint;
        }
        
        // Add fingerprint to form on page load (not URL for security)
        window.onload = function () {
            // Get the remaining time from server-side calculation
            const timerText = document.querySelector('#time').textContent;
            const [minutes, seconds] = timerText.split(':').map(Number);
            const totalSeconds = minutes * 60 + seconds;
            
            const display = document.querySelector('#time');
            
            if (totalSeconds > 0) {
                startTimer(totalSeconds, display);
            } else {
                display.textContent = "00:00";
                // Use textContent instead of innerHTML to prevent XSS
                const timerElement = document.querySelector('.timer');
                if (timerElement) {
                    timerElement.textContent = '';
                    const span = document.createElement('span');
                    span.style.color = '#ff4757';
                    span.textContent = 'OTP has expired. Please request a new one.';
                    timerElement.appendChild(span);
                }
            }
            
            // Get fingerprint from sessionStorage (not URL for security)
            const fingerprint = getStoredFingerprint();
            if (fingerprint) {
                // Add to form (fingerprint already validated and stored in session)
                const fingerprintInput = document.getElementById('browserFingerprintInput');
                if (fingerprintInput) {
                    fingerprintInput.value = fingerprint;
                }
            }
        };
        
        // Add fingerprint to form on submit
        const otpForm = document.getElementById('otpForm');
        if (otpForm) {
            otpForm.addEventListener('submit', function(e) {
                const fingerprint = getStoredFingerprint();
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