document.addEventListener('DOMContentLoaded', function() {
    // Find inputs that should trigger animations
    const textInputs = document.querySelectorAll('input[type="text"], input:not([type])');
    const passwordInputs = document.querySelectorAll('input[type="password"]');
    const submitButtons = document.querySelectorAll('input[type="submit"], button[type="submit"]');
    
    var current = null;
    
    // Text inputs animation
    textInputs.forEach(input => {
        input.addEventListener('focus', function(e) {
            if (current) current.pause();
            current = anime({
                targets: 'path',
                strokeDashoffset: {
                    value: 0,
                    duration: 700,
                    easing: 'easeOutQuart'
                },
                strokeDasharray: {
                    value: '240 1386',
                    duration: 700,
                    easing: 'easeOutQuart'
                }
            });
        });
    });
    
    // Password inputs animation
    passwordInputs.forEach(input => {
        input.addEventListener('focus', function(e) {
            if (current) current.pause();
            current = anime({
                targets: 'path',
                strokeDashoffset: {
                    value: -336,
                    duration: 700,
                    easing: 'easeOutQuart'
                },
                strokeDasharray: {
                    value: '240 1386',
                    duration: 700,
                    easing: 'easeOutQuart'
                }
            });
        });
    });
    
    // Submit buttons animation
    submitButtons.forEach(button => {
        button.addEventListener('focus', function(e) {
            if (current) current.pause();
            current = anime({
                targets: 'path',
                strokeDashoffset: {
                    value: -730,
                    duration: 700,
                    easing: 'easeOutQuart'
                },
                strokeDasharray: {
                    value: '530 1386',
                    duration: 700,
                    easing: 'easeOutQuart'
                }
            });
        });
    });
});