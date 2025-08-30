// Responsive Images Handler
// Detects mobile devices and replaces main.webp with main_mobile.webp

(function() {
    'use strict';
    
    // Function to detect if device is mobile
    function isMobileDevice() {
        return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ||
               window.innerWidth <= 768;
    }
    
    // Function to replace main.webp with main_mobile.webp in image tags
    function replaceHeroImages() {
        const heroImages = document.querySelectorAll('img[src*="main.webp"]');
        
        heroImages.forEach(img => {
            const currentSrc = img.src;
            const newSrc = currentSrc.replace('main.webp', 'main_mobile.webp');
            img.src = newSrc;
        });
    }
    
    // Function to update Open Graph meta tag
    function updateOpenGraphImage() {
        const ogImageTag = document.querySelector('meta[property="og:image"]');
        if (ogImageTag) {
            const currentContent = ogImageTag.content;
            const newContent = currentContent.replace('main.webp', 'main_mobile.webp');
            ogImageTag.content = newContent;
        }
    }
    
    // Initialize responsive images
    function initResponsiveImages() {
        if (isMobileDevice()) {
            // Replace hero images
            replaceHeroImages();
            
            // Update Open Graph image
            updateOpenGraphImage();
        }
    }
    
    // Run on DOMContentLoaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initResponsiveImages);
    } else {
        // DOM is already loaded
        initResponsiveImages();
    }
    
    // Also run on window resize to handle orientation changes
    window.addEventListener('resize', function() {
        // Only check if we're switching from desktop to mobile or vice versa
        const currentlyMobile = isMobileDevice();
        const htmlClass = document.documentElement.classList;
        const wasMobile = htmlClass.contains('mobile');
        
        if (currentlyMobile && !wasMobile) {
            htmlClass.add('mobile');
            replaceHeroImages();
            updateOpenGraphImage();
        } else if (!currentlyMobile && wasMobile) {
            htmlClass.remove('mobile');
            // Note: We don't revert to desktop images as it may cause layout shifts
        }
    });
})();