(function() {
    'use strict';
    
    function isMobileDevice() {
        return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ||
               window.innerWidth <= 768;
    }
    
    function replaceHeroImages() {
        const heroImages = document.querySelectorAll('img[src*="main.webp"]');
        
        heroImages.forEach(img => {
            const currentSrc = img.src;
            const newSrc = currentSrc.replace('main.webp', 'main_mobile.webp');
            img.src = newSrc;
            img.fetchPriority = 'high';
        });
    }
    
    function updateOpenGraphImage() {
        const ogImageTag = document.querySelector('meta[property="og:image"]');
        if (ogImageTag) {
            const currentContent = ogImageTag.content;
            const newContent = currentContent.replace('main.webp', 'main_mobile.webp');
            ogImageTag.content = newContent;
        }
    }
    
    function initResponsiveImages() {
        if (isMobileDevice()) {
            replaceHeroImages();
            
            updateOpenGraphImage();
        }
    }
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initResponsiveImages);
    } else {
        initResponsiveImages();
    }
    
    window.addEventListener('resize', function() {
        const currentlyMobile = isMobileDevice();
        const htmlClass = document.documentElement.classList;
        const wasMobile = htmlClass.contains('mobile');
        
        if (currentlyMobile && !wasMobile) {
            htmlClass.add('mobile');
            replaceHeroImages();
            updateOpenGraphImage();
        } else if (!currentlyMobile && wasMobile) {
            htmlClass.remove('mobile');
        }
    });
})();