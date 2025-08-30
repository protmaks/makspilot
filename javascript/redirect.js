(function() {
    if (window.location.pathname.endsWith('/index.html')) {
        var cleanPath = window.location.pathname.replace('/index.html', '/');
        
        window.location.replace(window.location.origin + cleanPath + window.location.search + window.location.hash);
    }
    
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', function() {
            navigator.serviceWorker.register('/sw.js')
                .then(function(registration) {
                    console.log('Service Worker registered with scope:', registration.scope);
                })
                .catch(function(error) {
                    console.log('Service Worker registration failed:', error);
                });
        });
    }
})();
