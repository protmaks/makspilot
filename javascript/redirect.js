// JavaScript redirect for index.html to clean URLs
(function() {
    // Check if current URL contains index.html
    if (window.location.pathname.endsWith('/index.html')) {
        // Get the directory path without index.html
        var cleanPath = window.location.pathname.replace('/index.html', '/');
        
        // Redirect to clean URL
        window.location.replace(window.location.origin + cleanPath + window.location.search + window.location.hash);
    }
})();
