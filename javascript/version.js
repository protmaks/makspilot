class VersionManager {
    constructor() {
        this.versionData = null;
        this.init();
    }

    async init() {
        await this.loadVersionData();
        this.updateVersionElements();
    }

    async loadVersionData() {
        try {
            const response = await fetch('/version.json');
            if (response.ok) {
                this.versionData = await response.json();
            } else {
                console.warn('Could not load version.json');
                this.versionData = { version: 'v 4.0 Beta' };
            }
        } catch (error) {
            console.warn('Error loading version data:', error);
            this.versionData = { version: 'v 4.0 Beta' };
        }
    }

    updateVersionElements() {
        if (!this.versionData) return;

        const versionElements = document.querySelectorAll('[data-version]');
        versionElements.forEach(element => {
            const versionType = element.getAttribute('data-version');
            
            switch (versionType) {
                case 'full':
                    element.textContent = this.versionData.version;
                    break;
                case 'number':
                    const numberMatch = this.versionData.version.match(/(\d+\.\d+)/);
                    element.textContent = numberMatch ? numberMatch[1] : this.versionData.version;
                    break;
                case 'date':
                    element.textContent = this.versionData.releaseDate || '';
                    break;
                default:
                    element.textContent = this.versionData.version;
            }
        });

        this.updateTitleVersions();
    }

    updateTitleVersions() {
        const h1Elements = document.querySelectorAll('h1');
        h1Elements.forEach(h1 => {
            const text = h1.textContent;
            const versionPattern = /v \d+\.\d+[^-]*/;
            if (versionPattern.test(text)) {
                h1.textContent = text.replace(versionPattern, this.versionData.version);
            }
        });
    }

    getVersion() {
        return this.versionData ? this.versionData.version : 'v 4.0 Beta';
    }

    getVersionNumber() {
        if (!this.versionData) return '4.0';
        const match = this.versionData.version.match(/(\d+\.\d+)/);
        return match ? match[1] : '4.0';
    }

    getReleaseDate() {
        return this.versionData ? this.versionData.releaseDate : '';
    }

    getChangelog() {
        return this.versionData ? this.versionData.changelog : {};
    }
}

let versionManager;

document.addEventListener('DOMContentLoaded', () => {
    versionManager = new VersionManager();
});

if (typeof module !== 'undefined' && module.exports) {
    module.exports = VersionManager;
}
