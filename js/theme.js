/**
 * Theme Management
 * Handles dark/light mode switching
 */

const Theme = {
    // Initialize theme from localStorage or system preference
    init() {
        const savedTheme = localStorage.getItem('theme');
        if (savedTheme) {
            document.documentElement.setAttribute('data-theme', savedTheme);
        } else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            document.documentElement.setAttribute('data-theme', 'dark');
        }
    },

    // Toggle between dark and light theme
    toggle() {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
    },

    // Get current theme
    getCurrent() {
        return document.documentElement.getAttribute('data-theme') || 'light';
    },

    // Set specific theme
    set(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('theme', theme);
    }
};

// Export for use in other modules
window.Theme = Theme;