// Application Configuration Constants
// Centralized configuration for easy modification and documentation

export const CONFIG = {
    // Looper / Batch Processing Settings
    LOOPER: {
        MAX_CONCURRENT_REQUESTS: 5,
        BATCH_SIZE: 30,
        REQUEST_TIMEOUT_MS: 30000,
        RETRY_DELAY_MS: 1000
    },

    // Five9 Integration Settings
    FIVE9: {
        POLL_INTERVAL_MS: 2000,
        CONNECTION_CHECK_INTERVAL_MS: 5000
    },

    // UI Colors (semantic naming)
    COLORS: {
        SUCCESS: '#10b981',
        ERROR: '#ef4444',
        WARNING: '#f59e0b',
        MUTED: '#6b7280',
        PRIMARY: '#3b82f6',
        PURPLE: '#8b5cf6'
    },

    // Timing / Delays
    TIMING: {
        DEBOUNCE_MS: 300,
        TOAST_DURATION_MS: 3000,
        AUTO_CLOSE_DELAY_MS: 2000
    },

    // Cache Settings
    CACHE: {
        DEFAULT_TTL_MS: 24 * 60 * 60 * 1000, // 24 hours
        MAX_ENTRIES: 1000
    }
};
