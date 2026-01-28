// Encryption Utility - Encrypts sensitive data for storage
// Uses Web Crypto API with AES-GCM encryption
// Version: 1.0

/**
 * Derives an encryption key from the extension ID using PBKDF2.
 * The key is unique to this extension installation.
 * @returns {Promise<CryptoKey>} The derived encryption key
 */
async function deriveKey() {
    // Use extension ID as the base for key derivation
    // This ensures the key is unique per extension but consistent across sessions
    const extensionId = chrome.runtime.id;
    const encoder = new TextEncoder();

    // Import the extension ID as key material
    const keyMaterial = await crypto.subtle.importKey(
        'raw',
        encoder.encode(extensionId),
        'PBKDF2',
        false,
        ['deriveBits', 'deriveKey']
    );

    // Use a static salt (unique to this extension)
    const salt = encoder.encode('SRK-ChromeExt-v1-Salt');

    // Derive an AES-GCM key
    return crypto.subtle.deriveKey(
        {
            name: 'PBKDF2',
            salt: salt,
            iterations: 100000,
            hash: 'SHA-256'
        },
        keyMaterial,
        { name: 'AES-GCM', length: 256 },
        false,
        ['encrypt', 'decrypt']
    );
}

/**
 * Encrypts a string value using AES-GCM.
 * @param {string} plaintext - The string to encrypt
 * @returns {Promise<string>} Base64 encoded encrypted data (includes IV)
 */
export async function encrypt(plaintext) {
    if (!plaintext) return plaintext;

    try {
        const key = await deriveKey();
        const encoder = new TextEncoder();
        const data = encoder.encode(plaintext);

        // Generate a random IV for each encryption
        const iv = crypto.getRandomValues(new Uint8Array(12));

        // Encrypt the data
        const encryptedData = await crypto.subtle.encrypt(
            { name: 'AES-GCM', iv: iv },
            key,
            data
        );

        // Combine IV + encrypted data
        const combined = new Uint8Array(iv.length + encryptedData.byteLength);
        combined.set(iv, 0);
        combined.set(new Uint8Array(encryptedData), iv.length);

        // Return as base64 with a prefix to identify encrypted values
        return 'enc:' + btoa(String.fromCharCode(...combined));
    } catch (error) {
        console.error('[Encryption] Failed to encrypt:', error);
        throw error;
    }
}

/**
 * Decrypts a string value that was encrypted with encrypt().
 * @param {string} encryptedValue - Base64 encoded encrypted data (with 'enc:' prefix)
 * @returns {Promise<string>} The decrypted plaintext
 */
export async function decrypt(encryptedValue) {
    if (!encryptedValue) return encryptedValue;

    // Check if value is encrypted (has our prefix)
    if (!encryptedValue.startsWith('enc:')) {
        // Return as-is if not encrypted (backwards compatibility)
        return encryptedValue;
    }

    try {
        const key = await deriveKey();

        // Remove prefix and decode base64
        const base64Data = encryptedValue.substring(4);
        const combined = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));

        // Extract IV (first 12 bytes) and encrypted data
        const iv = combined.slice(0, 12);
        const encryptedData = combined.slice(12);

        // Decrypt the data
        const decryptedData = await crypto.subtle.decrypt(
            { name: 'AES-GCM', iv: iv },
            key,
            encryptedData
        );

        // Decode the plaintext
        const decoder = new TextDecoder();
        return decoder.decode(decryptedData);
    } catch (error) {
        console.error('[Encryption] Failed to decrypt:', error);
        // Return empty string on decryption failure (corrupted data)
        return '';
    }
}

/**
 * Checks if a value is encrypted (has our encryption prefix).
 * @param {string} value - The value to check
 * @returns {boolean} True if the value is encrypted
 */
export function isEncrypted(value) {
    return typeof value === 'string' && value.startsWith('enc:');
}
