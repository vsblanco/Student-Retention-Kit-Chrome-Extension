/**
 * Simple markdown to HTML converter
 * Handles common markdown syntax for the About page
 */

/**
 * Converts markdown text to HTML
 * @param {string} markdown - The markdown text to convert
 * @returns {string} - The HTML output
 */
export function markdownToHtml(markdown) {
    let html = markdown;

    // Headers
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');

    // Bold
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

    // Horizontal rule
    html = html.replace(/^---$/gim, '<hr>');

    // Code (inline)
    html = html.replace(/`(.*?)`/g, '<code>$1</code>');

    // Unordered lists
    html = html.replace(/^\- (.*$)/gim, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');

    // Line breaks (paragraphs)
    const lines = html.split('\n');
    const processedLines = [];
    let inList = false;
    let inBlock = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Skip empty lines
        if (!line) {
            if (inList) {
                processedLines.push('</ul>');
                inList = false;
            }
            inBlock = false;
            continue;
        }

        // Check if line is already wrapped in a tag
        if (line.startsWith('<h') || line.startsWith('<hr') || line.startsWith('<ul') || line.startsWith('</ul>')) {
            if (line.startsWith('<ul>')) {
                inList = true;
            } else if (line.startsWith('</ul>')) {
                inList = false;
            }
            processedLines.push(line);
            inBlock = true;
        } else if (line.startsWith('<li>')) {
            if (!inList) {
                processedLines.push('<ul>');
                inList = true;
            }
            processedLines.push(line);
            inBlock = true;
        } else {
            // Regular text - wrap in paragraph
            if (!inBlock) {
                processedLines.push('<p>' + line + '</p>');
            } else {
                processedLines.push(line);
            }
        }
    }

    // Close any open lists
    if (inList) {
        processedLines.push('</ul>');
    }

    return processedLines.join('\n');
}

/**
 * Loads and renders a markdown file into the specified container
 * @param {string} filePath - Path to the markdown file
 * @param {HTMLElement} container - The container element to render into
 */
export async function loadAndRenderMarkdown(filePath, container) {
    try {
        const response = await fetch(filePath);
        if (!response.ok) {
            throw new Error(`Failed to load markdown file: ${response.statusText}`);
        }

        const markdown = await response.text();
        const html = markdownToHtml(markdown);
        container.innerHTML = html;

        console.log('✓ Markdown rendered successfully');
    } catch (error) {
        console.error('Error loading markdown:', error);
        container.innerHTML = '<p style="color: #ef4444;">Failed to load content. Please try again.</p>';
    }
}
