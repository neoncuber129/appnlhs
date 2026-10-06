const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const wwwrootDir = path.join(rootDir, 'ExcelDataEntryWeb', 'wwwroot');
const publishDir = path.join(rootDir, 'publish', 'local-web');

if (!fs.existsSync(publishDir)) {
    fs.mkdirSync(publishDir, { recursive: true });
}

// 1. Read CSS
const cssPath = path.join(wwwrootDir, 'css', 'style.css');
const cssContent = fs.readFileSync(cssPath, 'utf8');
fs.writeFileSync(path.join(publishDir, 'style.css'), cssContent, 'utf8');

// 2. Read SheetJS if present
let xlsxContent = '';
const xlsxPath = path.join(publishDir, 'xlsx.full.min.js');
if (fs.existsSync(xlsxPath)) {
    xlsxContent = fs.readFileSync(xlsxPath, 'utf8');
}

// 3. Process and Bundle JS files
const filesToBundle = [
    { name: 'debounce.js', path: path.join(wwwrootDir, 'js', 'utils', 'debounce.js') },
    { name: 'toast.js', path: path.join(wwwrootDir, 'js', 'utils', 'toast.js') },
    { name: 'storage.js', path: path.join(wwwrootDir, 'js', 'utils', 'storage.js') },
    { name: 'api.js', path: path.join(wwwrootDir, 'js', 'api.js') },
    { name: 'dropdown.js', path: path.join(wwwrootDir, 'js', 'ui', 'dropdown.js') },
    { name: 'sidebar.js', path: path.join(wwwrootDir, 'js', 'ui', 'sidebar.js') },
    { name: 'form.js', path: path.join(wwwrootDir, 'js', 'ui', 'form.js') },
    { name: 'modals.js', path: path.join(wwwrootDir, 'js', 'ui', 'modals.js') },
    { name: 'toolbar.js', path: path.join(wwwrootDir, 'js', 'ui', 'toolbar.js') },
    { name: 'app.js', path: path.join(wwwrootDir, 'js', 'app.js') }
];

let bundleJs = '/* Excel Data Entry Web - Standalone Local Bundle */\n';
bundleJs += '(function() {\n';
bundleJs += '  "use strict";\n\n';

for (const item of filesToBundle) {
    let content = fs.readFileSync(item.path, 'utf8');
    
    // Remove import lines
    content = content.replace(/^import\s+.*?from\s+['"].*?['"];?\r?\n/gm, '');
    
    // Remove export keywords (e.g. "export function", "export class", "export const")
    content = content.replace(/^export\s+(async\s+function|function|class|const|let|var)\s+/gm, '$1 ');
    content = content.replace(/^export\s+\{.*?\};?\r?\n/gm, '');
    content = content.replace(/^export\s+default\s+.*?;?\r?\n/gm, '');
    
    bundleJs += `  // --- MODULE: ${item.name} ---\n`;
    bundleJs += content.split('\n').map(line => '  ' + line).join('\n') + '\n\n';
}

bundleJs += '})();\n';

fs.writeFileSync(path.join(publishDir, 'app.bundle.js'), bundleJs, 'utf8');

// 4. Generate local index.html (multi-file local distribution)
const rawHtml = fs.readFileSync(path.join(wwwrootDir, 'index.html'), 'utf8');
let localHtml = rawHtml
    .replace('<link rel="stylesheet" href="css/style.css">', '<link rel="stylesheet" href="style.css">')
    .replace('<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>', '<script src="xlsx.full.min.js"></script>')
    .replace('<script type="module" src="js/app.js"></script>', '<script src="app.bundle.js"></script>');

fs.writeFileSync(path.join(publishDir, 'index.html'), localHtml, 'utf8');

// 5. Generate ExcelDataEntry_Offline.html (Single-file 100% self-contained)
let singleFileHtml = rawHtml
    .replace('<link rel="stylesheet" href="css/style.css">', `<style>\n${cssContent}\n</style>`)
    .replace('<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>', xlsxContent ? `<script>\n${xlsxContent}\n</script>` : '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>')
    .replace('<script type="module" src="js/app.js"></script>', `<script>\n${bundleJs}\n</script>`);

fs.writeFileSync(path.join(publishDir, 'ExcelDataEntry_Offline.html'), singleFileHtml, 'utf8');

console.log('Successfully published local web build to:', publishDir);
