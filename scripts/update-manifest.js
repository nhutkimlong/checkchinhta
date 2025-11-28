const fs = require('fs');
const path = require('path');

const manifestPath = path.join(__dirname, '..', 'manifest.xml');
const devUrl = 'https://localhost:3000';
const prodUrl = 'https://checkchinhta.netlify.app';

const mode = process.argv[2]; // 'dev' or 'prod'

if (!mode || (mode !== 'dev' && mode !== 'prod')) {
  console.error('Usage: node update-manifest.js [dev|prod]');
  process.exit(1);
}

try {
  let content = fs.readFileSync(manifestPath, 'utf8');
  let newContent = content;

  if (mode === 'prod') {
    console.log('Updating manifest.xml to PRODUCTION URLs...');
    // Replace devUrl with prodUrl
    // Use a global regex to replace all occurrences
    const regex = new RegExp(devUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    newContent = content.replace(regex, prodUrl);
    
    // Update DisplayName to remove (Dev) if present
    newContent = newContent.replace(/<DisplayName DefaultValue="AI Check Chính Tả \(Dev\)"\/>/, '<DisplayName DefaultValue="AI Check Chính Tả"/>');
    
  } else {
    console.log('Updating manifest.xml to DEVELOPMENT URLs...');
    // Replace prodUrl with devUrl
    const regex = new RegExp(prodUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    newContent = content.replace(regex, devUrl);

    // Update DisplayName to add (Dev)
    if (!newContent.includes('AI Check Chính Tả (Dev)')) {
        newContent = newContent.replace(/<DisplayName DefaultValue="AI Check Chính Tả"\/>/, '<DisplayName DefaultValue="AI Check Chính Tả (Dev)"/>');
    }
  }

  if (content !== newContent) {
    fs.writeFileSync(manifestPath, newContent, 'utf8');
    console.log('manifest.xml updated successfully.');
  } else {
    console.log('manifest.xml is already up to date.');
  }

} catch (err) {
  console.error('Error updating manifest.xml:', err);
  process.exit(1);
}
