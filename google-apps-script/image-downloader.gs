// Enhanced Google Apps Script - Implementing Research Findings
// Based on deep research into NHK anti-bot systems and UrlFetchApp limitations
// 
// Key Improvements:
// 1. Manual redirect handling
// 2. Missing headers (Accept-Encoding, etc.)
// 3. Exponential backoff for blocked requests
// 4. Better error detection and logging
// 5. User-Agent rotation

// Configuration
const CONFIG = {
  MAIN_DRIVE_FOLDER_ID: '1D-1xQifHDt8AVz83bZwCHiYugft_BFYq'
}

// Process remaining images (skip already processed and known problematic ones)
function processRemainingImagesOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const thumbnailCol = headers.indexOf('Thumbnail');
  const headlineCol = headers.indexOf('Headline');
  const publishCol = headers.indexOf('Publish date');
  const linkCol = headers.indexOf('Link');
  const driveUrlCol = headers.indexOf('Drive Image URL') !== -1 ? headers.indexOf('Drive Image URL') : headers.length;
  const statusCol = headers.indexOf('Image Status') !== -1 ? headers.indexOf('Image Status') : headers.length + 1;
  
  console.log('🎯 Processing ONLY remaining unprocessed images...');
  
  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let manualDownload = 0;
  
  // First, mark all problematic URLs for manual download
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const thumbnailUrl = row[thumbnailCol];
    const headline = row[headlineCol];
    const existingDriveUrl = row[driveUrlCol];
    const existingStatus = row[statusCol];
    
    // Skip if already processed
    if (!thumbnailUrl || existingDriveUrl || existingStatus) {
      continue;
    }
    
    // Check if should be skipped and mark for manual download
    if (shouldSkipUrl(thumbnailUrl)) {
      const seoFilename = createSEOFilename(headline);
      sheet.getRange(i + 1, driveUrlCol + 1).setValue(`📝 USE FILENAME: ${seoFilename}`);
      sheet.getRange(i + 1, statusCol + 1).setValue('🚫 Auto-Skip → Manual');
      console.log(`⏭️ Auto-skipped row ${i + 1}: ${headline.substring(0, 50)}...`);
      manualDownload++;
    }
  }
  
  console.log(`📝 Auto-marked ${manualDownload} known problematic URLs for manual download`);
  
  // Now process the remaining ones
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const thumbnailUrl = row[thumbnailCol];
    const headline = row[headlineCol];
    const publishDate = new Date(row[publishCol]);
    const articleUrl = row[linkCol];
    const existingDriveUrl = row[driveUrlCol];
    const existingStatus = row[statusCol];
    
    // Skip if already processed, has status, or should be skipped
    if (!thumbnailUrl || existingDriveUrl || existingStatus || shouldSkipUrl(thumbnailUrl)) {
      skipped++;
      continue;
    }
    
    console.log(`\n📸 Processing row ${i + 1}: ${headline}`);
    
    try {
      const startTime = new Date().getTime();
      
      // Use enhanced download function
      const imageBlob = downloadImageEnhanced(thumbnailUrl, articleUrl);
      const seoFilename = createSEOFilename(headline);
      const driveFile = saveToDrive(imageBlob, seoFilename, publishDate);
      const shareableUrl = getShareableUrl(driveFile);
      
      const endTime = new Date().getTime();
      console.log(`⏱️ Success in ${(endTime - startTime) / 1000}s`);
      
      // Update spreadsheet
      sheet.getRange(i + 1, driveUrlCol + 1).setValue(shareableUrl);
      sheet.getRange(i + 1, statusCol + 1).setValue('✅ Enhanced');
      processed++;
      
    } catch (error) {
      console.error(`❌ Failed row ${i + 1}: ${error.message}`);
      
      // Handle skip errors
      if (error.message.includes('SKIP')) {
        const seoFilename = createSEOFilename(headline);
        sheet.getRange(i + 1, driveUrlCol + 1).setValue(`📝 USE FILENAME: ${seoFilename}`);
        
        let errorStatus = '⚠️ Issue → Manual';
        if (error.message.includes('TIMEOUT_SKIP')) errorStatus = '⏰ Timeout → Manual';
        else if (error.message.includes('BLOCKED_SKIP')) errorStatus = '🚫 Blocked → Manual';
        else if (error.message.includes('JS_CHALLENGE_SKIP')) errorStatus = '🔒 JS Challenge → Manual';
        
        sheet.getRange(i + 1, statusCol + 1).setValue(errorStatus);
        manualDownload++;
      } else {
        sheet.getRange(i + 1, statusCol + 1).setValue('❌ Technical Error');
        failed++;
      }
    }
    
    // Short delay
    Utilities.sleep(2000);
    
    // Progress updates
    if (i % 10 === 0) {
      console.log(`📊 Progress: ${processed} success, ${manualDownload} manual, ${failed} failed`);
    }
  }
  
  console.log('\n🎉 Remaining images processed!');
  console.log(`✅ Successfully processed: ${processed}`);
  console.log(`📝 Marked for manual download: ${manualDownload}`);
  console.log(`❌ Technical failures: ${failed}`);
  console.log(`⏭️ Skipped (already done): ${skipped}`);
};

// User-Agent rotation pool (research shows this helps)
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
];

// Get random User-Agent
function getRandomUserAgent() {
  return USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
}

// Enhanced download function with aggressive timeout and skip mechanism
function downloadImageEnhanced(imageUrl, articleUrl = null, attempt = 1) {
  console.log(`🔄 Download attempt ${attempt} for: ${imageUrl}`);
  
  // Check if this URL is in our skip list (persistent problematic URLs)
  if (shouldSkipUrl(imageUrl)) {
    console.log('⏭️ URL in skip list - marking for manual download');
    throw new Error('SKIP_URL');
  }
  
  // Get optimal referer
  const referer = getBestReferer(imageUrl, articleUrl);
  console.log(`🔗 Using referer: ${referer}`);
  
  // Build enhanced headers based on research findings
  const headers = {
    // Rotating User-Agent to avoid detection
    'User-Agent': getRandomUserAgent(),
    
    // Critical missing headers identified in research
    'Accept': 'image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
    'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br', // ← This was missing!
    
    // Referer for authenticity
    'Referer': referer,
    
    // Additional headers from research
    'Cache-Control': 'no-cache',
    'DNT': '1', // Do Not Track
    'Upgrade-Insecure-Requests': '1',
    
    // Sec-Fetch headers (adjusted for Japanese sites)
    'Sec-Fetch-Dest': 'image',
    'Sec-Fetch-Mode': 'no-cors',
    'Sec-Fetch-Site': imageUrl.includes(referer.split('//')[1]?.split('/')[0]) ? 'same-site' : 'cross-site'
  };
  
  // More aggressive timeout for hanging URLs
  const timeoutMs = attempt === 1 ? 20000 : 15000; // Shorter on retries
  
  // Enhanced options based on research
  const options = {
    method: 'GET',
    headers: headers,
    followRedirects: false, // ← Manual redirect handling
    muteHttpExceptions: true, // ← Better error info
    validateHttpsCertificates: false,
    timeout: timeoutMs // ← Aggressive timeout
  };
  
  try {
    console.log(`📡 Making request with ${timeoutMs}ms timeout...`);
    
    // Add timeout wrapper to detect hanging requests
    const startTime = new Date().getTime();
    let response;
    
    try {
      response = UrlFetchApp.fetch(imageUrl, options);
    } catch (fetchError) {
      const elapsed = new Date().getTime() - startTime;
      
      // If it took too long, treat as hanging
      if (elapsed >= timeoutMs - 1000 || fetchError.message.includes('timeout')) {
        console.log(`⏰ Request hung for ${elapsed}ms - marking for skip`);
        throw new Error('TIMEOUT_SKIP');
      }
      
      throw fetchError;
    }
    
    let responseCode = response.getResponseCode();
    console.log(`📊 Initial response: ${responseCode}`);
    
    // Handle redirects manually (research shows this is critical)
    let redirectCount = 0;
    const maxRedirects = 3; // Reduced to avoid long chains
    
    while ([301, 302, 303, 307, 308].includes(responseCode) && redirectCount < maxRedirects) {
      const location = response.getHeaders()['Location'] || response.getHeaders()['location'];
      
      if (!location) {
        throw new Error(`Redirect response ${responseCode} but no Location header`);
      }
      
      console.log(`🔄 Following redirect ${redirectCount + 1}: ${location}`);
      
      // Check redirect timeout
      const redirectStart = new Date().getTime();
      response = UrlFetchApp.fetch(location, {...options, timeout: 15000}); // Shorter timeout for redirects
      const redirectTime = new Date().getTime() - redirectStart;
      
      if (redirectTime >= 14000) {
        console.log(`⏰ Redirect hung for ${redirectTime}ms - marking for skip`);
        throw new Error('REDIRECT_TIMEOUT_SKIP');
      }
      
      responseCode = response.getResponseCode();
      redirectCount++;
      
      Utilities.sleep(500);
    }
    
    // Handle Japanese site blocking patterns
    if (responseCode === 403) {
      console.log('🚫 403 Forbidden - Site blocking detected');
      
      if (attempt <= 2) { // Reduced retries to avoid long waits
        const delay = Math.pow(2, attempt) * 1500; // Shorter delays: 3s, 6s
        console.log(`⏳ Waiting ${delay}ms before retry...`);
        Utilities.sleep(delay);
        
        return downloadImageEnhanced(imageUrl, articleUrl, attempt + 1);
      } else {
        throw new Error('BLOCKED_SKIP');
      }
    }
    
    if (responseCode === 429) {
      console.log('⏰ 429 Rate Limited');
      throw new Error('RATE_LIMITED_SKIP'); // Don't retry rate limits, skip immediately
    }
    
    // Check for JavaScript challenge (research finding)
    if (responseCode === 200) {
      const contentType = response.getHeaders()['Content-Type'] || response.getHeaders()['content-type'] || '';
      
      if (contentType.includes('text/html')) {
        console.log('⚠️ Received HTML instead of image - possible JavaScript challenge');
        throw new Error('JS_CHALLENGE_SKIP');
      }
    }
    
    if (responseCode !== 200) {
      throw new Error(`HTTP_${responseCode}_SKIP`);
    }
    
    return validateAndReturnBlob(response);
    
  } catch (error) {
    console.error(`❌ Download failed (attempt ${attempt}): ${error.message}`);
    
    // Skip logic - don't retry certain errors
    if (error.message.includes('SKIP')) {
      throw error; // Let the caller handle skip errors
    }
    
    // Quick retry for network errors only
    if (attempt === 1 && (error.message.includes('network') || error.message.includes('connection'))) {
      console.log('🔄 Quick retry for network error...');
      Utilities.sleep(2000);
      return downloadImageEnhanced(imageUrl, articleUrl, 2);
    }
    
    throw new Error('NETWORK_ERROR_SKIP');
  }
}

// Validate response and return blob
function validateAndReturnBlob(response) {
  const blob = response.getBlob();
  const sizeInMB = (blob.getBytes().length / 1024 / 1024).toFixed(2);
  const contentType = blob.getContentType() || '';
  
  console.log(`📦 Response: ${sizeInMB} MB, Type: ${contentType}`);
  
  // Validate it's actually an image
  if (!contentType.startsWith('image/')) {
    // Sometimes content-type is wrong, check by size and magic bytes
    if (blob.getBytes().length < 1000) {
      throw new Error('Response too small - likely error page');
    }
    
    // Check magic bytes for common image formats
    const bytes = blob.getBytes();
    const isJPEG = bytes[0] === 0xFF && bytes[1] === 0xD8;
    const isPNG = bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47;
    const isWebP = bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50;
    
    if (!isJPEG && !isPNG && !isWebP) {
      throw new Error(`Invalid image data - Content-Type: ${contentType}`);
    }
    
    console.log('✅ Image validated by magic bytes');
  }
  
  if (blob.getBytes().length < 1000) {
    throw new Error('Image too small - likely error image');
  }
  
  console.log(`✅ Valid image: ${sizeInMB} MB`);
  return blob;
}

// Skip list for known problematic URLs/domains
function shouldSkipUrl(imageUrl) {
  const problematicPatterns = [
    // Known problematic domains that cause hangs
    'nhk.or.jp', // NHK always hangs with JavaScript challenges
    'japantimes.co.jp/wp-content', // Japan Times images often protected
    // Add more patterns as needed
  ];
  
  return problematicPatterns.some(pattern => imageUrl.includes(pattern));
}

// Add URL to skip list (for persistent problems)
function addToSkipList(imageUrl) {
  console.log(`📝 Adding to skip list: ${imageUrl}`);
  // In a real implementation, you could store this in PropertiesService
  // PropertiesService.getScriptProperties().setProperty('skipUrls', JSON.stringify(skipUrls));
}
function getBestReferer(imageUrl, articleUrl) {
  if (articleUrl && articleUrl.startsWith('http')) {
    return articleUrl;
  }
  
  // Site-specific referer mapping
  const url = new URL(imageUrl);
  const domain = url.hostname.toLowerCase();
  
  const refererMap = {
    'www3.nhk.or.jp': 'https://www3.nhk.or.jp/',
    'nhk.or.jp': 'https://www.nhk.or.jp/',
    'japantimes.co.jp': 'https://www.japantimes.co.jp/',
    'tokyoweekender.com': 'https://www.tokyoweekender.com/',
    'designboom.com': 'https://www.designboom.com/',
    'timeout.com': 'https://www.timeout.com/',
  };
  
  for (const [key, value] of Object.entries(refererMap)) {
    if (domain.includes(key)) {
      return value;
    }
  }
  
  return `https://${domain}/`;
}

// Enhanced main processing function
function processAllImagesEnhanced() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const thumbnailCol = headers.indexOf('Thumbnail');
  const headlineCol = headers.indexOf('Headline');
  const publishCol = headers.indexOf('Publish date');
  const linkCol = headers.indexOf('Link');
  const driveUrlCol = headers.indexOf('Drive Image URL') !== -1 ? headers.indexOf('Drive Image URL') : headers.length;
  const statusCol = headers.indexOf('Image Status') !== -1 ? headers.indexOf('Image Status') : headers.length + 1;
  
  // Add headers if needed
  if (headers.indexOf('Drive Image URL') === -1) {
    sheet.getRange(1, driveUrlCol + 1).setValue('Drive Image URL');
  }
  if (headers.indexOf('Image Status') === -1) {
    sheet.getRange(1, statusCol + 1).setValue('Image Status');
  }
  
  console.log(`🚀 Enhanced processing of ${data.length - 1} articles...`);
  
  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let manualDownload = 0; // Track URLs marked for manual download
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const thumbnailUrl = row[thumbnailCol];
    const headline = row[headlineCol];
    const publishDate = new Date(row[publishCol]);
    const articleUrl = row[linkCol];
    const existingDriveUrl = row[driveUrlCol];
    
    // Skip if already processed
    if (!thumbnailUrl || existingDriveUrl) {
      console.log(`⏭️ Skipping row ${i + 1}: already processed or no URL`);
      skipped++;
      continue;
    }
    
    console.log(`\n📸 Processing row ${i + 1}: ${headline}`);
    
    try {
      const startTime = new Date().getTime();
      
      // Use enhanced download function
      const imageBlob = downloadImageEnhanced(thumbnailUrl, articleUrl);
      const seoFilename = createSEOFilename(headline);
      const driveFile = saveToDrive(imageBlob, seoFilename, publishDate);
      const shareableUrl = getShareableUrl(driveFile);
      
      const endTime = new Date().getTime();
      console.log(`⏱️ Success in ${(endTime - startTime) / 1000}s`);
      
      // Update spreadsheet
      sheet.getRange(i + 1, driveUrlCol + 1).setValue(shareableUrl);
      sheet.getRange(i + 1, statusCol + 1).setValue('✅ Enhanced');
      processed++;
      
    } catch (error) {
      console.error(`❌ Failed row ${i + 1}: ${error.message}`);
      
      // Handle different skip reasons
      let errorStatus = '❌ Failed';
      let shouldContinue = true;
      
      if (error.message.includes('SKIP')) {
        // These are URLs we should skip and mark for manual download
        manualDownload++;
        
        // Generate the SEO filename they should use for manual download
        const seoFilename = createSEOFilename(headline);
        
        if (error.message.includes('TIMEOUT_SKIP') || error.message.includes('REDIRECT_TIMEOUT_SKIP')) {
          errorStatus = '⏰ Timeout → Manual';
          console.log('⏰ Request timed out - marking for manual download');
        } else if (error.message.includes('BLOCKED_SKIP')) {
          errorStatus = '🚫 Blocked → Manual';
          console.log('🚫 Site blocked - marking for manual download');
        } else if (error.message.includes('JS_CHALLENGE_SKIP')) {
          errorStatus = '🔒 JS Challenge → Manual';
          console.log('🔒 JavaScript challenge - marking for manual download');
        } else if (error.message.includes('RATE_LIMITED_SKIP')) {
          errorStatus = '🐌 Rate Limited → Manual';
          console.log('🐌 Rate limited - marking for manual download');
        } else if (error.message.includes('SKIP_URL')) {
          errorStatus = '⏭️ Skip List → Manual';
          console.log('⏭️ In skip list - marking for manual download');
        } else {
          errorStatus = '⚠️ Issue → Manual';
        }
        
        // Put the SEO filename in the Drive Image URL column for manual download
        sheet.getRange(i + 1, driveUrlCol + 1).setValue(`📝 USE FILENAME: ${seoFilename}`);
        console.log(`📝 Manual download filename: ${seoFilename}`);
        
      } else {
        // Regular failures - don't put filename since these are technical errors
        failed++;
        if (error.message.includes('Drive save')) {
          errorStatus = '💾 Drive Error';
        } else if (error.message.includes('timeout')) {
          errorStatus = '⏰ Network Timeout';
        }
      }
      
      sheet.getRange(i + 1, statusCol + 1).setValue(errorStatus);
    }
    
    // Shorter delay to speed up processing
    console.log('⏸️ Waiting 2 seconds...');
    Utilities.sleep(2000);
    
    // Progress updates
    if (i % 10 === 0) {
      console.log(`\n📊 Progress: ${i}/${data.length - 1} rows processed`);
      console.log(`✅ Success: ${processed}, 📝 Manual: ${manualDownload}, ❌ Failed: ${failed}, ⏭️ Skipped: ${skipped}`);
    }
  }
  
  console.log('\n🎉 Enhanced processing complete!');
  console.log(`📊 Final Results:`);
  console.log(`✅ Successfully processed: ${processed}`);
  console.log(`📝 Marked for manual download: ${manualDownload}`);
  console.log(`❌ Technical failures: ${failed}`);
  console.log(`⏭️ Already processed/skipped: ${skipped}`);
  
  if (manualDownload > 0) {
    console.log(`\n💡 ${manualDownload} images marked for manual download:`);
    console.log('- Check "Drive Image URL" column for filename to use');
    console.log('- Download manually and rename to the specified filename');
    console.log('- Upload to Drive and replace cell with Drive URL');
    console.log('- Run listManualDownloadUrls() for detailed instructions');
  }
}

// Test with enhanced approach
function testFirstRowEnhanced() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const thumbnailCol = headers.indexOf('Thumbnail');
  const headlineCol = headers.indexOf('Headline');
  const publishCol = headers.indexOf('Publish date');
  const linkCol = headers.indexOf('Link');
  const driveUrlCol = headers.indexOf('Drive Image URL') !== -1 ? headers.indexOf('Drive Image URL') : headers.length;
  const statusCol = headers.indexOf('Image Status') !== -1 ? headers.indexOf('Image Status') : headers.length + 1;
  
  console.log('🧪 Testing enhanced approach on first row...');
  
  if (data.length < 2) {
    console.log('❌ No data rows found');
    return;
  }
  
  const row = data[1];
  const thumbnailUrl = row[thumbnailCol];
  const headline = row[headlineCol];
  const publishDate = new Date(row[publishCol]);
  const articleUrl = row[linkCol];
  
  console.log(`📸 Testing: ${headline}`);
  console.log(`🔗 Image: ${thumbnailUrl}`);
  console.log(`📄 Article: ${articleUrl}`);
  
  try {
    const imageBlob = downloadImageEnhanced(thumbnailUrl, articleUrl);
    const seoFilename = createSEOFilename(headline);
    const driveFile = saveToDrive(imageBlob, seoFilename, publishDate);
    const shareableUrl = getShareableUrl(driveFile);
    
    sheet.getRange(2, driveUrlCol + 1).setValue(shareableUrl);
    sheet.getRange(2, statusCol + 1).setValue('✅ Enhanced');
    
    console.log('🎉 Enhanced approach SUCCESS!');
    console.log(`📁 Drive URL: ${shareableUrl}`);
    
  } catch (error) {
    console.error('❌ Enhanced approach failed:', error.message);
    
    // If it's a skip error, show the filename that would be used
    if (error.message.includes('SKIP')) {
      const seoFilename = createSEOFilename(headline);
      sheet.getRange(2, driveUrlCol + 1).setValue(`📝 USE FILENAME: ${seoFilename}`);
      sheet.getRange(2, statusCol + 1).setValue(`❌ ${error.message.replace('_SKIP', '')} → Manual`);
      
      console.log('📝 Manual download required');
      console.log(`👉 Download URL: ${thumbnailUrl}`);
      console.log(`👉 Save as filename: ${seoFilename}`);
    } else {
      sheet.getRange(2, statusCol + 1).setValue(`❌ ${error.message.substring(0, 30)}`);
    }
  }
}

// All the existing helper functions remain the same
function createSEOFilename(headline) {
  if (!headline) return `image-${Date.now()}.jpg`;
  
  return headline
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .substring(0, 50)
    .replace(/^-|-$/g, '')
    + '.jpg';
}

function saveToDrive(imageBlob, filename, publishDate) {
  try {
    let mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(CONFIG.MAIN_DRIVE_FOLDER_ID);
    } catch (error) {
      const folders = DriveApp.getFoldersByName('News Images Archive');
      if (folders.hasNext()) {
        mainFolder = folders.next();
      } else {
        throw new Error('Could not find main folder');
      }
    }
    
    const dateStr = Utilities.formatDate(publishDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dailyFolderName = `${dateStr} - News Images`;
    
    let dailyFolder;
    const existingFolders = mainFolder.getFoldersByName(dailyFolderName);
    if (existingFolders.hasNext()) {
      dailyFolder = existingFolders.next();
    } else {
      dailyFolder = mainFolder.createFolder(dailyFolderName);
    }
    
    const file = dailyFolder.createFile(imageBlob.setName(filename));
    return file;
    
  } catch (error) {
    throw new Error(`Drive save failed: ${error.message}`);
  }
}

function getShareableUrl(driveFile) {
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = driveFile.getId();
  return `https://drive.google.com/uc?export=view&id=${fileId}`;
}

function getProcessingStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const driveUrlCol = headers.indexOf('Drive Image URL');
  const thumbnailCol = headers.indexOf('Thumbnail');
  const statusCol = headers.indexOf('Image Status');
  
  if (driveUrlCol === -1) {
    console.log('📊 No Drive Image URL column found');
    return;
  }
  
  let totalImages = 0;
  let processedImages = 0;
  let failedImages = 0;
  let enhancedImages = 0;
  let manualDownloadImages = 0;
  let blockedImages = 0;
  let challengeImages = 0;
  let timeoutImages = 0;
  
  for (let i = 1; i < data.length; i++) {
    const thumbnailUrl = data[i][thumbnailCol];
    const driveUrl = data[i][driveUrlCol];
    const status = data[i][statusCol] || '';
    
    if (thumbnailUrl) {
      totalImages++;
      if (driveUrl && driveUrl.includes('drive.google.com')) {
        processedImages++;
        if (status.includes('Enhanced')) enhancedImages++;
      } else if (status.includes('Manual') || (driveUrl && driveUrl.includes('USE FILENAME'))) {
        manualDownloadImages++;
        if (status.includes('Blocked')) blockedImages++;
        if (status.includes('JS Challenge')) challengeImages++;
        if (status.includes('Timeout')) timeoutImages++;
      } else if (status.includes('❌')) {
        failedImages++;
      }
    }
  }
  
  console.log('📊 Enhanced Processing Statistics:');
  console.log(`📸 Total images: ${totalImages}`);
  console.log(`✅ Successfully processed: ${processedImages}`);
  console.log(`🔧 Enhanced method: ${enhancedImages}`);
  console.log(`📝 Marked for manual download: ${manualDownloadImages}`);
  console.log(`  - 🚫 Site blocked: ${blockedImages}`);
  console.log(`  - 🔒 JavaScript challenges: ${challengeImages}`);
  console.log(`  - ⏰ Timeouts: ${timeoutImages}`);
  console.log(`❌ Technical failures: ${failedImages}`);
  console.log(`⏳ Remaining to process: ${totalImages - processedImages - manualDownloadImages - failedImages}`);
  
  if (manualDownloadImages > 0) {
    console.log('\n📋 To find URLs for manual download, look for status containing "Manual"');
  }
}

// Helper function to list URLs that need manual download
function listManualDownloadUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const thumbnailCol = headers.indexOf('Thumbnail');
  const headlineCol = headers.indexOf('Headline');
  const statusCol = headers.indexOf('Image Status');
  const driveUrlCol = headers.indexOf('Drive Image URL');
  
  console.log('📋 URLs requiring manual download:');
  console.log('=====================================');
  
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol] || '';
    
    if (status.includes('Manual')) {
      count++;
      const thumbnailUrl = data[i][thumbnailCol];
      const headline = data[i][headlineCol];
      const driveUrlCell = data[i][driveUrlCol] || '';
      const reason = status.split('→')[0].trim();
      
      // Extract filename from the cell
      let filename = '';
      if (driveUrlCell.includes('USE FILENAME:')) {
        filename = driveUrlCell.replace('📝 USE FILENAME: ', '');
      }
      
      console.log(`${count}. Row ${i + 1}: ${reason}`);
      console.log(`   Title: ${headline}`);
      console.log(`   Download URL: ${thumbnailUrl}`);
      console.log(`   👉 SAVE AS: ${filename}`);
      console.log('');
    }
  }
  
  if (count === 0) {
    console.log('✅ No URLs requiring manual download found!');
  } else {
    console.log(`📝 Total URLs for manual download: ${count}`);
    console.log('\n💡 Manual Download Process:');
    console.log('1. Right-click the "Download URL" → Save image as...');
    console.log('2. Rename the file to the "SAVE AS" filename');
    console.log('3. Upload to your Google Drive folder');
    console.log('4. Replace the cell content with the Google Drive URL');
  }
}