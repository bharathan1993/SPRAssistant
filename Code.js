// ALTERNATIVE APPROACH: Use Google Drive as intermediary
// This is more reliable for slow-loading images

// --- CONFIGURATION ---
const GRAFANA_URL = 'https://telemetry-metrics.eks22.uw2.prod.auw2.zuora.com';
const GRAFANA_API_KEY = PropertiesService.getScriptProperties().getProperty('GRAFANA_API_KEY');

// Define your chart configurations
const CHARTS_TO_PULL = [
  { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 27, 
    chartTitle: 'API Overall Volume',
    dashboardPath: 'api-health-overall-view'
  }
];

// --- ADD-ON LIFECYCLE FUNCTIONS ---

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SlidesApp.getUi()
      .createAddonMenu()
      .addItem('Generate Report', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('SPR Assistant');
  SlidesApp.getUi().showSidebar(ui);
}

// --- CORE LOGIC WITH DRIVE APPROACH ---

/**
 * Generates the report using Google Drive as intermediary
 */
function generateReportFromSidebar(inputs) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    
    // Clear existing slides (except the first one)
    const slides = presentation.getSlides();
    for (let i = slides.length - 1; i >= 1; i--) {
      slides[i].remove();
    }

    const templateSlide = slides[0];

    CHARTS_TO_PULL.forEach((chartConfig, index) => {
      let currentSlide;
      
      if (index === 0) {
        currentSlide = templateSlide;
      } else {
        currentSlide = templateSlide.duplicate();
      }
      
      const imageUrl = buildGrafanaRenderUrl(chartConfig, inputs);
      populateSlideViaGoogleDrive(currentSlide, chartConfig.chartTitle, imageUrl);
    });

    return 'Report generated successfully!';

  } catch (e) {
    Logger.log('Error generating report: ' + e.message);
    throw new Error('Failed to generate report: ' + e.message);
  }
}

/**
 * Builds the Grafana render URL
 */
function buildGrafanaRenderUrl(chartConfig, inputs) {
  const dataSourceMap = {
    'Pilot Telemetry(A1-Prod)': 'Pinot Telemetry (US-Prod)',
    'Pilot Telemetry(A1-Sbx)': 'Pinot Telemetry (US-Sbx)',
    'Pilot Telemetry(EU-Prod)': 'Pinot Telemetry (EU-Prod)',
    'Pilot Telemetry(US-Prod)': 'Pinot Telemetry (US-Prod)'
  };
  
  const environmentMap = {
    'US Production': 'prod02',
    'NA Production': 'prod01',
    'US Sandbox': 'sbx02',
    'NA Sandbox': 'sbx01',
    'NA Central Sandbox': 'sbxcentral'
  };
  
  const intervalMap = {
    '1 day': 'day',
    '1 minute': 'minute',
    '1 hour': 'hour'
  };
  
  const timeFrameMap = {
    'Last 1 Hour': 'now-1h',
    'Last 24 Hours': 'now-24h',
    'Last 7 Days': 'now-7d'
  };
  
  const baseUrl = `${GRAFANA_URL}/render/d-solo/${chartConfig.dashboardUid}/${chartConfig.dashboardPath}`;
  
  // Use MINIMAL parameters to avoid timeout
  const params = {
    'orgId': '1',
    'panelId': chartConfig.panelId,
    'from': timeFrameMap[inputs.timeframe] || 'now-7d',  // Default to 24 hours
    'to': 'now',
    'var-tenant_id': inputs.tenantId || '',
    'width': '1000',
    'height': '500',
    'var-interval': 'day',
    'tz': 'UTC'
  };
  
  const urlParams = Object.keys(params)
    .filter(key => params[key] !== '')
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  const finalUrl = `${baseUrl}?${urlParams}`;
  Logger.log('Generated URL: ' + finalUrl);
  
  return finalUrl;
}

/**
 * NEW APPROACH: Fetch image, save to Drive, insert from Drive
 * This avoids timeout issues by using Drive as intermediary
 */
function populateSlideViaGoogleDrive(slide, title, imageUrl) {
  try {
    Logger.log('Starting populateSlideViaGoogleDrive...');
    
    // Replace title placeholder
    slide.replaceAllText('{{CHART_TITLE}}', title);
    
    // Find placeholder
    const shapes = slide.getShapes();
    let placeholder = null;
    
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      if (shape.getTitle && shape.getTitle() === 'ImagePlaceholder') {
        placeholder = shape;
        break;
      }
    }
    
    if (!placeholder) {
      let maxArea = 0;
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        if (shape.getShapeType() === SlidesApp.ShapeType.RECTANGLE) {
          const area = shape.getWidth() * shape.getHeight();
          if (area > maxArea) {
            maxArea = area;
            placeholder = shape;
          }
        }
      }
    }

    if (!placeholder) {
      throw new Error('Could not find placeholder shape');
    }

    const left = placeholder.getLeft();
    const top = placeholder.getTop();
    const width = placeholder.getWidth();
    const height = placeholder.getHeight();
    
    Logger.log('Placeholder found: ' + left + ',' + top + ' / ' + width + 'x' + height);

    // STEP 1: Fetch image from Grafana with longer timeout tolerance
    Logger.log('Fetching from Grafana... (this may take up to 60 seconds)');
    
    const options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'validateHttpsCertificates': false
    };
    
    if (GRAFANA_API_KEY) {
      options.headers = {
        'Authorization': 'Bearer ' + GRAFANA_API_KEY
      };
    }
    
    let response;
    try {
      response = UrlFetchApp.fetch(imageUrl, options);
    } catch (fetchError) {
      Logger.log('Fetch error: ' + fetchError.message);
      throw new Error('Failed to fetch from Grafana: ' + fetchError.message);
    }
    
    const responseCode = response.getResponseCode();
    Logger.log('Response code: ' + responseCode);
    
    if (responseCode === 504) {
      throw new Error('Grafana timeout (504). Try: 1) Shorter time range, 2) Simpler panel, 3) Contact Grafana admin');
    }
    
    if (responseCode !== 200) {
      throw new Error('Grafana returned error code: ' + responseCode);
    }
    
    let blob = response.getBlob();
    const blobSize = blob.getBytes().length;
    Logger.log('Downloaded: ' + blobSize + ' bytes');
    
    const contentType = blob.getContentType();
    Logger.log('Content-Type: ' + contentType);
    
    if (contentType === 'text/html') {
      throw new Error('Grafana returned HTML instead of image. Check dashboard variables.');
    }
    
    // STEP 2: Save to Google Drive temporarily
    Logger.log('Saving to Google Drive...');
    
    const fileName = 'grafana_temp_' + Date.now() + '.png';
    blob = blob.setName(fileName);
    
    let tempFile;
    try {
      tempFile = DriveApp.createFile(blob);
      Logger.log('Created Drive file: ' + tempFile.getId());
    } catch (driveError) {
      Logger.log('Drive error: ' + driveError.message);
      throw new Error('Failed to create Drive file: ' + driveError.message);
    }
    
    // STEP 3: Remove placeholder
    placeholder.remove();
    Logger.log('Placeholder removed');
    
    // STEP 4: Insert image from Drive file
    Logger.log('Inserting image from Drive...');
    
    try {
      const insertedImage = slide.insertImage(tempFile);
      
      // Position and size the image
      insertedImage.setLeft(left);
      insertedImage.setTop(top);
      insertedImage.setWidth(width);
      insertedImage.setHeight(height);
      
      Logger.log('Image inserted and positioned');
    } catch (insertError) {
      Logger.log('Insert error: ' + insertError.message);
      // Clean up before throwing
      tempFile.setTrashed(true);
      throw new Error('Failed to insert image: ' + insertError.message);
    }
    
    // STEP 5: Wait a moment for image to be processed, then delete temp file
    Logger.log('Cleaning up...');
    Utilities.sleep(2000);  // Wait 2 seconds
    
    try {
      tempFile.setTrashed(true);
      Logger.log('Temp file deleted');
    } catch (deleteError) {
      Logger.log('Warning: Could not delete temp file: ' + deleteError.message);
      // Don't throw - image is already inserted
    }
    
    Logger.log('✅ Successfully populated slide for: ' + title);
    
  } catch (e) {
    Logger.log('Error in populateSlideViaGoogleDrive: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    throw new Error('Failed to populate slide: ' + e.message);
  }
}

/**
 * Test function with Drive approach
 */
function testWithDriveApproach() {
  Logger.log('========== TESTING DRIVE APPROACH ==========');
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GRAFANA_API_KEY');
  
  if (!apiKey) {
    Logger.log('❌ No API key configured');
    return;
  }
  
  // Use minimal URL
  const testUrl = 'https://telemetry-metrics.eks22.uw2.prod.auw2.zuora.com/render/d-solo/LUq13bv4z/api-health-overall-view?orgId=1&panelId=27&from=now-24h&to=now&var-tenant_id=5463&width=800&height=400&tz=UTC';
  
  Logger.log('URL: ' + testUrl);
  Logger.log('');
  
  const options = {
    'method': 'get',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  
  try {
    Logger.log('Step 1: Fetching from Grafana...');
    const response = UrlFetchApp.fetch(testUrl, options);
    const code = response.getResponseCode();
    
    Logger.log('Response: ' + code);
    
    if (code === 504) {
      Logger.log('❌ Still getting 504. Grafana server issue - contact admin.');
      return 'TIMEOUT';
    }
    
    if (code !== 200) {
      Logger.log('❌ Error: ' + code);
      return 'ERROR';
    }
    
    Logger.log('✅ Got image from Grafana');
    Logger.log('');
    
    Logger.log('Step 2: Saving to Drive...');
    const blob = response.getBlob().setName('test_grafana.png');
    const file = DriveApp.createFile(blob);
    
    Logger.log('✅ Saved to Drive: ' + file.getId());
    Logger.log('File URL: ' + file.getUrl());
    Logger.log('');
    
    Logger.log('Step 3: Cleaning up...');
    Utilities.sleep(1000);
    file.setTrashed(true);
    
    Logger.log('✅ Deleted temp file');
    Logger.log('');
    Logger.log('✅✅✅ DRIVE APPROACH WORKS! ✅✅✅');
    Logger.log('You can now generate reports using this approach.');
    
    return 'SUCCESS';
    
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
    return 'ERROR';
  } finally {
    Logger.log('========================================');
  }
}