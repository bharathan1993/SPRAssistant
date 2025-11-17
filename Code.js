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
    chartTitle: 'Overall API Volume Trend',
    dashboardPath: 'api-health-overall-view'
  },
   { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 33, 
    chartTitle: 'Overall API Errors Trending',
    dashboardPath: 'api-health-overall-view'
  },
  { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 35, 
    chartTitle: 'API Errors Ranking by Zuora Response Code',
    dashboardPath: 'api-health-overall-view'
  },
  { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 60, 
    chartTitle: 'API Concurrency Usage',
    dashboardPath: 'api-health-overall-view'
  },
    { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 23, 
    chartTitle: 'Overall API Performance Trending',
    dashboardPath: 'api-health-overall-view'
  },
    { 
    dashboardUid: 'LUq13bv4z', 
    panelId: 67, 
    chartTitle: 'API Calls Count by Latency Range',
    dashboardPath: 'api-health-overall-view'
  }
  // Add more charts as needed
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
 * FIXED: Uses insertSlide with explicit index for precise ordering
 */
function generateReportFromSidebar(inputs) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    
    // Get all slides
    const slides = presentation.getSlides();
    
    // Clear ALL existing slides except the template (first one)
    for (let i = slides.length - 1; i >= 1; i--) {
      slides[i].remove();
    }

    const templateSlide = slides[0];
    
    // Process each chart
    CHARTS_TO_PULL.forEach((chartConfig, index) => {
      Logger.log('');
      Logger.log('=== Processing Chart ' + (index + 1) + ' ===');
      Logger.log('Panel ID: ' + chartConfig.panelId);
      Logger.log('Title: ' + chartConfig.chartTitle);
      
      // Insert a new slide at the END (after all existing slides)
      // This ensures order is preserved
      const currentSlide = presentation.appendSlide(templateSlide);
      Logger.log('Appended new slide at end of presentation');
      
      // Build the image URL with user inputs
      const imageUrl = buildGrafanaRenderUrl(chartConfig, inputs);
      
      // Populate this slide
      populateSlideViaGoogleDrive(currentSlide, chartConfig.chartTitle, imageUrl);
    });
    
    // Remove the template at the very end
    Logger.log('');
    Logger.log('Removing template slide...');
    templateSlide.remove();

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
    'Last 7 Days': 'now-7d',
    'Last 30 Days': 'now-30d'
  };
  
  const baseUrl = `${GRAFANA_URL}/render/d-solo/${chartConfig.dashboardUid}/${chartConfig.dashboardPath}`;
  
  // Build parameters using ALL user inputs from sidebar
  const params = {
    'orgId': '1',
    'panelId': chartConfig.panelId,
    'from': timeFrameMap[inputs.timeframe] || 'now-24h',
    'to': 'now',
    'var-data_source': dataSourceMap[inputs.data_source] || 'Pinot Telemetry (US-Prod)',  // FROM SIDEBAR
    'var-environment': environmentMap[inputs.environment] || 'prod02',  // FROM SIDEBAR
    'var-tenant_id': inputs.tenantId || '',  // FROM SIDEBAR
    'var-entity_id': '11e64eef-ad7b-6780-9658-00259058c29c',  // Hardcoded (may need to be dynamic)
    'var-API': 'All',  // Hardcoded
    'var-ZuoraResponseCode': 'All',  // Hardcoded
    'var-HttpStatus': 'All',  // Hardcoded
    'var-Client_twosinglequote': 'All',  // Hardcoded
    'var-Client_query_string': '',  // Hardcoded
    'var-GFW_Bucket': 'All',  // Hardcoded
    'var-interval': intervalMap[inputs.interval] || 'day',  // FROM SIDEBAR
    'var-gfw_time_range_from': '1.761966092427e+12',  // Hardcoded
    'var-restapi_entity_id_mapping_table': 'restapi_entity_id_mapping',  // Hardcoded
    'width': '1000',
    'height': '500',
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
    Logger.log('========================================');
    Logger.log('Starting populateSlideViaGoogleDrive for: ' + title);
    Logger.log('URL: ' + imageUrl);
    
    // Replace title placeholder
    slide.replaceAllText('{{CHART_TITLE}}', title);
    Logger.log('Replaced chart title with: ' + title);
    
    // Find placeholder - TRY MULTIPLE METHODS
    const shapes = slide.getShapes();
    let placeholder = null;
    let left = 50;   // Default position
    let top = 100;
    let width = 600;  // Default size
    let height = 400;
    
    Logger.log('Total shapes found: ' + shapes.length);
    
    // Method 1: Try to find by Alt Text Title
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      try {
        const shapeTitle = shape.getTitle();
        if (shapeTitle === 'ImagePlaceholder') {
          placeholder = shape;
          Logger.log('✅ Found placeholder by title at index ' + i);
          break;
        }
      } catch (e) {
        // Some shapes may not support getTitle()
      }
    }
    
    // Method 2: If not found by title, find the largest rectangle
    if (!placeholder) {
      Logger.log('Placeholder not found by title, searching for largest rectangle...');
      
      let maxArea = 0;
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        try {
          if (shape.getShapeType() === SlidesApp.ShapeType.RECTANGLE) {
            const w = shape.getWidth();
            const h = shape.getHeight();
            const area = w * h;
            
            Logger.log('Rectangle ' + i + ': ' + w + 'x' + h + ' = ' + area);
            
            if (area > maxArea) {
              maxArea = area;
              placeholder = shape;
            }
          }
        } catch (e) {
          Logger.log('Could not get dimensions for shape ' + i);
        }
      }
      
      if (placeholder) {
        Logger.log('✅ Found placeholder as largest rectangle (area: ' + maxArea + ')');
      }
    }

    // If placeholder found, get its dimensions and remove it
    if (placeholder) {
      try {
        left = placeholder.getLeft();
        top = placeholder.getTop();
        width = placeholder.getWidth();
        height = placeholder.getHeight();
        
        Logger.log('Placeholder position: (' + left + ', ' + top + ')');
        Logger.log('Placeholder size: ' + width + ' x ' + height);
        
        placeholder.remove();
        Logger.log('Placeholder removed');
      } catch (e) {
        Logger.log('⚠️ Could not remove placeholder: ' + e.message);
        Logger.log('Will insert at default position instead');
      }
    } else {
      Logger.log('⚠️ No placeholder found - will insert at default position');
      Logger.log('Default position: (' + left + ', ' + top + ')');
      Logger.log('Default size: ' + width + ' x ' + height);
    }

    // STEP 1: Fetch image from Grafana
    Logger.log('');
    Logger.log('Fetching from Grafana...');
    
    const options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'validateHttpsCertificates': false
    };
    
    if (GRAFANA_API_KEY) {
      options.headers = {
        'Authorization': 'Bearer ' + GRAFANA_API_KEY
      };
      Logger.log('Using API key authentication');
    }
    
    let response;
    try {
      const startTime = new Date().getTime();
      response = UrlFetchApp.fetch(imageUrl, options);
      const duration = new Date().getTime() - startTime;
      Logger.log('Fetch completed in ' + (duration/1000).toFixed(1) + ' seconds');
    } catch (fetchError) {
      Logger.log('❌ Fetch error: ' + fetchError.message);
      throw new Error('Failed to fetch from Grafana: ' + fetchError.message);
    }
    
    const responseCode = response.getResponseCode();
    Logger.log('Response code: ' + responseCode);
    
    if (responseCode === 504) {
      throw new Error('Grafana timeout (504). Try: 1) Shorter time range, 2) Contact Grafana admin');
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
      const htmlSnippet = blob.getDataAsString().substring(0, 200);
      Logger.log('HTML content: ' + htmlSnippet);
      throw new Error('Grafana returned HTML instead of image. Check panel ID and parameters.');
    }
    
    // STEP 2: Save to Google Drive temporarily
    Logger.log('');
    Logger.log('Saving to Google Drive...');
    
    const fileName = 'grafana_temp_' + Date.now() + '.png';
    blob = blob.setName(fileName);
    
    let tempFile;
    try {
      tempFile = DriveApp.createFile(blob);
      Logger.log('Created Drive file: ' + tempFile.getId());
    } catch (driveError) {
      Logger.log('❌ Drive error: ' + driveError.message);
      throw new Error('Failed to create Drive file: ' + driveError.message);
    }
    
    // STEP 3: Insert image from Drive file
    Logger.log('');
    Logger.log('Inserting image from Drive...');
    
    try {
      const insertedImage = slide.insertImage(tempFile);
      
      // Position and size the image
      insertedImage.setLeft(left);
      insertedImage.setTop(top);
      insertedImage.setWidth(width);
      insertedImage.setHeight(height);
      
      Logger.log('✅ Image inserted and positioned');
    } catch (insertError) {
      Logger.log('❌ Insert error: ' + insertError.message);
      // Clean up before throwing
      try {
        tempFile.setTrashed(true);
      } catch (e) {}
      throw new Error('Failed to insert image: ' + insertError.message);
    }
    
    // STEP 4: Wait a moment for image to be processed, then delete temp file
    Logger.log('');
    Logger.log('Cleaning up...');
    Utilities.sleep(2000);
    
    try {
      tempFile.setTrashed(true);
      Logger.log('Temp file deleted');
    } catch (deleteError) {
      Logger.log('⚠️ Warning: Could not delete temp file: ' + deleteError.message);
    }
    
    Logger.log('');
    Logger.log('✅✅ Successfully populated slide for: ' + title);
    Logger.log('========================================');
    
  } catch (e) {
    Logger.log('');
    Logger.log('❌❌ Error in populateSlideViaGoogleDrive: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    Logger.log('========================================');
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
