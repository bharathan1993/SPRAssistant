// ========================================
// CONFIGURATION
// ========================================

const GRAFANA_URL = 'https://telemetry-metrics.eks22.uw2.prod.auw2.zuora.com';
const GRAFANA_API_KEY = PropertiesService.getScriptProperties().getProperty('GRAFANA_API_KEY');

// Define all your dashboards here
const DASHBOARDS = {
  'api-health': {
    uid: 'LUq13bv4z',
    path: 'api-health-overall-view',
    name: 'API Health Overall View'
  },
  'api-breakdown': {
    uid: 'oHuq00DVz',
    path: 'api-health-breakdown-view',
    name: 'API Health Breakdown View'
  } 
};

// SPECIFY WHICH DASHBOARD TO USE HERE
const CURRENT_DASHBOARD = 'api-health';

// ========================================
// ADD-ON LIFECYCLE FUNCTIONS
// ========================================

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SlidesApp.getUi()
      .createAddonMenu()
      .addItem('Generate Report', 'showSidebar')
      .addSeparator()
      .addItem('List Available Panels', 'listAllPanels')
      .addItem('Debug: Show Slide Titles', 'debugShowSlideTitles')
      .addSeparator()
      .addItem('Refresh Panel Cache', 'refreshPanelCache')
      .addItem('Clear All Caches', 'clearAllCaches')
      .addToUi();
}

function showSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('SPR Assistant');
  SlidesApp.getUi().showSidebar(ui);
}

// ========================================
// PANEL FETCHING WITH CACHING
// ========================================

/**
 * Fetch all panels from dashboard including those in collapsed rows
 */
function fetchDashboardPanels(dashboardKey) {
  const dashboard = DASHBOARDS[dashboardKey];
  
  if (!dashboard) {
    throw new Error('Dashboard not found: ' + dashboardKey);
  }
  
  const url = `${GRAFANA_URL}/api/dashboards/uid/${dashboard.uid}`;
  
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + GRAFANA_API_KEY,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };
  
  try {
    Logger.log('Fetching dashboard: ' + dashboard.name);
    Logger.log('UID: ' + dashboard.uid);
    Logger.log('URL: ' + url);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error('Grafana API returned error code: ' + responseCode);
    }
    
    const data = JSON.parse(response.getContentText());
    const allPanels = [];
    
    function extractAllPanels(panels, depth = 0) {
      const indent = '  '.repeat(depth);
      
      panels.forEach(panel => {
        if (panel.type === 'row') {
          Logger.log(indent + 'Row: "' + panel.title + '"');
          
          if (panel.panels && Array.isArray(panel.panels)) {
            Logger.log(indent + '  → Contains ' + panel.panels.length + ' nested panels');
            extractAllPanels(panel.panels, depth + 1);
          }
        } else {
          allPanels.push({
            id: panel.id,
            title: panel.title,
            type: panel.type,
            dashboardUid: dashboard.uid,
            dashboardPath: dashboard.path
          });
          Logger.log(indent + 'Panel: ID=' + panel.id + ' "' + panel.title + '"');
        }
      });
    }
    
    Logger.log('');
    Logger.log('Extracting panels...');
    extractAllPanels(data.dashboard.panels);
    
    Logger.log('');
    Logger.log('✓ Total panels extracted: ' + allPanels.length);
    
    return allPanels;
    
  } catch (e) {
    Logger.log('✗ Error fetching dashboard: ' + e.message);
    throw new Error('Failed to fetch dashboard panels: ' + e.message);
  }
}

/**
 * Cache panel mappings in script properties
 */
function cachePanelMappings(dashboardKey) {
  Logger.log('Caching panels for: ' + dashboardKey);
  
  const panels = fetchDashboardPanels(dashboardKey);
  const mapping = {};
  
  panels.forEach(panel => {
    mapping[panel.title] = panel;
    const trimmed = panel.title.trim();
    if (trimmed !== panel.title) {
      mapping[trimmed] = panel;
    }
  });
  
  const cache = PropertiesService.getScriptProperties();
  const cacheKey = 'PANEL_CACHE_' + dashboardKey;
  cache.setProperty(cacheKey, JSON.stringify(mapping));
  
  Logger.log('✓ Cached ' + Object.keys(mapping).length + ' panel entries for ' + dashboardKey);
  
  return mapping;
}

/**
 * Get panels from cache or fetch if not cached
 */
function getPanelsWithCache(dashboardKey) {
  const cache = PropertiesService.getScriptProperties();
  const cacheKey = 'PANEL_CACHE_' + dashboardKey;
  const cached = cache.getProperty(cacheKey);
  
  if (cached) {
    Logger.log('✓ Using cached panels for ' + dashboardKey);
    return JSON.parse(cached);
  }
  
  Logger.log('Cache miss - fetching and caching panels for ' + dashboardKey);
  return cachePanelMappings(dashboardKey);
}

/**
 * Refresh panel cache for current dashboard
 */
function refreshPanelCache() {
  try {
    Logger.log('Refreshing panel cache for: ' + CURRENT_DASHBOARD);
    
    cachePanelMappings(CURRENT_DASHBOARD);
    
    const ui = SlidesApp.getUi();
    ui.alert(
      'Cache Refreshed',
      'Panel cache refreshed for dashboard: ' + DASHBOARDS[CURRENT_DASHBOARD].name,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log('Error: ' + e.message);
    SlidesApp.getUi().alert('Error', e.message, SlidesApp.getUi().ButtonSet.OK);
  }
}

/**
 * Clear all panel caches
 */
function clearAllCaches() {
  try {
    const cache = PropertiesService.getScriptProperties();
    const keys = cache.getKeys();
    
    let cleared = 0;
    keys.forEach(key => {
      if (key.startsWith('PANEL_CACHE_')) {
        cache.deleteProperty(key);
        cleared++;
      }
    });
    
    Logger.log('✓ Cleared ' + cleared + ' panel caches');
    
    const ui = SlidesApp.getUi();
    ui.alert(
      'Caches Cleared',
      'Cleared ' + cleared + ' panel cache(s)',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log('Error: ' + e.message);
    SlidesApp.getUi().alert('Error', e.message, SlidesApp.getUi().ButtonSet.OK);
  }
}

// ========================================
// MAIN REPORT GENERATION
// ========================================

/**
 * Generate report using cached panels
 */
function generateReportFromSidebar(inputs) {
  try {
    Logger.log('');
    Logger.log('========== STARTING REPORT GENERATION ==========');
    Logger.log('Dashboard: ' + CURRENT_DASHBOARD + ' (' + DASHBOARDS[CURRENT_DASHBOARD].name + ')');
    Logger.log('Inputs:');
    Logger.log('  Tenant ID: ' + inputs.tenantId);
    Logger.log('  Data Source: ' + inputs.data_source);
    Logger.log('  Environment: ' + inputs.environment);
    Logger.log('  Interval: ' + inputs.interval);
    Logger.log('  Timeframe: ' + inputs.timeframe);
    
    if (!inputs.tenantId || inputs.tenantId.trim() === '') {
      throw new Error('Tenant ID is required');
    }
    
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    
    Logger.log('Presentation has ' + slides.length + ' slides');
    
    if (slides.length === 0) {
      throw new Error('No slides found in presentation');
    }
    
    const panelMap = getPanelsWithCache(CURRENT_DASHBOARD);
    
    Logger.log('');
    Logger.log('Available panels (' + Object.keys(panelMap).length + ' total):');
    const uniquePanels = {};
    Object.values(panelMap).forEach(panel => {
      if (!uniquePanels[panel.id]) {
        uniquePanels[panel.id] = panel;
        Logger.log('  ID ' + panel.id + ': "' + panel.title + '"');
      }
    });
    
    let processedCount = 0;
    let skippedCount = 0;
    const errors = [];
    
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const slideTitle = extractSlideTitle(slide);
      
      Logger.log('');
      Logger.log('========================================');
      Logger.log('Slide ' + (i + 1) + ' of ' + slides.length);
      Logger.log('========================================');
      
      if (!slideTitle) {
        Logger.log('⊗ No title found, skipping');
        skippedCount++;
        continue;
      }
      
      Logger.log('Slide title: "' + slideTitle + '"');
      
      const panel = panelMap[slideTitle] || panelMap[slideTitle.trim()];
      
      if (panel) {
        Logger.log('✓ Found Panel ID: ' + panel.id);
        
        try {
          const chartConfig = {
            dashboardUid: panel.dashboardUid,
            panelId: panel.id,
            chartTitle: slideTitle,
            dashboardPath: panel.dashboardPath
          };
          
          const imageUrl = buildGrafanaRenderUrl(chartConfig, inputs);
          populateSlideViaGoogleDrive(slide, slideTitle, imageUrl);
          
          processedCount++;
          Logger.log('✓✓✓ SLIDE COMPLETED ✓✓✓');
          
        } catch (slideError) {
          Logger.log('✗ Error: ' + slideError.message);
          errors.push('Slide "' + slideTitle + '": ' + slideError.message);
          skippedCount++;
        }
        
      } else {
        Logger.log('⊗ No matching panel found');
        skippedCount++;
      }
    }
    
    Logger.log('');
    Logger.log('========== COMPLETE ==========');
    Logger.log('✓ Processed: ' + processedCount);
    Logger.log('⊗ Skipped: ' + skippedCount);
    Logger.log('==============================');
    
    let message = '';
    if (processedCount > 0) {
      message = 'Report generated successfully!\n\n' +
                '✓ Updated ' + processedCount + ' chart(s)\n';
      if (skippedCount > 0) {
        message += '⊗ Skipped ' + skippedCount + ' slide(s)';
      }
    } else {
      message = 'No charts generated!\n\n' +
                '⊗ Skipped ' + skippedCount + ' slide(s)\n\n' +
                'Make sure slide titles match panel names exactly.';
    }
    
    if (errors.length > 0) {
      message += '\n\nErrors:\n' + errors.slice(0, 3).join('\n');
      if (errors.length > 3) {
        message += '\n... and ' + (errors.length - 3) + ' more';
      }
    }
    
    return message;
    
  } catch (e) {
    Logger.log('✗ FATAL ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    throw new Error('Failed to generate report: ' + e.message);
  }
}

// ========================================
// URL BUILDING AND IMAGE INSERTION
// ========================================

/**
 * Build Grafana render URL with user inputs
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
  
  const params = {
    'orgId': '1',
    'panelId': chartConfig.panelId,
    'from': timeFrameMap[inputs.timeframe] || 'now-24h',
    'to': 'now',
    'var-data_source': dataSourceMap[inputs.data_source] || 'Pinot Telemetry (US-Prod)',
    'var-environment': environmentMap[inputs.environment] || 'prod02',
    'var-tenant_id': inputs.tenantId || '',
    'var-entity_id': '11e64eef-ad7b-6780-9658-00259058c29c',
    'var-API': 'All',
    'var-ZuoraResponseCode': 'All',
    'var-HttpStatus': 'All',
    'var-Client_twosinglequote': 'All',
    'var-Client_query_string': '',
    'var-GFW_Bucket': 'All',
    'var-interval': intervalMap[inputs.interval] || 'day',
    'var-gfw_time_range_from': '1.761966092427e+12',
    'var-restapi_entity_id_mapping_table': 'restapi_entity_id_mapping',
    'width': '1000',
    'height': '500',
    'tz': 'UTC'
  };
  
  const urlParams = Object.keys(params)
    .filter(key => params[key] !== '')
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  return `${baseUrl}?${urlParams}`;
}

/**
 * Populate slide with chart image via Google Drive
 */
function populateSlideViaGoogleDrive(slide, title, imageUrl) {
  try {
    Logger.log('  → Processing chart: ' + title);
    
    const shapes = slide.getShapes();
    let placeholder = null;
    let left = 50;
    let top = 150;
    let width = 600;
    let height = 400;
    
    Logger.log('  → Scanning ' + shapes.length + ' shapes for placeholder');
    
    // Method 1: Find by alt text "ImagePlaceholder"
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      try {
        const shapeTitle = shape.getTitle();
        if (shapeTitle === 'ImagePlaceholder') {
          placeholder = shape;
          Logger.log('  → Found placeholder by alt text');
          break;
        }
      } catch (e) {}
    }
    
    // Method 2: Find largest non-text rectangle
    if (!placeholder) {
      Logger.log('  → Searching for largest rectangle...');
      let maxArea = 0;
      
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        try {
          const shapeType = shape.getShapeType();
          
          if (shapeType === SlidesApp.ShapeType.RECTANGLE) {
            let hasText = false;
            try {
              const text = shape.getText().asString().trim();
              hasText = text.length > 0;
            } catch (e) {}
            
            if (!hasText) {
              const w = shape.getWidth();
              const h = shape.getHeight();
              const area = w * h;
              
              if (area > maxArea) {
                maxArea = area;
                placeholder = shape;
              }
            }
          }
        } catch (e) {}
      }
      
      if (placeholder) {
        Logger.log('  → Found placeholder as largest rectangle');
      }
    }
    
    if (placeholder) {
      left = placeholder.getLeft();
      top = placeholder.getTop();
      width = placeholder.getWidth();
      height = placeholder.getHeight();
      Logger.log('  → Placeholder: (' + left.toFixed(0) + ', ' + top.toFixed(0) + ') ' + width.toFixed(0) + 'x' + height.toFixed(0));
      placeholder.remove();
    } else {
      Logger.log('  → No placeholder found, using defaults');
    }
    
    Logger.log('  → Fetching image from Grafana...');
    
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
    
    const response = UrlFetchApp.fetch(imageUrl, options);
    const responseCode = response.getResponseCode();
    
    Logger.log('  → Response: ' + responseCode);
    
    if (responseCode !== 200) {
      throw new Error('Grafana returned error code ' + responseCode);
    }
    
    let blob = response.getBlob();
    Logger.log('  → Downloaded: ' + (blob.getBytes().length / 1024).toFixed(1) + ' KB');
    
    const fileName = 'grafana_temp_' + Date.now() + '.png';
    blob = blob.setName(fileName);
    
    Logger.log('  → Saving to Drive...');
    const tempFile = DriveApp.createFile(blob);
    
    Logger.log('  → Inserting into slide...');
    const insertedImage = slide.insertImage(tempFile);
    insertedImage.setLeft(left);
    insertedImage.setTop(top);
    insertedImage.setWidth(width);
    insertedImage.setHeight(height);
    
    Logger.log('  ✓ Chart inserted');
    
    Utilities.sleep(2000);
    tempFile.setTrashed(true);
    Logger.log('  ✓ Cleanup complete');
    
  } catch (e) {
    Logger.log('  ✗ Error: ' + e.message);
    throw new Error('Failed to populate slide "' + title + '": ' + e.message);
  }
}

// ========================================
// UTILITY FUNCTIONS
// ========================================

/**
 * List all available panels
 */
function listAllPanels() {
  try {
    const panelMap = getPanelsWithCache(CURRENT_DASHBOARD);
    
    const uniquePanels = {};
    Object.values(panelMap).forEach(panel => {
      uniquePanels[panel.id] = panel;
    });
    
    const panels = Object.values(uniquePanels);
    
    Logger.log('========== AVAILABLE PANELS ==========');
    panels.forEach(panel => {
      Logger.log('ID: ' + panel.id + ' | Title: "' + panel.title + '"');
    });
    Logger.log('======================================');
    
    const ui = SlidesApp.getUi();
    const panelList = panels
      .sort((a, b) => a.id - b.id)
      .map(p => 'ID ' + p.id + ': ' + p.title)
      .join('\n');
    
    ui.alert(
      'Available Panels (' + panels.length + ' total)',
      'Dashboard: ' + DASHBOARDS[CURRENT_DASHBOARD].name + '\n\n' + panelList,
      ui.ButtonSet.OK
    );
    
    return panels;
    
  } catch (e) {
    Logger.log('Error: ' + e.message);
    SlidesApp.getUi().alert('Error', e.message, SlidesApp.getUi().ButtonSet.OK);
    throw e;
  }
}

/**
 * Debug: Show all slide titles
 */
function debugShowSlideTitles() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    
    const titlesFound = [];
    const noTitleSlides = [];
    
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const title = extractSlideTitle(slide);
      
      if (title) {
        titlesFound.push((i + 1) + '. "' + title + '"');
      } else {
        noTitleSlides.push(i + 1);
      }
    }
    
    const ui = SlidesApp.getUi();
    let message = 'Found ' + titlesFound.length + ' slide(s) with titles:\n\n' +
                  titlesFound.join('\n');
    
    if (noTitleSlides.length > 0) {
      message += '\n\nSlides without titles: ' + noTitleSlides.join(', ');
    }
    
    ui.alert('Slide Titles', message, ui.ButtonSet.OK);
    
  } catch (e) {
    SlidesApp.getUi().alert('Error', e.message, SlidesApp.getUi().ButtonSet.OK);
  }
}

/**
 * Extract the title from a slide
 */
function extractSlideTitle(slide) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    
    try {
      const textRange = shape.getText();
      const text = textRange.asString().trim();
      
      if (text.length > 0 && text.length < 200) {
        return text;
      }
    } catch (e) {}
  }
  
  return null;
}