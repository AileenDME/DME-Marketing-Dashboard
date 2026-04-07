import axios from 'axios';

// This serverless function handles OAuth and fetches Excel data from SharePoint
// Deploy this to your Vercel project

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  try {
    // Get access token from Azure AD
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      }
    );

    const accessToken = tokenResponse.data.access_token;

    // SharePoint file ID and site info
    const siteId = process.env.SHAREPOINT_SITE_ID;
    const driveId = process.env.SHAREPOINT_DRIVE_ID;
    const itemId = process.env.SHAREPOINT_ITEM_ID;

    // Fetch workbook sessions and sheet data
    const headers = { Authorization: `Bearer ${accessToken}` };

    // Get all worksheets
    const sheetsResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/workbook/worksheets`,
      { headers }
    );

    const sheets = sheetsResponse.data.value;
    const sheetMap = {};

    // Fetch data from each sheet
    for (const sheet of sheets) {
      const sheetName = sheet.name;
      
      try {
        const rangeResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`,
          { headers }
        );

        sheetMap[sheetName] = rangeResponse.data.values;
      } catch (error) {
        console.error(`Error fetching sheet ${sheetName}:`, error.message);
        sheetMap[sheetName] = [];
      }
    }

    // Parse and format the data
    const formattedData = parseExcelData(sheetMap);

    res.status(200).json(formattedData);
  } catch (error) {
    console.error('Error fetching Excel data:', error);
    res.status(500).json({ 
      error: 'Failed to fetch Excel data',
      message: error.message 
    });
  }
}

function parseExcelData(sheetMap) {
  const parseSheet = (data, valueIndex = 1) => {
    if (!data || data.length < 2) return [];
    
    return data.slice(1).map(row => ({
      date: formatDate(row[0]),
      value: parseFloat(row[valueIndex]) || 0
    })).filter(item => item.date);
  };

  const formatDate = (excelDate) => {
    if (!excelDate) return null;
    
    if (typeof excelDate === 'string') {
      return excelDate;
    }
    
    // Excel date serial number
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    return date.toISOString().split('T')[0];
  };

  const parseProjects = (data) => {
    if (!data || data.length < 2) return [];
    
    return data.slice(1).map(row => ({
      name: row[0],
      percentage: (parseFloat(row[1]) || 0),
      status: row[2] || 'On Track'
    })).filter(p => p.name);
  };

  return {
    leadsGenerated: parseSheet(sheetMap['Leads Generated'], 1),
    leadsVelocity: parseSheet(sheetMap['Leads Velocity'], 1),
    domainAuthority: parseSheet(sheetMap['Domain Authority'], 1),
    usTraffic: parseSheet(sheetMap['US Traffic'], 1),
    linkedinFollowers: parseSheet(sheetMap['LinkedIn Followers'], 1),
    salesDeck: parseSheet(sheetMap['Sales Deck Created'], 1),
    rfpResponses: parseSheet(sheetMap['RFP Responses'], 1),
    projects: parseProjects(sheetMap['Projects']),
    sharePointLink: 'https://dmeservicesolutions-my.sharepoint.com/:x:/g/personal/aileen_manalo_dmeserve_com/IQBIgn_igqUHQa0xmaavGaToAeuVpToC4cHb8IrqkN-JryQ'
  };
}
