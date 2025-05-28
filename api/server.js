const express = require('express');

const http = require('http');
const WebSocket = require('ws');

const axios = require('axios');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const multer = require('multer');
require('dotenv').config();

const app = express();

const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// Configure multer for file upload
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are allowed'));
    }
  }
});

// Rate limiting configuration
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100 // limit each IP to 100 requests per windowMs
});

// Apply rate limiting to all routes
app.use(limiter);

// ✅ Enable CORS with more specific configuration
app.use(cors({
  origin: ['http://localhost:3000', 'http://localhost:3001'],
  methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization", "Accept"],
  exposedHeaders: ["Content-Range", "X-Content-Range"],
  credentials: true,
  maxAge: 86400 // 24 hours
}));

// Handle preflight requests
app.options('*', cors());

// Add specific CORS headers for the upload endpoint
app.use('/api/upload-to-created-folder', (req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'http://localhost:3001');
  res.header('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');
  res.header('Access-Control-Allow-Credentials', 'true');
  next();
});

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

let latestRecords = new Map(); // ✅ Initialize this

let lastDataMap = new Map(); // ✅ Initialize this

// 🔐 Microsoft credentials from environment variables
const config = {
  tenantId: process.env.TENANT_ID,
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  shareUrl: "u!aHR0cHM6Ly9pbm5vdmFzZW5zZWNvbS1teS5zaGFyZXBvaW50LmNvbS9wZXJzb25hbC9zYW11ZWxfY2hhenlfaW5ub3Zhc2Vuc2VfY29tL19sYXlvdXRzLzE1L29uZWRyaXZlLmFzcHg",
  baseFolderId: "01KT57QGQJ5XUE33KCW5F2XFC3SXMTIZO7"
};
const DRIVE_ID = 'b!ph_tWRNmS0SagtDw028XXogvGAIEG6JEogACR-9Y9Nrlg5qfq5gXQ6FfFDTMThV1';
const PARENT_FOLDER_ID = '01KT57QGTKNYB4RVSPIRDKMSOFUDOKC3S7';
// Validate configuration
function validateConfig() {
  const requiredFields = ['tenantId', 'clientId', 'clientSecret', 'shareUrl', 'baseFolderId'];
  const missingFields = requiredFields.filter(field => !config[field]);
  
  if (missingFields.length > 0) {
    throw new Error(`Missing required environment variables: ${missingFields.join(', ')}`);
  }
}

// Validate config on startup
validateConfig();

// 🔑 Get access token with error handling
async function getAccessToken() {
  try {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    
    const response = await axios.post(
      tokenUrl,
      new URLSearchParams({
        grant_type: "client_credentials",
        client_id: config.clientId,
        client_secret: config.clientSecret,
        scope: "https://graph.microsoft.com/.default",
      }),
      { 
        headers: { 
          'Content-Type': 'application/x-www-form-urlencoded' 
        },
        timeout: 10000 // 10 second timeout
      }
    );
    
    if (!response.data.access_token) {
      throw new Error('No access token received');
    }
    
    return response.data.access_token;
  } catch (error) {
    console.error('Token acquisition error:', error.response?.data || error.message);
    throw new Error(`Failed to get access token: ${error.response?.data?.error_description || error.message}`);
  }
}

// Validate folder name
function validateFolderName(folderName) {
  // SharePoint has restrictions on folder names
  const invalidChars = /[<>:"/\\|?*\x00-\x1F]/g;
  if (invalidChars.test(folderName)) {
    throw new Error('Folder name contains invalid characters');
  }
  
  // Check for reserved names
  const reservedNames = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'LPT1', 'LPT2', 'LPT3'];
  if (reservedNames.includes(folderName.toUpperCase())) {
    throw new Error('Folder name is a reserved name');
  }
  
  return true;
}


let clients = [];


// wss.on('connection', async(ws) => {

//   try {
//     const response = await axios.get('https://api.airtable.com/v0/app6R5bTSGcKo2gmV/tblon8HRet4lsDOUe', {
//       headers: {
//         Authorization: `Bearer pat3fMqN9X4eRWFmd.b31cffaf020d8e4666de0f657adc110e17127c9c38b093cf69d0996fe8e8dfcc`,
//         // Cookie: "brw=brwiMeamMoDgk2PG7; brwConsent=opt-in; AWSALBTG=6eNqOtl90/Hx3Rtqf+hKPZAjhL4G1y7el1V6H0DcIrCDoXON1Zgo+sTBoxAk0nTKQqrWsPbBoicfljEl2ufTJHsVWqgCBxTqjwKAmiMbCOSD5BOq43jRytIwD91E4WvCZiM7ZalTLVCuUoB/MbbbqKHiOyQe1n5PuJtiF0kho7sTp17qCio=; AWSALBTGCORS=6eNqOtl90/Hx3Rtqf+hKPZAjhL4G1y7el1V6H0DcIrCDoXON1Zgo+sTBoxAk0nTKQqrWsPbBoicfljEl2ufTJHsVWqgCBxTqjwKAmiMbCOSD5BOq43jRytIwD91E4WvCZiM7ZalTLVCuUoB/MbbbqKHiOyQe1n5PuJtiF0kho7sTp17qCio=",
//         'Content-Type': 'application/json',
//       },
//     });

//     const allRecords = response.data.records;
//     ws.send(JSON.stringify({ type: 'INIT', records: allRecords }));

//   } catch (error) {
//     console.error('❌ Error sending initial Airtable data:', error);
//   }

// });

// async function fetchAirtableData() {
//   try {
//     const response = await axios.get(
//       `https://api.airtable.com/v0/app6R5bTSGcKo2gmV/tblon8HRet4lsDOUe`,
//       {
//         headers: {
//           Authorization: `Bearer pat3fMqN9X4eRWFmd.b31cffaf020d8e4666de0f657adc110e17127c9c38b093cf69d0996fe8e8dfcc`,
//           // Cookie: "brw=brwiMeamMoDgk2PG7; brwConsent=opt-in; AWSALBTG=6eNqOtl90/Hx3Rtqf+hKPZAjhL4G1y7el1V6H0DcIrCDoXON1Zgo+sTBoxAk0nTKQqrWsPbBoicfljEl2ufTJHsVWqgCBxTqjwKAmiMbCOSD5BOq43jRytIwD91E4WvCZiM7ZalTLVCuUoB/MbbbqKHiOyQe1n5PuJtiF0kho7sTp17qCio=; AWSALBTGCORS=6eNqOtl90/Hx3Rtqf+hKPZAjhL4G1y7el1V6H0DcIrCDoXON1Zgo+sTBoxAk0nTKQqrWsPbBoicfljEl2ufTJHsVWqgCBxTqjwKAmiMbCOSD5BOq43jRytIwD91E4WvCZiM7ZalTLVCuUoB/MbbbqKHiOyQe1n5PuJtiF0kho7sTp17qCio=",
//           'Content-Type': 'application/json',
//         },
//       }
//     );

//     const records = response.data.records;
//     latestRecords = records; // Update stored records for new clients

//     const changedRecords = [];

//     records.forEach((record) => {
//       const prev = lastDataMap.get(record.id);
//       if (!prev || JSON.stringify(prev.fields) !== JSON.stringify(record.fields)) {
//         changedRecords.push(record);
//         lastDataMap.set(record.id, record);
//       }
//     });

//     if (changedRecords.length > 0) {
//       const payload = {
//         type: 'update',
//         changed: changedRecords,
//         all: records,
//       };

//       broadcast(payload);
//     }
//   } catch (error) {
//     console.error('❌ Error fetching Airtable data:', error.message);
//   }
// }

// // Broadcast to all connected clients
// function broadcast(payload) {
//   const message = JSON.stringify(payload);
//   wss.clients.forEach((client) => {
//     if (client.readyState === WebSocket.OPEN) {
//       client.send(message);
//     }
//   });
// }


// setInterval(fetchAirtableData, 50000);

//new approach 

const airtableTables = [
  {
    id: 'tblcdsocb7LuTWZ1B',
    url: 'https://api.airtable.com/v0/app6R5bTSGcKo2gmV/tblcdsocb7LuTWZ1B',
    name: 'Phone_call',
  },
  {
    id: 'tblon8HRet4lsDOUe',
    url: 'https://api.airtable.com/v0/app6R5bTSGcKo2gmV/tblon8HRet4lsDOUe',
    name: 'Screening',
  }
];
// let clients = [];
const airtableHeaders = {
  Authorization: `Bearer pat3fMqN9X4eRWFmd.b31cffaf020d8e4666de0f657adc110e17127c9c38b093cf69d0996fe8e8dfcc`,
  'Content-Type': 'application/json',
};

// const lastDataMap = new Map();

function broadcast(payload) {
  const msg = JSON.stringify(payload);
  wss.clients.forEach((client) => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(msg);
    }
  });
}

async function fetchTableData(table) {
  try {
    const response = await axios.get(table.url, { headers: airtableHeaders });
    const records = response.data.records;

    const changed = [];

    for (const record of records) {
      const key = `${table.id}:${record.id}`;
      const prev = lastDataMap.get(key);

      if (!prev || JSON.stringify(prev.fields) !== JSON.stringify(record.fields)) {
        changed.push(record);
        lastDataMap.set(key, record);
      }
    }

    return { changed, all: records };
  } catch (err) {
    console.error(`❌ Failed to poll ${table.name}:`, err.response?.status || err.message);
    return { changed: [], all: [] };
  }
}

async function pollAllTables() {
  for (const table of airtableTables) {
    const { changed, all } = await fetchTableData(table);
    if (changed.length > 0) {
      broadcast({
        type: 'update',
        tableName: table.name,
        tableId: table.id,
        changed,
        all,
      });
    }
  }
}

// WebSocket connection
wss.on('connection', async (ws) => {
  console.log('🔌 Client connected');

  try {
    for (const table of airtableTables) {
      const response = await axios.get(table.url, { headers: airtableHeaders });
      const records = response.data.records;

      records.forEach((record) => {
        const key = `${table.id}:${record.id}`;
        lastDataMap.set(key, record);
      });

      ws.send(
        JSON.stringify({
          type: 'INIT',
          tableName: table.name,
          tableId: table.id,
          records,
        })
      );
    }
  } catch (err) {
    console.error('❌ Error on init:', err.message);
  }
});

setInterval(pollAllTables, 5000); // Poll every 5 seconds


app.post('/api/create-folder', async (req, res) => {
  console.log('Create folder request received:', req.body);
  
  const { folderName } = req.body;

  try {
    // Validation
    if (!folderName || typeof folderName !== 'string') {
      return res.status(400).json({ 
        error: "Missing or invalid folderName",
        details: "folderName must be a non-empty string" 
      });
    }

    if (folderName.trim().length === 0) {
      return res.status(400).json({ 
        error: "Folder name cannot be empty" 
      });
    }

    // Validate folder name format
    try {
      validateFolderName(folderName);
    } catch (error) {
      return res.status(400).json({
        error: "Invalid folder name",
        details: error.message
      });
    }

    // Sanitize folder name
    const sanitizedFolderName = folderName.trim().substring(0, 255);

    // Get access token
    const token = await getAccessToken();
    console.log('Access token acquired successfully');

    // Create folder in the specified location
    // const endpoint = `https://graph.microsoft.com/v1.0/shares/${config.shareUrl}/driveItem/children`;
    const endpoint ='https://graph.microsoft.com/v1.0/drives/b!ph_tWRNmS0SagtDw028XXogvGAIEG6JEogACR-9Y9Nrlg5qfq5gXQ6FfFDTMThV1/items/01KT57QGQJ5XUE33KCW5F2XFC3SXMTIZO7/children'
    
    console.log('Creating folder at endpoint:', endpoint);

    const createResponse = await axios.post(
      endpoint,
      {
        name: sanitizedFolderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename"
      },
      {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        timeout: 15000 // 15 second timeout
      }
    );

    console.log('Folder created successfully:', createResponse.data.name);
    
    // Return simplified response
    res.status(200).json({
      success: true,
      folder: {
        id: createResponse.data.id,
        name: createResponse.data.name,
        webUrl: createResponse.data.webUrl,
        createdDateTime: createResponse.data.createdDateTime
      }
    });

  } catch (error) {
    console.error('Create folder error:', {
      message: error.message,
      response: error.response?.data,
      status: error.response?.status
    });

    // Handle different types of errors
    if (error.response) {
      // Microsoft Graph API error
      const graphError = error.response.data;
      const statusCode = error.response.status;
      
      // Map common SharePoint errors to appropriate status codes
      const errorMapping = {
        'invalidRequest': 400,
        'itemNotFound': 404,
        'accessDenied': 403,
        'quotaExceeded': 429
      };

      const mappedStatus = errorMapping[graphError.error?.code] || statusCode;
      
      return res.status(mappedStatus).json({
        error: "Microsoft Graph API error",
        details: graphError.error?.message || graphError.error_description || "Unknown API error",
        code: graphError.error?.code
      });
    } else if (error.code === 'ECONNABORTED') {
      // Timeout error
      return res.status(408).json({
        error: "Request timeout",
        details: "The request took too long to complete"
      });
    } else {
      // Other errors (network, etc.)
      return res.status(500).json({
        error: "Internal server error",
        details: error.message
      });
    }
  }
});


// 📄 Upload PDF to folder endpoint
app.post('/api/upload-to-created-folder', upload.single('pdfFile'), async (req, res) => {
  // Set CORS headers for this specific route
  res.header('Access-Control-Allow-Origin', 'http://localhost:3001');
  res.header('Access-Control-Allow-Credentials', 'true');

  const pdfFile = req.file;
  const { folderId } = req.body;
  
  if (!pdfFile || !folderId) {
    return res.status(400).json({ error: 'Missing folder ID or file' });
  }
  
  try {
    const accessToken = await getAccessToken();

    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(pdfFile.originalname)}:/content`;

    const uploadResponse = await axios.put(uploadUrl, pdfFile.buffer, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/pdf',
        'Content-Length': pdfFile.size
      }
    });

    res.status(200).json({
      success: true,
      uploadedFile: {
        id: uploadResponse.data.id,
        name: uploadResponse.data.name,
        webUrl: uploadResponse.data.webUrl,
        size: uploadResponse.data.size
      }
    });

  } catch (error) {
    console.error('Upload error:', error.response?.data || error.message);
    res.status(error.response?.status || 500).json({
      error: 'Upload failed',
      details: error.response?.data || error.message
    });
  }
});


//get all files
app.get('/api/list-files', async (req, res) => {
    const { folderId } = req.query;
  
    if (!folderId) {
      return res.status(400).json({ error: 'Missing folderId in query params' });
    }
  
    try {
      const accessToken = await getAccessToken();
  
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${folderId}/children`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
  
      const files = response.data.value;
      res.json({ files });
    } catch (err) {
      console.error('Graph API Error:', err?.response?.data || err.message);
      res.status(500).json({ error: 'Failed to list files' });
    }
  });


// Health check endpoint
app.get('/api/health', (req, res) => {
  res.status(200).json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    service: 'SharePoint Folder API'
  });
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    error: 'Internal server error',
    details: process.env.NODE_ENV === 'development' ? err.message : 'Something went wrong'
  });
});

// 404 handler
app.use('*', (req, res) => {
  res.status(404).json({
    error: 'Not found',
    details: `Route ${req.method} ${req.originalUrl} not found`
  });
});

const PORT = process.env.PORT || 5001;

server.listen(5174, () => {
  console.log('WebSocket server running on ws://localhost:5172');
});

app.listen(PORT, () => {
  console.log(`✅ Backend running on http://localhost:${PORT}`);
  console.log(`📁 Health check: http://localhost:${PORT}/api/health`);
  console.log(`🔗 Create folder: POST http://localhost:${PORT}/api/create-folder`);
  console.log(`📄 Upload PDF: POST http://localhost:${PORT}/api/upload-to-created-folder`);
});

module.exports = app;