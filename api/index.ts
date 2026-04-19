import express from 'express';
import { S3Client, ListObjectsV2Command, GetObjectCommand } from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import session from 'express-session';
import cookieParser from 'cookie-parser';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

// AWS S3 Config
const s3Client = new S3Client({
  region: process.env.AWS_REGION || 'us-east-1',
  credentials: {
    accessKeyId: process.env.AWS_ACCESS_KEY_ID || 'MISSING',
    secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY || 'MISSING',
  },
});

app.use(cookieParser());
app.use(express.json());
app.use(session({
  secret: process.env.SESSION_SECRET || 'kavach-secret',
  resave: false,
  saveUninitialized: true,
  cookie: { 
    secure: true, 
    sameSite: 'none',
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));

// AWS S3 API Routes
app.get('/api/aws/files', async (req, res) => {
  const { AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_BUCKET_NAME, AWS_REGION } = process.env;
  
  if (!AWS_ACCESS_KEY_ID || !AWS_SECRET_ACCESS_KEY || !AWS_BUCKET_NAME) {
    // If we are in a demo/development environment, we can return mock files
    // to allow the user to see the dashboard in action.
    console.log('AWS Configuration missing. Returning mock files for demo purposes.');
    
    const mockFiles = [
      {
        id: 'demo/20260328_VAPI_RFCOMM_TR.csv',
        name: 'demo/20260328_VAPI_RFCOMM_TR.csv',
        mimeType: 'text/csv',
        size: 1024,
        modifiedTime: new Date().toISOString(),
        source: 'aws',
        isMock: true
      },
      {
        id: 'demo/20260328_VAPI_RFCOMM_ST.csv',
        name: 'demo/20260328_VAPI_RFCOMM_ST.csv',
        mimeType: 'text/csv',
        size: 2048,
        modifiedTime: new Date().toISOString(),
        source: 'aws',
        isMock: true
      },
      {
        id: 'demo/20260328_VAPI_TRNMSNMA.csv',
        name: 'demo/20260328_VAPI_TRNMSNMA.csv',
        mimeType: 'text/csv',
        size: 512,
        modifiedTime: new Date().toISOString(),
        source: 'aws',
        isMock: true
      }
    ];

    return res.json({ 
      files: mockFiles,
      warning: 'Using Mock Data: AWS Configuration is missing. Please add AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, and AWS_BUCKET_NAME in Settings > Environment Variables for real data.'
    });
  }

  console.log(`Attempting to list files in bucket: ${AWS_BUCKET_NAME} (Region: ${AWS_REGION || 'us-east-1'})`);

  try {
    const command = new ListObjectsV2Command({
      Bucket: process.env.AWS_BUCKET_NAME,
    });
    const response = await s3Client.send(command);
    const files = response.Contents?.map(file => ({
      id: file.Key,
      name: file.Key,
      mimeType: 'application/octet-stream',
      size: file.Size,
      modifiedTime: file.LastModified,
      source: 'aws'
    })) || [];
    res.json({ files });
  } catch (error: any) {
    console.error('Error listing S3 files:', error);
    res.status(500).json({ 
      error: 'Failed to list AWS S3 files', 
      details: error.message || String(error) 
    });
  }
});

app.get('/api/aws/download', async (req, res) => {
  const key = req.query.key as string;
  
  if (!key) {
    return res.status(400).json({ error: 'Missing file key' });
  }

  // Handle mock files
  if (key.startsWith('demo/')) {
    return res.json({ url: `/api/mock/data?key=${encodeURIComponent(key)}` });
  }

  try {
    const command = new GetObjectCommand({
      Bucket: process.env.AWS_BUCKET_NAME,
      Key: key,
    });
    const url = await getSignedUrl(s3Client, command, { expiresIn: 3600 });
    res.json({ url });
  } catch (error: any) {
    console.error('Error generating S3 signed URL:', error);
    res.status(500).json({ 
      error: 'Failed to generate download link',
      details: error.message || String(error)
    });
  }
});

// Mock Data Serving Route
app.get('/api/mock/data', (req, res) => {
  const key = req.query.key as string;
  
  res.setHeader('Content-Type', 'text/csv');
  
  // Correcting mock data headers to match user's screenshot exactly
  if (key.includes('RFCOMM_TR')) {
    res.send(`Date,Time,Loco Id,Station Id,Station Name,Direction,Expected,Received,Percentage,Radio
28-03-2026,10:00:00,37887,VAPI,VAPI STATION,Nominal,100,98,98,1
28-03-2026,10:05:00,37887,VAPI,VAPI STATION,Reverse,100,95,95,1
28-03-2026,10:10:00,37424,UVD,UDVADA STATION,Nominal,100,92,92,1
28-03-2026,10:15:00,37424,UVD,UDVADA STATION,Reverse,100,88,88,1`);
  } else if (key.includes('RFCOMM_ST')) {
    res.send(`Date,Time,Loco Id,Station Id,Station Name,Direction,Expected,Received,Percentage,Radio
28-03-2026,10:00:00,37887,VAPI,VAPI STATION,Nominal,100,99,99,1
28-03-2026,10:05:00,37887,VAPI,VAPI STATION,Reverse,100,97,97,1
07-04-2026,10:00:00,37424,BIM,BIM STATION,Nominal,1000,976.8,97.68,1
07-04-2026,10:05:00,37424,GVD,GVD STATION,Nominal,10000,9453,94.53,1`);
  } else if (key.includes('TRNMSNMA')) {
    res.send(`Date,Time,Loco Id,Station Id,Station Name,Radio,Event
28-03-2026,10:00:00,37887,VAPI,VAPI STATION,1,MA Received
28-03-2026,10:05:00,37887,VAPI,VAPI STATION,1,MA Received
07-04-2026,10:00:00,37424,BIM,BIM STATION,1,MA Received`);
  } else {
    res.send('Date,Time,Info\n28-03-2026,10:00:00,Mock Data');
  }
});

export default app;
