const express = require('express');
const path = require('path');
const fs = require('fs');
const useragent = require('express-useragent');
const axios = require('axios');
const multer = require('multer');
const docxConverter = require('docx-pdf');
require('dotenv').config();

// 1. MODULAR FIREBASE IMPORT WORKAROUND
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore, FieldValue } = require('firebase-admin/firestore');
const { getStorage } = require('firebase-admin/storage');

const app = express();
const PORT = process.env.PORT || 3000;

// ----------------------------------------------------
// FIREBASE INITIALIZATION (Live Safe Mode)
// ----------------------------------------------------
let serviceAccount;

if (process.env.FIREBASE_SERVICE_ACCOUNT_JSON) {
    // Live Deployment: Parse the raw JSON string stored in your environment variable
    try {
        serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON);
    } catch (e) {
        console.error("FATAL ERROR: Failed to parse FIREBASE_SERVICE_ACCOUNT_JSON environment variable.", e.message);
        process.exit(1);
    }
} else {
    // Local Development: Fall back to reading the local JSON file
    try {
        serviceAccount = require('./firebase-service-account.json');
    } catch (e) {
        console.warn("WARNING: No local firebase-service-account.json found. Ensure environment variables are configured.");
    }
}

// Ensure we have a valid configuration before trying to initialize Firebase
if (!serviceAccount) {
    console.error("FATAL ERROR: Firebase credentials are missing. Please define FIREBASE_SERVICE_ACCOUNT_JSON in your environment or provide a local firebase-service-account.json file.");
    process.exit(1);
}

// Initialize the App
const firebaseApp = initializeApp({
    credential: cert(serviceAccount),
    storageBucket: process.env.FIREBASE_STORAGE_BUCKET || "courses-pq.firebasestorage.app"
});

// Grab database and storage instances
const db = getFirestore(firebaseApp);
const bucket = getStorage(firebaseApp).bucket();

// Access Pins (Pulls securely from your environment variables)
const STUDENT_PIN = process.env.STUDENT_PIN || "";
const ADMIN_PIN = process.env.ADMIN_PIN || "";

// Middleware
app.use(express.json());
app.use(useragent.express());
app.use(express.static(path.join(__dirname, 'public')));

// Temp local directory for converting files
const TEMP_UPLOADS_DIR = path.join(__dirname, 'temp_uploads');
if (!fs.existsSync(TEMP_UPLOADS_DIR)) {
    fs.mkdirSync(TEMP_UPLOADS_DIR, { recursive: true });
}

// Multer Storage configurations
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, TEMP_UPLOADS_DIR);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + path.extname(file.originalname));
    }
});

const fileFilter = (req, file, cb) => {
    const filetypes = /doc|docx|application\/msword|application\/vnd.openxmlformats-officedocument.wordprocessingml.document/;
    const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = filetypes.test(file.mimetype);

    if (extname && mimetype) {
        return cb(null, true);
    } else {
        cb(new Error('Only Microsoft Word (.doc, .docx) documents are allowed!'));
    }
};

const upload = multer({ storage, fileFilter });

// New Promise wrapper using docx-pdf (No LibreOffice required)
const convertDocToPdf = (inputPath, outputPath) => {
    return new Promise((resolve, reject) => {
        docxConverter(inputPath, outputPath, (err, result) => {
            if (err) {
                return reject(err);
            }
            resolve(outputPath);
        });
    });
};

// ----------------------------------------------------
// ENDPOINTS
// ----------------------------------------------------

// PIN Authorization Endpoints
app.post('/api/auth/student', (req, res) => {
    const { pin } = req.body;
    if (pin === STUDENT_PIN) {
        return res.json({ success: true, message: "Access Granted" });
    }
    return res.status(401).json({ success: false, message: "Invalid Student Access PIN" });
});

app.post('/api/auth/admin', (req, res) => {
    const { pin } = req.body;
    if (pin === ADMIN_PIN) {
        return res.json({ success: true, message: "Access Granted" });
    }
    return res.status(401).json({ success: false, message: "Invalid Admin Access PIN" });
});

// Logs Endpoint
app.post('/api/logs', async (req, res) => {
    const { email } = req.body;
    
    // 1. Safely extract the client IP address
    let ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || '';
    
    if (ip.includes(',')) {
        ip = ip.split(',')[0].trim();
    }

    // Clean IPv6 tunnel prefix if present (e.g. ::ffff:127.0.0.1)
    if (ip.startsWith('::ffff:')) {
        ip = ip.substring(7);
    }

    const deviceType = req.useragent.isMobile ? 'Mobile' : req.useragent.isTablet ? 'Tablet' : 'Desktop';
    const os = req.useragent.os;
    const browser = req.useragent.browser;

    // Fallback location object structure
    let locationData = {
        text: "Unknown Location",
        lat: null,
        lon: null
    };

    // Detect if development localhost IP environment is running
    const isLocal = !ip || 
                    ip === '::1' || 
                    ip === '127.0.0.1' || 
                    ip.startsWith('10.') || 
                    ip.startsWith('192.168.') || 
                    ip.startsWith('172.16.');

    try {
        let lookupIp = ip;

        // LOCALHOST DEVELOPMENT WORKAROUND:
        if (isLocal) {
            try {
                const publicIpResponse = await axios.get('https://api.ipify.org?format=json', { timeout: 3000 });
                if (publicIpResponse.data && publicIpResponse.data.ip) {
                    lookupIp = publicIpResponse.data.ip;
                }
            } catch (ipifyErr) {
                console.warn("Could not retrieve external public IP via ipify, using fallback location.");
            }
        }

        // Perform Geolocation lookup on the resolved IP
        if (lookupIp && lookupIp !== '127.0.0.1' && lookupIp !== '::1') {
            const geoRes = await axios.get(`http://ip-api.com/json/${lookupIp}`);
            if (geoRes.data && geoRes.data.status === 'success') {
                const city = geoRes.data.city || "Unknown City";
                const region = geoRes.data.regionName || "";
                const country = geoRes.data.country || "Unknown Country";
                const lat = geoRes.data.lat || null;
                const lon = geoRes.data.lon || null;
                
                locationData.text = region ? `${city}, ${region}, ${country}` : `${city}, ${country}`;
                locationData.lat = lat;
                locationData.lon = lon;
            } else {
                locationData.text = `IP (${lookupIp}) Failed to Resolve`;
            }
        } else {
            locationData.text = "Localhost Sandbox Loopback";
        }
    } catch (e) {
        console.error("Geo-location lookup failed:", e.message);
        locationData.text = "Location Lookup Error";
    }

    const logEntry = {
        timestamp: FieldValue.serverTimestamp(),
        displayTime: new Date().toLocaleString(),
        email: email || "Anonymous (PIN Only)",
        ip: ip || "Unknown IP",
        location: locationData, // Saved as an object with text, lat, and lon
        device: `${deviceType} (${os} - ${browser})`
    };

    try {
        await db.collection('logs').add(logEntry);
        res.json({ success: true });
    } catch (err) {
        console.error("Firestore logging error:", err);
        res.status(500).json({ error: "Failed to write session log" });
    }
});

// Admin fetching logs (returns the Firestore document ID)
app.get('/api/admin/logs', async (req, res) => {
    try {
        const snapshot = await db.collection('logs').orderBy('timestamp', 'desc').limit(100).get();
        const logs = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            
            // Backwards-compatible location parser 
            let locationObj = { text: "Unknown", lat: null, lon: null };
            if (data.location) {
                if (typeof data.location === 'object') {
                    locationObj = data.location;
                } else {
                    locationObj.text = data.location;
                }
            }

            logs.push({
                id: doc.id, // Capture and pass the Firestore Document ID
                timestamp: data.displayTime || "N/A",
                email: data.email,
                location: locationObj,
                device: data.device
            });
        });
        res.json(logs);
    } catch (err) {
        console.error("Firestore fetching logs error:", err);
        res.status(500).json({ error: "Failed to read session logs" });
    }
});

// DELETE log endpoint
app.delete('/api/admin/logs/:id', async (req, res) => {
    const logId = req.params.id;
    try {
        const docRef = db.collection('logs').doc(logId);
        const doc = await docRef.get();

        if (!doc.exists) {
            return res.status(404).json({ success: false, error: "Log entry not found" });
        }

        // Delete Firestore Log Document
        await docRef.delete();

        res.json({ success: true, message: "Log entry deleted successfully." });
    } catch (err) {
        console.error("Error deleting log:", err);
        res.status(500).json({ error: "Internal Server Error occurred during log deletion." });
    }
});

// Get Courses Catalog
app.get('/api/courses', async (req, res) => {
    try {
        const snapshot = await db.collection('courses').orderBy('createdAt', 'desc').get();
        const courses = [];
        snapshot.forEach(doc => {
            courses.push({ id: doc.id, ...doc.data() });
        });
        res.json(courses);
    } catch (err) {
        console.error("Firestore loading courses error:", err);
        res.status(500).json({ error: "Failed to retrieve courses" });
    }
});

// Upload, Convert, and Store Route
app.post('/api/courses', upload.single('wordFile'), async (req, res) => {
    const { code, title, questions } = req.body;
    
    if (!code || !title) {
        return res.status(400).json({ error: "Missing required fields" });
    }

    if (!req.file) {
        return res.status(400).json({ error: "Please attach a Word Solution Document (.docx / .doc)" });
    }

    const localWordPath = req.file.path;
    const pdfFilename = `${path.basename(req.file.filename, path.extname(req.file.filename))}.pdf`;
    const localPdfPath = path.join(TEMP_UPLOADS_DIR, pdfFilename);

    try {
        // 1. Convert Word to PDF
        await convertDocToPdf(localWordPath, localPdfPath);

        // 2. Upload output PDF to Cloud Storage
        const destination = `courses/${pdfFilename}`;
        const [uploadedFile] = await bucket.upload(localPdfPath, {
            destination: destination,
            metadata: {
                contentType: 'application/pdf',
            }
        });

        // 3. Make public
        await uploadedFile.makePublic();
        const publicUrl = `https://storage.googleapis.com/${bucket.name}/${destination}`;

        // 4. Cleanup local temp files
        fs.unlinkSync(localWordPath);
        fs.unlinkSync(localPdfPath);

        let parsedQuestions = [];
        if (questions) {
            parsedQuestions = typeof questions === 'string' ? JSON.parse(questions) : questions;
        }

        // 5. Store metadata structure in Firestore
        const newCourse = {
            code,
            title,
            questions: parsedQuestions,
            pdfUrl: publicUrl,
            storagePath: destination, // Retain storage reference so we can easily delete it later
            originalFileName: req.file.originalname.replace(/\.[^/.]+$/, "") + ".pdf",
            createdAt: FieldValue.serverTimestamp()
        };

        const docRef = await db.collection('courses').add(newCourse);
        
        res.json({ success: true, courseId: docRef.id });
    } catch (err) {
        console.error("Failure processing document:", err);
        
        // Cleanup local files if they exist on crash
        if (fs.existsSync(localWordPath)) fs.unlinkSync(localWordPath);
        if (fs.existsSync(localPdfPath)) fs.unlinkSync(localPdfPath);

        res.status(500).json({ error: "Conversion or Database error occurred." });
    }
});

// DELETE Endpoint: Removes course from Firestore database and cleans up Firebase Storage file
app.delete('/api/courses/:id', async (req, res) => {
    const courseId = req.params.id;
    try {
        const docRef = db.collection('courses').doc(courseId);
        const doc = await docRef.get();

        if (!doc.exists) {
            return res.status(404).json({ success: false, error: "Course record not found" });
        }

        const courseData = doc.data();

        // 1. Delete associated physical PDF from storage bucket
        if (courseData.storagePath) {
            const file = bucket.file(courseData.storagePath);
            await file.delete().catch(err => {
                console.warn("File was already removed or inaccessible in Storage bucket:", err.message);
            });
        }

        // 2. Delete Firestore Document
        await docRef.delete();

        res.json({ success: true, message: "Course document and cloud file deleted successfully." });
    } catch (err) {
        console.error("Error deleting course:", err);
        res.status(500).json({ error: "Internal Server Error occurred during deletion process." });
    }
});

app.listen(PORT, () => console.log(`🚀 Portal active on http://localhost:${PORT}`));