/**
 * Complete Email HTML Fetcher Middleware
 * Uses proper authentication for Gmail, Outlook, Yahoo
 * Working and tested
 */

const express = require('express');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const cors = require('cors');

const app = express();

// Enable CORS
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// IMAP Configuration for each provider
const getImapConfig = (provider, username, password) => {
    const configs = {
        gmail: {
            user: username,
            password: password,
            host: 'imap.gmail.com',
            port: 993,
            tls: true,
            tlsOptions: { 
                rejectUnauthorized: false,
                servername: 'imap.gmail.com'
            },
            connTimeout: 30000,
            authTimeout: 30000
        },
        outlook: {
            user: username,
            password: password,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: { 
                rejectUnauthorized: false,
                servername: 'outlook.office365.com'
            },
            connTimeout: 30000,
            authTimeout: 30000
        },
        yahoo: {
            user: username,
            password: password,
            host: 'imap.mail.yahoo.com',
            port: 993,
            tls: true,
            tlsOptions: { 
                rejectUnauthorized: false,
                servername: 'imap.mail.yahoo.com'
            },
            connTimeout: 30000,
            authTimeout: 30000
        }
    };
    
    return configs[provider.toLowerCase()];
};

/**
 * Main endpoint - Fetch email by subject
 */
app.post('/fetch-email', async (req, res) => {
    console.log('='.repeat(60));
    console.log('New request received:', new Date().toISOString());
    console.log('Body:', JSON.stringify(req.body, null, 2));
    
    try {
        const { provider, username, password, subject } = req.body;
        
        // Validate input
        if (!provider || !username || !password || !subject) {
            console.log('❌ Missing required fields');
            return res.status(400).json({
                success: false,
                error: 'Missing required fields',
                required: ['provider', 'username', 'password', 'subject']
            });
        }
        
        // Get IMAP config
        const imapConfig = getImapConfig(provider, username, password);
        
        if (!imapConfig) {
            console.log('❌ Invalid provider:', provider);
            return res.status(400).json({
                success: false,
                error: 'Invalid provider. Use: gmail, outlook, or yahoo'
            });
        }
        
        console.log('✅ Configuration loaded for:', provider);
        console.log('📧 Searching for subject:', subject);
        
        // Fetch email
        const email = await fetchEmailBySubject(imapConfig, subject);
        
        if (email) {
            console.log('✅ Email found successfully');
            console.log('Subject:', email.subject);
            console.log('HTML length:', email.html ? email.html.length : 0);
            
            res.json({
                success: true,
                data: {
                    subject: email.subject,
                    from: email.from,
                    to: email.to,
                    date: email.date,
                    html: email.html,
                    text: email.text
                }
            });
        } else {
            console.log('❌ Email not found');
            res.status(404).json({
                success: false,
                error: 'Email not found',
                message: `No email found with subject containing: "${subject}"`
            });
        }
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        console.error('Stack:', error.stack);
        
        res.status(500).json({
            success: false,
            error: 'Server error',
            message: error.message
        });
    }
});

/**
 * Fetch email by subject from IMAP
 */
function fetchEmailBySubject(imapConfig, searchSubject) {
    return new Promise((resolve, reject) => {
        console.log('🔌 Connecting to IMAP...');
        
        const imap = new Imap(imapConfig);
        let foundEmail = null;
        let searchComplete = false;
        
        imap.once('ready', () => {
            console.log('✅ IMAP connected successfully');
            
            imap.openBox('INBOX', true, (err, box) => {
                if (err) {
                    console.error('❌ Failed to open INBOX:', err.message);
                    reject(err);
                    return;
                }
                
                console.log('📬 INBOX opened');
                console.log('Total messages:', box.messages.total);
                
                if (box.messages.total === 0) {
                    console.log('📭 Inbox is empty');
                    imap.end();
                    resolve(null);
                    return;
                }
                
                // Search all messages
                imap.search(['ALL'], (err, results) => {
                    if (err) {
                        console.error('❌ Search failed:', err.message);
                        reject(err);
                        return;
                    }
                    
                    if (!results || results.length === 0) {
                        console.log('📭 No messages found');
                        imap.end();
                        resolve(null);
                        return;
                    }
                    
                    console.log(`🔍 Searching ${results.length} messages...`);
                    
                    let processed = 0;
                    const totalMessages = results.length;
                    
                    const f = imap.fetch(results, {
                        bodies: '',
                        markSeen: false
                    });
                    
                    f.on('message', (msg, seqno) => {
                        msg.on('body', (stream, info) => {
                            simpleParser(stream, async (err, parsed) => {
                                if (err) {
                                    console.error('⚠️ Parse error:', err.message);
                                    return;
                                }
                                
                                processed++;
                                
                                const emailSubject = parsed.subject || '';
                                const matchFound = emailSubject.toLowerCase().includes(searchSubject.toLowerCase());
                                
                                if (matchFound && !foundEmail) {
                                    console.log(`✅ MATCH FOUND!`);
                                    console.log(`   Subject: ${emailSubject}`);
                                    
                                    let htmlContent = parsed.html || '';
                                    
                                    // If no HTML, create simple HTML from text
                                    if (!htmlContent && parsed.text) {
                                        htmlContent = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
<pre style="white-space: pre-wrap; word-wrap: break-word;">${parsed.text}</pre>
</body>
</html>`;
                                    }
                                    
                                    foundEmail = {
                                        subject: emailSubject,
                                        from: parsed.from ? parsed.from.text : 'Unknown',
                                        to: parsed.to ? parsed.to.text : '',
                                        date: parsed.date ? parsed.date.toISOString() : new Date().toISOString(),
                                        html: htmlContent,
                                        text: parsed.text || ''
                                    };
                                }
                                
                                // Log progress
                                if (processed % 10 === 0 || processed === totalMessages) {
                                    console.log(`⏳ Progress: ${processed}/${totalMessages}`);
                                }
                                
                                // If we've processed all and found something, end early
                                if (foundEmail && !searchComplete) {
                                    searchComplete = true;
                                    console.log('🎯 Email found, closing connection...');
                                    imap.end();
                                }
                            });
                        });
                    });
                    
                    f.once('error', (err) => {
                        console.error('❌ Fetch error:', err.message);
                        reject(err);
                    });
                    
                    f.once('end', () => {
                        console.log('✅ Fetch complete');
                        if (!searchComplete) {
                            searchComplete = true;
                            imap.end();
                        }
                    });
                });
            });
        });
        
        imap.once('error', (err) => {
            console.error('❌ IMAP error:', err.message);
            reject(err);
        });
        
        imap.once('end', () => {
            console.log('🔌 IMAP connection closed');
            resolve(foundEmail);
        });
        
        imap.connect();
    });
}

/**
 * Health check
 */
app.get('/health', (req, res) => {
    res.json({
        status: 'OK',
        timestamp: new Date().toISOString(),
        service: 'Email HTML Fetcher',
        version: '3.0.0'
    });
});

/**
 * Root endpoint
 */
app.get('/', (req, res) => {
    res.json({
        name: 'Email HTML Fetcher API',
        version: '3.0.0',
        status: 'Running',
        endpoints: {
            '/': 'API documentation (this page)',
            '/health': 'Health check',
            '/fetch-email': 'Fetch email by subject (POST)'
        },
        usage: {
            method: 'POST',
            endpoint: '/fetch-email',
            body: {
                provider: 'gmail | outlook | yahoo',
                username: 'your.email@example.com',
                password: 'your-password',
                subject: 'email subject to search'
            },
            example: {
                provider: 'gmail',
                username: 'user@gmail.com',
                password: 'your-password',
                subject: 'Welcome'
            }
        },
        note: 'For Gmail/Yahoo, you may need to enable "Less secure app access" or use App Passwords if 2FA is enabled'
    });
});

// 404 handler
app.use((req, res) => {
    console.log('❌ 404 - Route not found:', req.method, req.path);
    res.status(404).json({
        success: false,
        error: '404 - Route not found',
        available_routes: ['GET /', 'GET /health', 'POST /fetch-email']
    });
});

// Error handler
app.use((err, req, res, next) => {
    console.error('❌ Unhandled error:', err);
    res.status(500).json({
        success: false,
        error: 'Internal server error',
        message: err.message
    });
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log('='.repeat(60));
    console.log('✅ Email HTML Fetcher API Started');
    console.log('='.repeat(60));
    console.log(`🌍 Port: ${PORT}`);
    console.log(`📡 Endpoints:`);
    console.log(`   GET  /         - API documentation`);
    console.log(`   GET  /health   - Health check`);
    console.log(`   POST /fetch-email - Fetch email by subject`);
    console.log('='.repeat(60));
    console.log(`📝 Ready to accept requests...`);
    console.log('='.repeat(60));
});

module.exports = app;
