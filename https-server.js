const https = require('https');
const fs = require('fs');
const path = require('path');

// Read PFX certificate
const pfxPath = path.join(__dirname, '.certs', 'localhost.pfx');
const pfxData = fs.readFileSync(pfxPath);

const options = {
  pfx: pfxData,
  passphrase: 'password'
};

const handler = require('serve-handler');

const server = https.createServer(options, (request, response) => {
  return handler(request, response, {
    public: __dirname,
    cleanUrls: false,
    directoryListing: false
  });
});

server.listen(3000, () => {
  console.log('HTTPS Server running at https://localhost:3000');
  console.log('Press Ctrl+C to stop');
});
