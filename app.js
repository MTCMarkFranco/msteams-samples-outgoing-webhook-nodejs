//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
const { CardFactory } = require('botbuilder');
const crypto = require('crypto');
const restify = require('restify');
const path = require('path');
const ENV_FILE = path.join(__dirname, '.env');
const fs = require('fs');
require('dotenv').config({ path: ENV_FILE });

// Setup SSL Config using local certificates
var https_options = {
    certificate: fs.readFileSync('./certs/certificate.crt'),
    ca: fs.readFileSync('./certs/full-chain-cert.p7b'),
    key: fs.readFileSync('./certs/key.key')
};

// Build the node restify server with SSL Bindings...
const server = restify.createServer(https_options);
// Enable bodyParser Plugin for payload capture...
server.use(restify.plugins.bodyParser());

// Setup Ports and start listening...
server.listen(process.env.port || process.env.PORT || 443, function(request, response) {
    console.log(`\n${ server.name } listening to ${ server.url }.`);
	console.log('\nSend Messages from Teams!!!\n');
	console.log('\*************************************\n');
});

// Catch-All for every request (Used for debugging...)
server.on('request', (request, response) => {
	console.log(`INFORMATION: Incoming Request - Route: "${ request.url }" Attatched Headers: "${ JSON.stringify(request.headers,null,4) }"`);
  });

// Listen on the following routes
server.post('/api', function (request, response, next) {

	// Retrieve payload from incoming request
	var payloadJson = JSON.stringify(request.body);

	// Retrieve authorizatin HMAC information provided by TEAMS infrastructure
	var teamsHmac = request.headers['authorization'];

	// Get payload intoa UTF-8 encoded Byte Stream
	var utfMessage = Buffer.from(payloadJson, 'utf8');

	// Get Shared Secret from ENVIRONMENT Variables and store in b64 Buffer
	var b64Secret = Buffer.from(process.env.TeamsSharedSecret, "base64");

	// Calculate HMAC on the message we've received using the shared secret			
	var messageHmac = "HMAC " + crypto.createHmac('sha256', b64Secret).update(utfMessage).digest("base64");

	// Ouput HMACs for debugging purposes
	console.log("Computed HMAC: " + messageHmac);
	console.log("Received HMAC: " + teamsHmac);

	// Validate that you got the message from proper TEAMS channel by analyzing the HMACs
	var responseCard = '';
	if (messageHmac === teamsHmac) {
		responseCard = returnCard(JSON.parse(payloadJson).text);

	} else {
		responseCard = '{ "type": "message", "text": "Error: message sender cannot be authenticated. Make sure your authentication secret is still valid!" }';
	}

	// Write the response and send the Information back to TEAMS6
	response.writeHead(200);
	response.write(responseCard);
	response.end();

});

function returnCard (message) {
	
	const heroCard = CardFactory.heroCard('OutBound WebHook Sample',
            'This sample demonstrates how to handle outbound webhooks in Teams.  Please review the readme for more information.');
		heroCard.content.subtitle = message;
		heroCard.content.images = [{"url":"https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcTnfpHwTiWB3TOPt4jOaqIP9cJCnmcq_Ysjgg&usqp=CAU"}];
		heroCard.content.buttons = [{"type": "openUrl", "title": "Microsoft","value": "https://www.microsoft.com"}];
		const attachment = { ...heroCard, heroCard };
		
		return JSON.stringify({
				type: 'result',
				attachmentLayout: 'list',
				attachments: [
					attachment
				]
		});
};