var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');
var microsoftGraph = require("@microsoft/microsoft-graph-client"); // this is the Outlook API

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/mail'] = mail; //mail route

server.start(router.route, handle);

function home(response, request) {
	
	console.log('Request handler \'home\' was called.');
  	response.writeHead(200, {'Content-Type': 'text/html'});
	response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
 	//response.write('<p>Please <a href='#'>sign in</a> with your Office 365 or Outlook.com account.</p>');
 	response.end();
	
 	/*	
  	console.log('Request handler \'home\' was called.');
  	response.writeHead(200, {'Content-Type': 'text/html'});
  	response.write('<p>Hello Outlook User!</p>');
  	response.end();
	*/
	
	
}

//invoke authorize function
var url = require('url');

	function authorize(response, request) {
  		console.log('Request handler \'authorize\' was called.');

  		// The authorization code is passed as a query parameter
  		var url_parts = url.parse(request.url, true);
  		var code = url_parts.query.code;
  		console.log('Code: ' + code);
		authHelper.getTokenFromCode(code, tokenReceived, response);
		
		/*
  		response.writeHead(200, {'Content-Type': 'text/html'});
  		response.write('<p>Received auth code: ' + code + '</p>');
  		response.end();
		*/
	}
	

	function getUserEmail(token, callback) {
  // Create a Graph client
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  client
    .api('/me')
    .get((err, res) => {
      if (err) {
        callback(err, null);
      } else {
        callback(null, res.mail);
      }
    });
}
	
//Callback function tokenReceived - store in cookie
	function tokenReceived(response, error, token) {
  if (error) {
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  } else {
    getUserEmail(token.token.access_token, function(error, email){
      if (error) {
        console.log('getUserEmail returned an error: ' + error);
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
      } else if (email) {
		
		 //save the refresh token and the expiration time in a session cookie 
		 var cookies = ['OutlookConnectByJules-token=' + token.token.access_token + ';Max-Age=4000',
                       'OutlookConnectByJules-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                       'OutlookConnectByJules-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                       'OutlookConnectByJules-email=' + email + ';Max-Age=4000'];
		  
		/*  
        var cookies = ['OutlookConnectByJules-token=' + token.token.access_token + ';Max-Age=3600',
                       'OutlookConnectByJules-email=' + email + ';Max-Age=3600'];
		*/  
        response.setHeader('Set-Cookie', cookies);
        response.writeHead(302, {'Location': 'http://localhost:8000/mail'});
        response.end();
      }
    }); 
  }
}
	
	
//read cookie values
	function getValueFromCookie(valueName, cookie) {
  if (cookie.indexOf(valueName) !== -1) {
    var start = cookie.indexOf(valueName) + valueName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}
	
	
//retrieve the cached token, check if it is expired, and refresh
	function getAccessToken(request, response, callback) {
  var expiration = new Date(parseFloat(getValueFromCookie('OutlookConnectByJules-token-expires', request.headers.cookie)));

  if (expiration <= new Date()) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    var refresh_token = getValueFromCookie('OutlookConnectByJules-refresh-token', request.headers.cookie);
    authHelper.refreshAccessToken(refresh_token, function(error, newToken){
      if (error) {
        callback(error, null);
      } else if (newToken) {
        var cookies = ['OutlookConnectByJules-token=' + newToken.token.access_token + ';Max-Age=4000',
                       'OutlookConnectByJules-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                       'OutlookConnectByJules-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        callback(null, newToken.token.access_token);
      }
    });
  } else {
    // Return cached token
    var access_token = getValueFromCookie('OutlookConnectByJules-token', request.headers.cookie);
    callback(null, access_token);
  }
}
	
//access the mail api
	function mail(response, request) {
  getAccessToken(request, response, function(error, token) {
    console.log('Token found in cookie: ', token);
    var email = getValueFromCookie('OutlookConnectByJules-email', request.headers.cookie);
    console.log('Email found in cookie: ', email);
    if (token) {
      response.writeHead(200, {'Content-Type': 'text/html'});
      response.write('<div><h1>Your inbox</h1></div>');

      // Create a Graph client
      var client = microsoftGraph.Client.init({
        authProvider: (done) => {
          // Just return the token
          done(null, token);
        }
      });

      // Get the 10 newest messages
      client
        .api('/me/mailfolders/inbox/messages')
        .header('X-AnchorMailbox', email)
        .top(10)
        .select('subject,from,receivedDateTime,isRead')
        .orderby('receivedDateTime DESC')
        .get((err, res) => {
          if (err) {
            console.log('getMessages returned an error: ' + err);
            response.write('<p>ERROR: ' + err + '</p>');
            response.end();
          } else {
            console.log('getMessages returned ' + res.value.length + ' messages.');
            response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
            res.value.forEach(function(message) {
              console.log('  Subject: ' + message.subject);
              var from = message.from ? message.from.emailAddress.name : 'NONE';
              response.write('<tr><td>' + from + 
                '</td><td>' + (message.isRead ? '' : '<b>') + message.subject + (message.isRead ? '' : '</b>') +
                '</td><td>' + message.receivedDateTime.toString() + '</td></tr>');
            });

            response.write('</table>');
            response.end();
          }
        });
    } else {
      response.writeHead(200, {'Content-Type': 'text/html'});
      response.write('<p> No token found in cookie!</p>');
      response.end();
    }
  });
}