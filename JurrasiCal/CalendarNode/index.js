// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");
var outlook = require("node-outlook");

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize;
handle["/mail"] = mail;
handle["/events"] = events;
handle["/createevent"] = createevent;

server.start(router.route, handle);

function home(response, request) {
  console.log("Request handler 'home' was called.");
  response.writeHead(200, {"Content-Type": "text/html"});
  response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 account.</p>');
  response.end();
}

var url = require("url");
function authorize(response, request) {
  console.log("Request handler 'authorize' was called.");
  
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(request.url, true);
  var code = url_parts.query.code;
  console.log("Code: " + code);
  var token = authHelper.getTokenFromCode(code, 'https://outlook.office365.com/', tokenReceived, response);
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log("Access token error: ", error.message);
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  }
  else {
    response.setHeader('Set-Cookie', ['node-tutorial-token =' + token.token.access_token + ';Max-Age=3600']);
    response.writeHead(302, {'Location': 'http://localhost:8010/events'});
    response.end();
  }
}

function mail(response, request) {
  var cookieName = 'node-tutorial-token';
  var cookie = request.headers.cookie;
  if (cookie && cookie.indexOf(cookieName) !== -1) {
    console.log("Cookie: ", cookie);
    // Found our token, extract it from the cookie value
    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    var token = cookie.substring(start, end);
    console.log("Token found in cookie: " + token);
    
    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0', 
      authHelper.getAccessTokenFn(token));
    
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<div><span>Your inbox</span></div>');
    response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
    
    outlookClient.me.messages.getMessages()
    .orderBy('DateTimeReceived desc')
    .select('DateTimeReceived,From,Subject').fetchAll(10).then(function (result) {
      result.forEach(function (message) {
        var from = message.from ? message.from.emailAddress.name : "NONE";
        response.write('<tr><td>' + from + 
          '</td><td>' + message.subject +
          '</td><td>' + message.dateTimeReceived.toString() + '</td></tr>');
      });
      
      response.write('</table>');
      response.end();
    },function (error) {
      console.log(error);
      response.write("<p>ERROR: " + error + "</p>");
      response.end();
    });
  }
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}

function events(response, request) {
  var cookieName = 'node-tutorial-token';
  var cookie = request.headers.cookie;
  if (cookie && cookie.indexOf(cookieName) !== -1) {
    console.log("Cookie: ", cookie);
    // Found our token, extract it from the cookie value
    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    var token = cookie.substring(start, end);
    console.log("Token found in cookie: " + token);
    
    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0', 
      authHelper.getAccessTokenFn(token));
    
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<div><span>Your events</span></div>');
    response.write('<table><tr><th>Start</th><th>End</th><th>Subject</th><th>Location</th><th>Attendees</th></tr>');
    
    outlookClient.me.events.getEvents()
    .orderBy('Start desc')
    .select('Start,End,Subject,Location,Attendees,Body').fetchAll(10).then(function (result) {
      result.forEach(function (event) {
        //var from = message.from ? message.from.emailAddress.name : "NONE";
		var attendees=[];
		event.attendees.forEach(function (attendee) {
			var disp = attendee._EmailAddress._Name;
			var email = attendee._EmailAddress._Address;
			attendees.push(disp + "(" + email + ")");
		});
        response.write('<tr><td>' + event.start + 
          '</td><td>' + event.end +
          '</td><td>' + event.subject + 
          '</td><td>' + event.location._DisplayName +
          '</td><td>' + attendees.toString() +
		  '</td></tr>');
		response.write('<tr><td>' + event.body.content + '</tr></td>');

      });
      
      response.write('</table>');
      response.end();
    },function (error) {
      console.log(error);
      response.write("<p>ERROR: " + error + "</p>");
      response.end();
    });
  }
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}

function createevent(response, request) {
  var cookieName = 'node-tutorial-token';
  var cookie = request.headers.cookie;
  if (cookie && cookie.indexOf(cookieName) !== -1) {
    console.log("Cookie: ", cookie);
    // Found our token, extract it from the cookie value
    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    var token = cookie.substring(start, end);
    console.log("Token found in cookie: " + token);
    
    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0', 
      authHelper.getAccessTokenFn(token));
    
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<div><span>Creating event</span></div>');
    response.write('<table><tr><th>Start</th><th>End</th><th>Subject</th><th>Location</th><th>Attendees</th></tr>');
    

	var ev = new outlook.Microsoft.OutlookServices.Event();
	ev.subject = "Test event";
	//ev.body.contentType = "HTML";
	var body = new outlook.Microsoft.OutlookServices.ItemBody();
	body.content = "Test Body";
	body.contentType = "HTML";
	//ev.body = body;
	//ev.start = new Date("6/25/2105 8:00");
	//ev.startTimeZone = "Pacific Standard Time";
	//ev.end = new Date("2014-02-25T19:00:00-08:00");
	ev.subject = "Test Event";
	//ev.location = "Test Location";
	var attendee =  new outlook.Microsoft.OutlookServices.Attendee;
	var emailAddress =  new outlook.Microsoft.OutlookServices.EmailAddress;
	emailAddress.address = "shawnmc@awesome.onmicrosoft.com";
	emailAddress.name = "Shawn Test";
	attendee.emailAddress = emailAddress;
	//ev.attendees = attendee;

    outlookClient.me.events.addEvent(ev)
    	.then(function (result) {
    		console.log(JSON.stringify(result));
			response.write('<P>Success</P>');
			response.write('</table>');
			response.end();
			},function (error) {
			console.log(error);
			response.write("<p>ERROR: " + JSON.stringify(error) + "</p>");
			response.end()
		}); 
      
      

  }
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
	
}
