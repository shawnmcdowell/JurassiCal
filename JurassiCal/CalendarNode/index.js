// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");
var outlook = require("node-outlook");
var http = require('http');

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize;
handle["/mail"] = mail;
handle["/events"] = events;
handle["/createevent"] = createevent;
handle["/postrequest"] = postrequest;

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
//    response.writeHead(302, {'Location': 'http://localhost:8010/events'});
    response.writeHead(302, {'Location': 'http://localhost:63233/app/home/home.html'});

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
    

	//Deal with POST to create
	
    var bodyData = "";
    if(request.method == 'POST') {
        request.on('data', function(data) {
            bodyData += data;
            if(bodyData.length > 1e6) {
                bodyData = "";
                response.writeHead(413, {'Content-Type': 'text/plain'}).end();
                request.connection.destroy();
            }
        });

        request.on('end', function() {
            var params = request.post = JSON.parse(bodyData);
			console.log("done");
			console.log(bodyData);
			_createOutlookEvent(outlookClient, params, response);
        });

    } else {
        response.writeHead(405, {'Content-Type': 'text/plain'});
        response.end();
    }	
      

  }
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
	
}

function _createOutlookEvent(outlookClient, params, response) {
	var ev = new outlook.Microsoft.OutlookServices.Event();
	
	////Add the event fields to an object
	//Subject
	ev.subject = params.Subject;
	
	//Body
	var body = new outlook.Microsoft.OutlookServices.ItemBody();
	body.content = params.Body
	body.contentType = outlook.Microsoft.OutlookServices.BodyType.HTML;
	ev.body = body;
	
	//Start
	var startDate = new Date(params.Start); 
	ev.start = startDate.toISOString();
	//NYI: ev.startTimeZone = "Pacific Standard Time";
	
	//End
	var endDate = new Date(params.End); 
	ev.end = endDate.toISOString();
	//NYI: ev.endTimeZone = "Pacific Standard Time";
	
	//Location
	var loc = new outlook.Microsoft.OutlookServices.Location;
	loc.displayName = params.Location;
	//NYI: loc.address = "1 Microsoft Way, Redmond, WA";
	ev.location = loc;
	
	//Attendees
	params.Attendees.forEach(function(attendeeToAdd){
		//set emailAddress object
		var emailAddress =  new outlook.Microsoft.OutlookServices.EmailAddress;
		emailAddress.address = attendeeToAdd.Address;
		//emailAddress.name = attendeeToAdd.Address;		//NYI: add the Name value to JSON object
		
		//set attendee parent object
		var attendee =  new outlook.Microsoft.OutlookServices.Attendee;
		attendee.emailAddress = emailAddress;
		
		//append each attendee object to the attendees object
		ev.attendees.push(attendee);		
	});

	//Make the call to add a calendar event
    outlookClient.me.events.addEvent(ev)
    	.then(function (result) {
			console.log("------------RESULT--------------");
			var r = new outlook.Microsoft.OutlookServices.Event(result);
		    response.writeHead(200, {"Content-Type": "text/html"});
    		console.log(JSON.stringify(r));
			console.log("------------RESULT--------------");
			response.write(JSON.stringify(result));
			response.end();
			},function (error) {
			response.writeHead(500, {"Content-Type": "text/html"});
			response.write(JSON.stringify(error));
			response.end()
		}); 

}

function postrequest(response, request) {
    var bodyData = "";

    if(request.method == 'POST') {
        request.on('data', function(data) {
            bodyData += data;
            if(bodyData.length > 1e6) {
                bodyData = "";
                response.writeHead(413, {'Content-Type': 'text/plain'}).end();
                request.connection.destroy();
            }
        });

        request.on('end', function() {
            request.post = querystring.parse(bodyData);
			console.log("done");
			console.log(bodyData);
			response.writeHead(200, {'Content-Type': 'text/plain'});
			response.write("Success");
        });

    } else {
        response.writeHead(405, {'Content-Type': 'text/plain'});
        response.end();
    }	
}
