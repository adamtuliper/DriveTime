// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />

(function () {
    "use strict";



    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#findRoute').click(createAppointments);
            checkLocation();
            loadProps();

        });
    };

    function addMinutes(date, minutes) {
        return new Date(date.getTime() + minutes * 60000);
    }


    function createAppointmentByEws(start, end, subject, item)
    {
        var requestPreAppt = '<?xml version="1.0" encoding="utf-8"?>' +
       '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
       'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
       'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
       'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
  '<soap:Header>' +
    '<t:RequestServerVersion Version="Exchange2007_SP1" />' +
    '<t:TimeZoneContext>' +
      '<t:TimeZoneDefinition Id="Pacific Standard Time" />' +
    '</t:TimeZoneContext>' +
  '</soap:Header>' +
  '<soap:Body>' +
    '<m:CreateItem SendMeetingInvitations="SendToNone">' +
      '<m:Items>' +
        '<t:CalendarItem>' +
          '<t:Subject>' + subject + '</t:Subject>' +
          '<t:Body BodyType="HTML">' + 'DriveTime Related Item:' + item.itemId + '</t:Body>' +
          //'<t:ReminderDueBy>2013-09-19T14:37:10.732-07:00</t:ReminderDueBy>' +
          '<t:Start>' + start.toISOString() + '</t:Start>' +
          '<t:End>' + end.toISOString() + '</t:End>' +
          '<t:Location></t:Location>' +
          '<t:MeetingTimeZone TimeZoneName="Pacific Standard Time" />' +
        '</t:CalendarItem>' +
      '</m:Items>' +
    '</m:CreateItem>' +
  '</soap:Body>' +
'</soap:Envelope>'

        
        Office.context.mailbox.makeEwsRequestAsync(requestPreAppt,
        function (asyncResult, result) {
            if (asyncResult.status == "failed") {
                console.log("Action failed with error: " + asyncResult.error.message);
            } else {
                console.log("Success: " + JSON.stringify(asyncResult));
            }
        });

    }
    function createAppointments() {

        //  map.entities.clear(); 
        //  var geoLocationProvider = new Microsoft.Maps.GeoLocationProvider(map);  
        // geoLocationProvider.getCurrentPosition();
        var item = Office.context.mailbox.item;

        //alert(minutes);
        console.log('item...');
        console.log(item);
        var startTime = item.start;
        if (!startTime) {
            console.log('No start time found on item');
            return;
        }
        var endTime = item.end;
        console.log('startTime: ' + startTime);

        //var start = startTime.getHours();
        //var end = new Date();

        var millis = item.end - item.start;
        var minutes = millis / 1000 / 60;

        //end.setMinutes(end.getMinutes() + minutes);
        //item.end.setMinutes(-minutes);
        var start = item.start;
        console.log(start);
        //start.setMinutes(-minutes);
        minutes = 15;
        start.setMinutes(-minutes);
        var end = item.start;

        console.log('appt start:' + start);
        console.log('appt end:' + end)


        //var request = '<?xml version="1.0" encoding="utf-8"?>' +
        //'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <soap:Header>' +
        //'<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        //'</soap:Header>' +
        //'<soap:Body>' +
        //'<GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <ItemShape>' +
        //'<t:BaseShape>IdOnly</t:BaseShape> <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
        //'</ItemShape>' +
        //'<ItemIds>' +
        //'<t:ItemId Id="' + Office.context.mailbox.item.itemId + '"/>' +
        //'</ItemIds>' +
        //'</GetItem>' +
        //'</soap:Body>' +
        //'</soap:Envelope>';


        createAppointmentByEws(start, end, 'DriveTime - Commute Time Pre', item);


        console.log('made request');
       // sendRequest();
        //Office.context.mailbox.displayNewAppointmentForm(
        //  {
        //      requiredAttendees: [Office.context.mailbox.userProfile.emailAddress],
        //      optionalAttendees: [''],
        //      start: start,
        //      end: item.start,
        //      location: '',
        //      resources: [''],
        //      subject: 'DriveTime - Commute Time Pre',
        //      body: 'DriveTime Related Item:' + item.itemId
        //  });

        //console.log('setting timeout');

        var postEndTime = endTime;
        //postEndTime.setMinutes(minutes);
        postEndTime.setMinutes(15);

        createAppointmentByEws(item.end, postEndTime, 'DriveTime - Commute Time Post', item);

        //setTimeout(function () {

        //    Office.context.mailbox.displayNewAppointmentForm(
        //   {
        //       requiredAttendees: [Office.context.mailbox.userProfile.emailAddress],
        //       optionalAttendees: [''],
        //       start: item.end,
        //       end: postEndTime,
        //       location: '',
        //       resources: [''],
        //       subject: 'DriveTime - Commute Time Post',
        //       body: 'DriveTime Related Item:' + item.itemId
        //   });

        //}, 2000);

        setTimeout(function () {
            window.location.href = '/completed.html';
        }, 2000);
    }


    function getSoapEnvelope(request) {
        // Wrap an Exchange Web Services request in a SOAP envelope. 
        var result =

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '  <t:RequestServerVersion Version="Exchange2013"/>' +
        '  </soap:Header>' +
        '  <soap:Body>' +

        request +

        '  </soap:Body>' +
        '</soap:Envelope>';

        return result;
    };

    function getSubjectRequest(id) {
        // Return a GetItem EWS operation request for the subject of the specified item.  
        var result =

     '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
     '      <ItemShape>' +
     '        <t:BaseShape>IdOnly</t:BaseShape>' +
     '        <t:AdditionalProperties>' +
     '            <t:FieldURI FieldURI="item:Subject"/>' +
     '        </t:AdditionalProperties>' +
     '      </ItemShape>' +
     '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
     '    </GetItem>';

        return result;
    };

    // Send an EWS request for the message's subject. 
    function sendRequest() {
        // Create a local variable that contains the mailbox. 
        var mailbox = Office.context.mailbox;
        var request = getSubjectRequest(mailbox.item.itemId);
        var envelope = getSoapEnvelope(request);

        mailbox.makeEwsRequestAsync(envelope, callback);
    };

    // Function called when the EWS request is complete. 
    function callback(asyncResult) {
        var response = asyncResult.value;
        var context = asyncResult.context;

        // Process the returned response here. 
        //var responseSpan = document.getElementById("response");
        //responseSpan.innerText = response;
        console.log(response);

    };


    // Take an array of AttachmentDetails objects and
    // build a list of attachment names, separated by a line-break
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from a Message object
    function loadMessageProps(item) {
        $('#message-props').show();

        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#conversationId').text(item.conversationId);
        $('#from').html(buildEmailAddressString(item.from));
        $('#internetMessageId').text(item.internetMessageId);
        $('#normalizedSubject').text(item.normalizedSubject);
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));
    }

    // Load properties from an Appointment object
    function loadAppointmentProps(item) {
        $('#appointment-props').show();

        $('#appt-attachments').html(buildAttachmentsString(item.attachments));
        $('#end').text(item.end.toLocaleString());
        $('#location').text(item.location);
        $('#appt-normalizedSubject').text(item.normalizedSubject);
        $('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
        $('#organizer').html(buildEmailAddressString(item.organizer));
        $('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));
        $('#resources').html(buildEmailAddressesString(item.resources));
        $('#start').text(item.start.toLocaleString());
        $('#appt-subject').text(item.subject);
    }


    function checkLocation() {
        var item = Office.context.mailbox.item;





        //if(!item.location)
        if (true) {
            item.body.getAsync('text', function (result) {
                if (result.status === 'succeeded') {
                    var body = result.value;

                    // var start = new Date();
                    // var end = new Date();
                    // end.setHours(start.getHours() + 1);

                    // Office.context.mailbox.displayNewAppointmentForm(
                    //   {
                    //     requiredAttendees: [''], 
                    //     optionalAttendees: [''], 
                    //     start: start, 
                    //     end: end,// item.end, 
                    //     location: 'Travel', 
                    //     resources: [''], 
                    //     subject: 'meeting', 
                    //     body: 'Travel spot booked by DriveTime!'
                    //   });


                    Office.context.mailbox.item.notificationMessages.addAsync("testttt", {
                        type: "informationalMessage",
                        icon: "blue-icon-16",
                        message: 'test notification - item end:' + item.end.toLocaleString(),
                        persistent: false
                    });

                    //TODO determine if address is in body

                    //$('#bodyText').text(result.value);
                }

            });
        }

        $('#startingAddress').val('24832 Via Del Rio, Lake Forest CA 92630');
        $('#destinationAddress').val(item.location);


    }
    // Load properties from the Item base object, then load the
    // type-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;

        $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        $('#itemClass').text(item.itemClass);
        $('#itemId').text(item.itemId);
        $('#itemType').text(item.itemType);

        item.body.getAsync('html', function (result) {
            if (result.status === 'succeeded') {
                $('#bodyHtml').text(result.value);
            }
        });

        item.body.getAsync('text', function (result) {
            if (result.status === 'succeeded') {
                $('#bodyText').text(result.value);
            }
        });

        if (item.itemType == Office.MailboxEnums.ItemType.Message) {
            loadMessageProps(item);
        }
        else {
            loadAppointmentProps(item);
        }
    }
})();

// MIT License: 

// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 

// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.