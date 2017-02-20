/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayEntityDocument();

            displayBuiltInEntities();
        });
    };

    // Displays the "EntityDocument" fields, based on the current mail item
    function displayEntityDocument() {

        //Grab mailbox and make EWS 
        var mailbox = Office.context.mailbox;
        $('#itemId').text(mailbox.item.itemId);
        var request = getItemRequest(mailbox.item.itemId);
        var envelope = getSoapEnvelope(request);

        mailbox.makeEwsRequestAsync(envelope, ewsCallback);
    }

    // Displays the built-in entities, based on the current mail item
    function displayBuiltInEntities()
    {
        var entities = Office.context.mailbox.item.getEntities();

        displayEntities(entities.addresses, "Addresses");
        displayEntities(entities.contacts, "Contacts");
        displayEntities(entities.emailAddresses, "Email Addresses");
        displayEntities(entities.meetingSuggestions, "Meeting Suggestions");
        displayEntities(entities.phoneNumbers, "Phone Numbers");
        displayEntities(entities.taskSuggestions, "Task Suggestions");
        displayEntities(entities.urls, "URLs");
    }

    function displayEntities(entities, typeName)
    {
        if (entities == null || entities.length == 0)
        {
            return;
        }
        
        $('#entity').append("<p>" + typeName + "</p>");
        var appendText = "<ul>";
        for (var i = 0; i < entities.length; i++)
        {
            appendText += "<li>";
            appendText += JSON.stringify(entities[i]);
            appendText += "</li>";
        }
        appendText += "</ul>";
        $('#entity').append(appendText);
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
        '    <t:RequestServerVersion Version="Exchange2010"/>' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        request +
        '  </soap:Body>' +
        '</soap:Envelope>';

        return result;
    };

    function getItemRequest(id) {
        // Return a GetItem EWS operation request for the subject of the specified item. 
        var result =
        '<GetItem ' +
        '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '        <t:AdditionalProperties>' +
        '            <t:ExtendedFieldURI DistinguishedPropertySetId="Common"' +
        '                                PropertyName="EntityDocument"' +
        '                                PropertyType="String"/>' +
        '        </t:AdditionalProperties>' +
        '      </ItemShape>' +
        '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
        '    </GetItem>';    
        return result;
    };

    function ewsCallback(asyncResult) {
        var response = asyncResult.value;
        var context = asyncResult.context;

        var $xml = $(response);

        var jsonString = $xml.find('t\\:Value').text();

        if (!jsonString)
            $('#json').text("no EntityDocument entities found :(");
        else
            $('#json').html( JSON.stringify(JSON.parse(jsonString), null, '    '));
    
    }
})();