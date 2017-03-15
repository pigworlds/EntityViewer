/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItem(Office.context.mailbox.item);

            if (isPersistenceSupported()) {
                // Set up ItemChanged event
                Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
            }
        });
    };

    function displayItem(item) {
        clearOutput();
        displayEntityDocument(item);
        displayBuiltInEntities(item);
    }

    function clearOutput() {
        $('#itemId').text('');
        $('#json').text('');
        $('#entity').text('');
        $('#property').text('');
    }

    function isPersistenceSupported() {
        // This feature is part of the preview 1.5 req set
        // Since 1.5 isn't fully implemented, just check that the 
        // method is defined.
        // Once 1.5 is implemented, we can replace this with
        // Office.context.requirements.isSetSupported('Mailbox', 1.5)
        return Office.context.mailbox.addHandlerAsync !== undefined;
    }

    function itemChanged(eventArgs) {
        displayItem(Office.context.mailbox.item);
    }

    // Displays the "EntityDocument" fields, based on the current mail item
    function displayEntityDocument(item) {
        //Grab mailbox and make rest call
        var itemId = getItemRestId(item);
        $('#itemId').text(itemId);

        if (itemId != null) {
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true, asyncContext: item }, getTokenCallback);
        }
    }

    function getTokenCallback(result)
    {
        if (result.status === "succeeded")
        {
            var accessToken = result.value;
            var item = result.asyncContext;

            // Use the access token
            getCurrentItem(accessToken, item);
        }
        else
        {
            getTokenFailed(result);
        }
    }

    // Displays the built-in entities, based on the current mail item
    function displayBuiltInEntities(item)
    {
        if (item == null)
        {
            return;
        }

        var entities = item.getEntities();

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

    function getItemRestId(item)
    {
        if (item == null)
        {
            return null;
        }

        // Currently the only Outlook Mobile version that supports add-ins
        // is Outlook for iOS.
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS')
        {
            // itemId is already REST-formatted
            return item.itemId;
        }
        else
        {
            // Convert to an item ID for API v2.0
            return Office.context.mailbox.convertToRestId(
              item.itemId,
              Office.MailboxEnums.RestVersion.v2_0
            );
        }
    }

    function getRestUrl(itemId)
    {
        // Construct the REST URL to the current item
        // Details for formatting the URL can be found at 
        // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest

        var restUrl = Office.context.mailbox.restUrl;
        if (restUrl == null)
        {
            restUrl = "https://outlook.office365.com/api";
        }
        else if (!restUrl.match("/api$"))
        {
            restUrl += "/api";
        }

        return restUrl
            + "/v2.0/me/messages/"
            + itemId
            + "/?$select = id"
            + "&$expand = SingleValueExtendedProperties($filter = "
            + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name EntityDocument')"
            + " or "
            + "(PropertyId eq 'Boolean {00062008-0000-0000-c000-000000000046} Name EntityExtractionSuccess')"
            + " or "
            + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name EntityExtractionServiceDiagnosticContext')"
            + " or "
            + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name TeeVersion')"
            + ")";
    }

    function getCurrentItem(accessToken, item)
    {
        $('#json').text("calling REST API");

        // Get the item's REST ID
        var itemId = getItemRestId(item);

        var getMessageUrl = getRestUrl(itemId);

        $.ajax({ url: getMessageUrl, dataType: 'json', headers: { 'Authorization': 'Bearer ' + accessToken } })
            .done(restCallback)
            .fail(restCallbackFailed);
    }

    function restCallback(item)
    {
        var svp = item.SingleValueExtendedProperties;

        var found = false;
        if (svp != null && svp.length > 0)
        {
            for (var prop in svp)
            {
                var id = svp[prop].PropertyId;
                var value = svp[prop].Value; 
                switch (id)
                {
                    case 'String {00062008-0000-0000-c000-000000000046} Name EntityDocument':
                        found = true;
                        $('#json').JSONView(JSON.parse(value));
                        break;
                    case 'Boolean {00062008-0000-0000-c000-000000000046} Name EntityExtractionSuccess':
                        $('#property').append("<p>Entity Extraction Success: " + value + "</p>");
                        break;
                    case 'String {00062008-0000-0000-c000-000000000046} Name EntityExtractionServiceDiagnosticContext':
                        var js = $('<div />').JSONView(JSON.parse(value));
                        $('#property').append("<p>Debug context: </p>").append(js);
                        break;
                    case 'String {00062008-0000-0000-c000-000000000046} Name TeeVersion':
                        $('#property').append("<p>TeeVersion: " + value + "</p>");
                        break;
                }
            }
        }

        if (!found)
        {
            $('#json').text("no EntityDocument entities found :(");
        }
    }

    function restCallbackFailed(error)
    {
        $('#json').text("call REST failed :( " + error.responseText);
    }

    function getTokenFailed(error)
    {
        $('#json').text("get REST token failed :(");
    }
})();
