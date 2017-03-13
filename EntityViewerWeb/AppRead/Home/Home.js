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
        //Grab mailbox and make rest call
        var itemId = getItemRestId();
        $('#itemId').text(itemId);

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, getTokenCallback);
    }

    function getTokenCallback(result)
    {
        if (result.status === "succeeded")
        {
            var accessToken = result.value;

            // Use the access token
            getCurrentItem(accessToken);
        }
        else
        {
            getTokenFailed(result);
        }
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

    function getItemRestId()
    {
        // Currently the only Outlook Mobile version that supports add-ins
        // is Outlook for iOS.
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS')
        {
            // itemId is already REST-formatted
            return Office.context.mailbox.item.itemId;
        }
        else
        {
            // Convert to an item ID for API v2.0
            return Office.context.mailbox.convertToRestId(
              Office.context.mailbox.item.itemId,
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
            + '/v2.0/me/messages/'
            + itemId
            + "/?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00062008-0000-0000-C000-000000000046} Name EntityDocument')";
    }

    function getCurrentItem(accessToken)
    {
        $('#json').text("calling REST API");

        // Get the item's REST ID
        var itemId = getItemRestId();

        var getMessageUrl = getRestUrl(itemId);

        $.ajax({ url: getMessageUrl, dataType: 'json', headers: { 'Authorization': 'Bearer ' + accessToken } })
            .done(restCallback)
            .fail(restCallbackFailed);
    }

    function restCallback(item)
    {
        var svp = item.SingleValueExtendedProperties;

        if (svp == null || svp.length == 0)
        {
            $('#json').text("no EntityDocument entities found :(");
        }
        else
        {
            var jsonString = svp[0].Value;
            $('#json').JSONView(JSON.parse(jsonString));
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
