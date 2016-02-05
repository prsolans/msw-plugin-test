// The initialize function is required for all apps.
Office.initialize = function(reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function() {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
        var initialAddress = readContentControl("Address", "results");
        document.getElementById("results").innerText = initialAddress;
    });
}

function ReadData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(result) {
        if (result.status === "succeeded") {
            printData(result.value);
        } else {
            printData(result.error.name + ":" + err.message);
        }
    });
}

function printData(data) {
    {
        var printOut = "";

        for (var x = 0; x < data.length; x++) {
            for (var y = 0; y < data[x].length; y++) {
                printOut += data[x][y];
            }
        }
        document.getElementById("results").innerText = printOut;
    }
}

function readContentControl(tag, displayLocation) {

    var printOut = "";

    // Run a batch operation against the Word object model.
    Word.run(function(context) {

            // Create a proxy object for the content controls collection that contains a specific tag.
            var contentControlsWithTag = context.document.contentControls.getByTag(tag);

            // Queue a command to load the text property for all of content controls with a specific tag. 
            context.load(contentControlsWithTag, 'text');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function() {
                if (contentControlsWithTag.items.length === 0) {
                    printOut = "There isn't a content control with a tag of " + thisTag + " in this document.";
                    document.getElementById(displayLocation).innerText = printOut;
                } else {
                    printOut = contentControlsWithTag.items[0].text;
                    document.getElementById(displayLocation).innerText = printOut;
                }

            });
        })
        .catch(function(error) {
            document.getElementById("control-results").innerText = 'Error: ' + JSON.stringify(error);
            if (error instanceof OfficeExtension.Error) {
                document.getElementById("control-results").innerText = 'Debug info: ' + JSON.stringify(error.debugInfo);
            }
        });

}
