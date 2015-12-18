// The initialize function is required for all apps.
Office.initialize = function(reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function() {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
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

function readContentControl() {
    // Run a batch operation against the Word object model.
    Word.run(function(context) {

            // Create a proxy object for the content controls collection that contains a specific tag.
            var contentControlsWithTag = context.document.contentControls.getByTag('Address');
            printData(contentControlsWithTag);

            // Queue a command to load the tag property for all of content controls. 
            context.load(contentControlsWithTag, 'tag');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function() {
                if (contentControlsWithTag.items.length === 0) {
                    console.log('No content control found.');
                    printData("no content control found");

                } else {
                    // Queue a command to get the HTML contents of the first content control.
                    var html = contentControlsWithTag.items[0].getHtml();

                    printData(html);

                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync()
                        .then(function() {
                            console.log('Content control HTML: ' + html.value);
                        });
                }
            });
        })
        .catch(function(error) {
            console.log('Error: ' + JSON.stringify(error));
                                printData(error);

            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

}
