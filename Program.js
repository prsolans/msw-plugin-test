// The initialize function is required for all apps.
Office.initialize = function(reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function() {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
        var initialAddress = readContentControl("ClauseA", "cc-orig-ClauseA");
        document.getElementById("cc-orig-ClauseA").innerText = initialAddress;
        var clauseB = readContentControl("ClauseB", "cc-orig-ClauseB");
        document.getElementById("cc-orig-ClauseB").innerText = clauseB;
    });
}

function readData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(result) {
        if (result.status === "succeeded") {
            printData(result.value);
        } else {
            printData(result.error.name + ":" + err.message);
        }
    });
}

function printData(data, displayLocation) {
    {
        var printOut = "";

        for (var x = 0; x < data.length; x++) {
            for (var y = 0; y < data[x].length; y++) {
                printOut += data[x][y];
            }
        }
        document.getElementById(displayLocation).innerText = printOut;
    }
}

function compareContent(contentControl, displayLocation) {

    var result = "changed";

    var initialValue = document.getElementById("cc-orig-" + contentControl).innerText;
    if (document.getElementById(contentControl) == null) {
        readContentControl(contentControl, displayLocation);
    }
    var currentValue = document.getElementById(displayLocation).innerText;

    if (initialValue === currentValue) {
        result = "unchanged";
        document.getElementById(displayLocation).style.backgroundColor = "green";
    } else {
        highlightContentControl(contentControl);
        document.getElementById(displayLocation).style.backgroundColor = "red";
    }
    document.getElementById(displayLocation).style.color = "white";
    document.getElementById(displayLocation).innerText = result;
}

function highlightContentControl(tag) {
    // Run a batch operation against the Word object model.
    Word.run(function(context) {

            // Create a proxy object for the content controls collection that contains a specific tag.
            var contentControlsWithTag = context.document.contentControls.getByTag(tag);

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function() {
                if (contentControlsWithTag.items.length === 0) {} else {
                    contentControlsWithTag.items[0].style.color = "red";
                }
            });
        })
        .catch(function(error) {
            document.getElementById(displayLocation).innerText = 'Error: ' + JSON.stringify(error);
            if (error instanceof OfficeExtension.Error) {
                document.getElementById(displayLocation).innerText = 'Debug info: ' + JSON.stringify(error.debugInfo);
            }
        });
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
                    printOut = "There isn't a content control with a tag of " + tag + " in this document.";
                    document.getElementById(displayLocation).innerText = printOut;
                } else {
                    printOut = contentControlsWithTag.items[0].text;
                    document.getElementById("cc-" + tag).innerText = tag;
                    document.getElementById(displayLocation).innerText = printOut;
                }
            });
        })
        .catch(function(error) {
            document.getElementById(displayLocation).innerText = 'Error: ' + JSON.stringify(error);
            if (error instanceof OfficeExtension.Error) {
                document.getElementById(displayLocation).innerText = 'Debug info: ' + JSON.stringify(error.debugInfo);
            }
        });

}
