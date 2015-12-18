// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}
var MyArray = [['Berlin'],['Munich'],['Duisburg']];

function writeData() {
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: 'matrix' });
}

function ReadData() {
    Office.context.document.getSelectedDataAsync("matrix", function (result) {
        if (result.status === "succeeded"){
            printData(result.value);
        }

        else{
            printData(result.error.name + ":" + err.message);
        }
    });
}

      function printData(data) {
    {
        var printOut = "";

        for (var x = 0 ; x < data.length; x++) {
            for (var y = 0; y < data[x].length; y++) {
                printOut += data[x][y] + ",";
            }
        }
       document.getElementById("results").innerText = printOut;
    }
}
