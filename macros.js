(function()
{
    //function to get the api txt 
    function makeRequest(url, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);

        xhr.onload = function() {
            if (xhr.status === 200) {
                // Pass the response to the callback function
                callback(xhr.responseText);
            }
            else {
                // If the request was not successful, log the error
                console.error(xhr.statusText);
            }
        };
        xhr.send();
    }

    //function to get the right string
    function findThirdString(str) {
        // Use a regular expression to match the third string
        let regex = /~(.*?)~/g;

        // Loop through the matches and return the third one
        for (let i = 0; i < 3; i += 1) {
            let matches = regex.exec(str);
            if (i === 1) {
                return matches[1];
            }
        }
    }

    //main function
    var oWorksheet = Api.GetActiveSheet();
    if (oWorksheet.GetName()==="Sheet4"){ //check if the active sheet is the right one, in this exzample in sheet4
        //console.log("666");
    }
    else{
        return;
    }
  //loop to fill cell
    var url, stockNumber, price;
    for (let run = 1; run <= 200; run++) {
        
            // Get the cell containing the stock number (e.g. A1)
            stockNumber = oWorksheet.GetRange("A" + run).GetValue();
            url = "http://qt.gtimg.cn/q=s_" + stockNumber;
            //get price for stock
            if (url != "http://qt.gtimg.cn/q=s_") {
                makeRequest(url, function(response) {
                    // Access the response here

                    price = findThirdString(response);//get the price from string
                    Api.GetActiveSheet().GetRange("B" + run).Value = price;
                    //console.log(run);
                    //console.log(response);
                    //console.log(price);
                });
            }
    }
})();
