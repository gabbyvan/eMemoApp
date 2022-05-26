// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });
        $('#submit2').click(function () {
            createTemplate();
        });
        $('#SearchUser').click(function () {
            var user = document.getElementById("repUser").value;
            searchRecipient(user);
        });
        

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}

// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}

// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}

function sendSlice(slice, state) {
    var data = slice.data;
    
    // If the slice contains data, create an HTTP request.
    if (data) {
        var ToNames = $('#toAddressBox').html();
        var ccUsers = $('#CcUser').val();
        var memoBody = $('#memoBody').val();
        console.log("containED")
        var base64EncodedStr = btoa(String.fromCharCode.apply(null, new Uint8Array(data)));
        console.log(base64EncodedStr);
        sendDATA(base64EncodedStr, ToNames, ccUsers, "eMemo_doc_001", memoBody);
        
        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        
        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
      /*  var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
            
        }

        request.open("POST", "https://prod-163.westeurope.logic.azure.com:443/workflows/199bb6c9539643d99485d502b130febd/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JoxIqGPwx5cBAqaNHtSH2FFpzLZ758bYe3_WzHZZxnw");
        request.setRequestHeader("Content-Type", "application/json;charset=UTF-8");

        // Send the file as the body of an HTTP POST
        // request to the web server.

        request.send(JSON.stringify({ "name": fileData }));
        console.log(slice.index)*/

        
        


    } else {
        console.log("data was null" + console.log(JSON.stringify(data)));
    }
}

function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}

function sendDATA(datas, tos, CCs, docID,memoBODY) {
    $.ajax({
        contentType: "application/json",
        url: "https://prod-163.westeurope.logic.azure.com:443/workflows/199bb6c9539643d99485d502b130febd/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JoxIqGPwx5cBAqaNHtSH2FFpzLZ758bYe3_WzHZZxnw",
        type: "POST",
        data: JSON.stringify({ memoData: datas, To: tos, Cc: CCs, DocID: docID, memoBody: memoBODY }),
        success: function (res) {
            console.log(res);
            
        },
        error: function (res) {
            console.log(res);
            
        }
    })
}

function createTemplate() {
    // insertBody();
    var ToNames = $('#toAddressBox').html();
    /*console.log(ToNames);
    insertCC();
    insertTo(ToNames);
    insertMeMoHeader();*/
    
}
function insertMeMoHeader() {
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('Multithread ICT Solutions Memo \n', Word.InsertLocation.start);
        var Paragraph = body.paragraphs.getFirst();
        Paragraph.alignment = 'Centered';



        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added memo header.');
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
function insertTo(To) {
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('To: ' + To +' \n', Word.InsertLocation.start);
        
       
        



        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added memo header.');
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
function insertCC() {
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('CC: \n', Word.InsertLocation.start);






        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added memo header.');
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
function insertBody() {
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('Insert body of memo below: \n', Word.InsertLocation.start);
        var Paragraph = body.paragraphs.getFirst();
        Paragraph.alignment = 'Centered';
        Paragraph.font.underline = "Single";
        





        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added memo header.');
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
var array = [];
function searchRecipient(UserName) {
    array.length = 0;
    $('#liveToEmailsearch').empty();
    $.ajax({
        contentType: "application/json",
        url: "https://prod-165.westeurope.logic.azure.com:443/workflows/67d65cd109234f269adedea023a855a3/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gw-cd5SK4Qwz8B8_voyyVd8mx7lcvuizCF_oNyeGP1o",
        type: "POST",
        data: JSON.stringify({ user: UserName }),
        success: function (res) {
            console.log(JSON.stringify(res));
            array = res;
            for (var i = 0; i < array.length; i++) {
                $('#liveToEmailsearch').append('<a href="#" onclick="add(\'' + array[i] + '\')"><p id="usesrname">' + array[i] + '</p></a>');
                
            }
            var x = document.getElementById("liveToEmailsearch").style.display = "block";
            document.getElementById("liveToEmailsearch").style.border = "1px solid #A5ACB2";
        },
        error: function (res) {
            console.log(res);

        }
    })
}
var tempArray = [];
function add(setValue) {
    tempArray.push(setValue);
    $('#toAddressBox').append('' + setValue + ";" +'');
   // $("#repUser").val(tempArray.join(";"));
     var x = document.getElementById("liveToEmailsearch");
     if (x.style.display === "none") {
         x.style.display = "block";
     } else {
         x.style.display = "none";
     }
}