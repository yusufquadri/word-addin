(function () {

    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
                $('#getOOXMLData').click(function () { getOOXML_newAPI(); });
                $('#setOOXMLData').click(function () { setOOXML_newAPI(); });
                console.log('This code is using Word 2016 or greater.');
            } else {
                $('#getOOXMLData').click(function () { getOOXML(); });
                $('#setOOXMLData').click(function () { setOOXML(); });
                console.log('This code is using Word 2013.');
            }
        });
    };

    var currentOOXML = "";

    function postData(input) {

        $.ajax({
            type: "GET",
            url: "http://localhost:60177/spell/",
            data: { param: input },
            async: false,
            success: function (data) {
                $("#demo").html(data);
            }
        });

        /* 
        
         $.ajax({
            type: "GET",
            url: "http://localhost:51871/spell/",
            data:{param : input},
            success: function (response) {
                document.write("message sent");
            }
            
            $.ajax({
            url: "http://localhost:57371", success: function (result) {
                document.write(result);
            }
        });*/
    }

  // function callbackFunc(response) {
        // do something with the response
      //  document.write(response);
   // }

    function getOOXML_newAPI() {

        var report = document.getElementById("status");

        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        var textArea = document.getElementById("dataOOXML");

        Word.run(function (context) {

            var body = context.document.body;

            var bodyOOXML = body.getOoxml();
           // var bodyHTML = body.getHtml();
          //  report.innerText = bodyOOXML;

            return context.sync().then(function () {
                // currentooxml contains the entire text
                currentOOXML = bodyOOXML.value;
                //var currentHTML = bodyHTML.value;
                //report.innerText = currentHTML;
             //report.innerText = currentOOXML;
               // document.write(currentOOXML);
                if (window.DOMParser) {
                    parser = new DOMParser();
                    xmlDoc = parser.parseFromString(currentOOXML, "application/xml");
                    //document.write("here");
                }
                else // Internet Explorer
                {
                    xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                    xmlDoc.async = false;
                    xmlDoc.loadXML(currentOOXML);
                }
                /*for (var i = 0; i < 3; i++) {
                    var tt = xmlDoc.getElementsByTagName("w:t")[i].childNodes[0].nodeValue;
                    report.innerText = tt;
                }
                var i;
                var len = xmlDoc.getElementsByTagName("w:t").len;
              //  alert(len);
                for (i = 0; i < len; i++) {
                    document.getElementById("demo").innerHTML =
                    xmlDoc.getElementsByTagName("w:t")[1].childNodes[0].nodeValue;
                }*/
                var x, i, txt;
              //  xmlDoc = xml.responseXML;
                txt = "";
                x = xmlDoc.getElementsByTagName("w:t");
                for (i = 0; i < x.length; i++) {
                    txt += x[i].childNodes[0].nodeValue + " ";
                }
                document.getElementById("demo").innerHTML = txt;

               

                postData(txt);
               // document.write('done and dusted');
                while (textArea.hasChildNodes()) {
                    textArea.removeChild(textArea.lastChild);
                };

                setTimeout(function () {
                    textArea.appendChild(document.createTextNode(currentOOXML));
                    report.innerText = "The getOOXML function succeeded!";
                }, 400);
               // document.write("outta nowhere");
                setTimeout(function () {
                    report.innerText = "";
                }, 200);

            });
        })
        .catch(function (error) {

            currentOOXML = "";
            report.innerText = error.message;

            console.log("Error: " + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function setOOXML_newAPI() {
        var report = document.getElementById("status");

        currentOOXML = document.getElementById("dataOOXML").textContent

        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        if (currentOOXML != "") {

            Word.run(function (context) {

                var body = context.document.body;

                body.insertOoxml(currentOOXML, Word.InsertLocation.end);
                body.select();
                return context.sync().then(function () {

                    report.innerText = "The setOOXML function succeeded!";
                    setTimeout(function () {
                        report.innerText = "";
                    }, 2000);
                });
            })
            .catch(function (error) {

                while (textArea.hasChildNodes()) {
                    textArea.removeChild(textArea.lastChild);
                }
                report.innerText = error.message;

                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });

        } else {
            report.innerText = 'Add some OOXML data before trying to set the contents.';
        }
    }

    function getOOXML() {
        var report = document.getElementById("status");
        var textArea = document.getElementById("dataOOXML");
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        Office.context.document.getSelectedDataAsync("ooxml",
            function (result) {

                if (result.status == "succeeded") {

                    currentOOXML = result.value;

                    while (textArea.hasChildNodes()) {
                        textArea.removeChild(textArea.lastChild);
                        report.innerText = "";
                    };
                    setTimeout(function () {
                        textArea.appendChild(document.createTextNode(currentOOXML));
                        report.innerText = "The getOOXML function succeeded!";
                    }, 400);

                    setTimeout(function () {
                        report.innerText = "";
                    }, 2000);
                }
                else {
                    currentOOXML = "";
                    report.innerText = result.error.message;
                }
            });
    }

    function setOOXML() {
        var report = document.getElementById("status");

        var report1 = report;

        currentOOXML = document.getElementById("dataOOXML").textContent

        while (report.hasChildNodes())
        {
            report.removeChild(report.lastChild);
        }

        if (currentOOXML != "") {
            Office.context.document.setSelectedDataAsync(
                currentOOXML, { coercionType: "ooxml" },
                function (result) {
                    if (result.status == "succeeded") {
                        report.innerText = "The setOOXML function succeeded!";
                        setTimeout(function () {
                            report.innerText = "";
                        }, 2000);
                    }
                    else {
                        report.innerText = result.error.message;

                        while (textArea.hasChildNodes()) {
                            textArea.removeChild(textArea.lastChild);
                        }
                    }
                });
        }
        else {

            report.innerText = "There is currently no OOXML to insert!"
                + " Please select some of your document and click [Get OOXML] first!";
        }
    }
})();
