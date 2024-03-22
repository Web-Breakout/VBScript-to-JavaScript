// The translateVbsToJs and translateJsToVbs functions
function translateVbsToJs(vbsCode) {
    let jsCode = vbsCode;

    // Replace VBScript Dim with JavaScript let
    jsCode = jsCode.replace(/Dim/g, 'let');

    // Replace VBScript MsgBox function with JavaScript confirm function
    jsCode = jsCode.replace(/MsgBox\("Testing screen", vbOKCancel, "Title"\)/g, 'window.confirm("Testing screen")');

    // Replace VBScript If...Then...Else structure with JavaScript if...else structure
    //If the variable is named something other than 'result', it will not work
    jsCode = jsCode.replace(/If result = vbOK Then/g, 'if (result) {');
    jsCode = jsCode.replace(/If result2 = vbOK Then/g, 'if (result2) {');
    jsCode = jsCode.replace(/If result3 = vbOK Then/g, 'if (result3) {');
    jsCode = jsCode.replace(/If result4 = vbOK Then/g, 'if (result4) {');
    jsCode = jsCode.replace(/Else/g, '} else {');
    jsCode = jsCode.replace(/End If/g, '}');

    // Replace VBScript MsgBox function with JavaScript alert function
    jsCode = jsCode.replace(/MsgBox "/g, 'alert("');
    jsCode = jsCode.replace(/"/g, '")');

    // Replace VBScript WScript.Quit function with JavaScript window.close function
    jsCode = jsCode.replace(/WScript.Quit/g, 'window.close()');

    // Replace VBScript comments with JavaScript comments
    jsCode = jsCode.replace(/' Code to execute when the OK button is clicked/g, '// Code to execute when the OK button is clicked');

    // Replace VBScript Sub with JavaScript function
    jsCode = jsCode.replace(/Sub /g, 'function ');
    jsCode = jsCode.replace(/Function /g, 'function ');
    return jsCode;
}

function translateJsToVbs(jsCode) {
    let vbsCode = jsCode;

    // Replace JavaScript let with VBScript Dim
    vbsCode = vbsCode.replace(/let/g, 'Dim');

    // Replace JavaScript confirm function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/window.confirm\("Testing screen"\)/g, 'MsgBox("Testing screen", vbOKCancel, "Title")');

    // Replace JavaScript if...else structure with VBScript If...Then...Else structure
    // If the variable is named something other than 'result', it will not work
    vbsCode = vbsCode.replace(/if \(result\) \{/g, 'If result = vbOK Then');
    vbsCode = vbsCode.replace(/if \(result2\) \{/g, 'If result2 = vbOK Then');
    vbsCode = vbsCode.replace(/if \(result3\) \{/g, 'If result3 = vbOK Then');
    vbsCode = vbsCode.replace(/if \(result4\) \{/g, 'If result4 = vbOK Then');
    vbsCode = vbsCode.replace(/\} else \{/g, 'Else');
    vbsCode = vbsCode.replace(/\}/g, 'End If');

    // Replace JavaScript alert function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/alert\("/g, 'MsgBox "');
    vbsCode = vbsCode.replace(/\)"/g, '"');

    // Replace JavaScript window.close function with VBScript WScript.Quit function
    vbsCode = vbsCode.replace(/window.close\(\)/g, 'WScript.Quit');

    // Replace JavaScript comments with VBScript comments
    vbsCode = vbsCode.replace(/\/\/ /g, "' ");

    // Replace JavaScript function with VBScript Sub
    vbsCode = vbsCode.replace(/function /g, 'Sub ');
    return vbsCode;
}

// The sanitizeInput function
function sanitizeInput(js) {
    // Certain inputs are not allowed
    js = js.replace(/eval\(/g, 'var usnafe = (');
    js = js.replace(/var/g, 'let');
    js = js.replace(/const/g, 'let');
    js = js.replace(/'use-strict'/g, "' This file originated from a strict JavaScript file.");
    return js;
}

// The downloadCode function

function downloadCode(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);

    element.style.display = 'none';
    document.body.appendChild(element);

    element.click();

    document.body.removeChild(element);
}

// The handleFileSelect function

function handleFileSelect(evt) {
    var file = evt.target.files[0];
    if (!file) {
        return;
    }

    var reader = new FileReader();
    reader.onload = function(e) {
        var contents = e.target.result;
        var jsCode = translateVbsToJs(contents);
        jsCode = sanitizeInput(jsCode);

        // Run the JavaScript code
        try {
            eval(jsCode);
        } catch (error) {
            console.error("An error occurred while executing the code:", error);
        }

        // Download the JavaScript code
        downloadCode('translated.js', jsCode);

        // Translate the JavaScript code back to VBScript and download it
        var vbsCode = translateJsToVbs(jsCode);
        downloadCode('translated.vbs', vbsCode);
    };
    reader.readAsText(file);
}

document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);
