function translateVbsToJs(vbsCode) {
    let jsCode = vbsCode;

    // Replace VBScript Dim with JavaScript let
    jsCode = jsCode.replace(/Dim/g, 'let');

    // Replace VBScript MsgBox function with JavaScript confirm function
    jsCode = jsCode.replace(/MsgBox\("Testing screen", vbOKCancel, "Title"\)/g, 'window.confirm("Testing screen");\n');

    // Replace VBScript If...Then...Else structure with JavaScript if...else structure
    jsCode = jsCode.replace(/If (\w+) = vbOK Then/g, 'if ($1) {\n');
    jsCode = jsCode.replace(/Else/g, 'Else\n');
    jsCode = jsCode.replace(/End If/g, '}\n');

    // Replace VBScript MsgBox function with JavaScript alert function
    jsCode = jsCode.replace(/MsgBox "/g, 'alert("');
    jsCode = jsCode.replace(/"/g, '");\n');

    // Replace VBScript WScript.Quit function with JavaScript window.close function
    jsCode = jsCode.replace(/WScript.Quit/g, 'window.close();\n');

    // Replace VBScript comments with JavaScript comments
    jsCode = jsCode.replace(/' Code to execute when the OK button is clicked/g, '// Code to execute when the OK button is clicked\n');

    // Replace VBScript Sub with JavaScript function
    jsCode = jsCode.replace(/Sub (\w+)\(\)\n/g, 'function $1() {\n');
    jsCode = jsCode.replace(/End Sub/g, '}\n');

    // Replace VBScript Function with JavaScript function with return statement
    jsCode = jsCode.replace(/Function (\w+)\(\)\n/g, 'function $1() {\n');
    jsCode = jsCode.replace(/End Function/g, 'return $1;\n}\n');

    return jsCode;
}

function translateJsToVbs(jsCode) {
    let vbsCode = jsCode;

    // Replace JavaScript let with VBScript Dim
    vbsCode = vbsCode.replace(/let/g, 'Dim');

    // Replace JavaScript confirm function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/window.confirm\("Testing screen"\);/g, 'MsgBox "Testing screen", vbOKCancel, "Title"\n');

    // Replace JavaScript if...else structure with VBScript If...Then...Else structure
    vbsCode = vbsCode.replace(/if \((\w+)\) \{/g, 'If $1 = vbOK Then\n');
    vbsCode = vbsCode.replace(/\} else \{/g, 'Else\n');
    vbsCode = vbsCode.replace(/\}/g, 'End If\n');

    // Replace JavaScript alert function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/alert\("/g, 'MsgBox "');
    vbsCode = vbsCode.replace(/\);/g, '"\n');

    // Replace JavaScript window.close function with VBScript WScript.Quit function
    vbsCode = vbsCode.replace(/window.close\(\);/g, 'WScript.Quit\n');

    // Replace JavaScript comments with VBScript comments
    vbsCode = vbsCode.replace(/\/\/ /g, "' ");

    // Replace JavaScript function with VBScript Sub
    vbsCode = vbsCode.replace(/function (\w+)\(\) \{\n/g, 'Sub $1()\n');
    vbsCode = vbsCode.replace(/\}\n/g, 'End Sub\n');

    // Replace JavaScript function with return statement with VBScript Function
    vbsCode = vbsCode.replace(/function (\w+)\(\) \{\nreturn \w+;\n\}\n/g, 'Function $1()\nEnd Function\n');

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
