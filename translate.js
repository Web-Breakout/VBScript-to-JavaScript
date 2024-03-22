function translateVbsToJs(vbsCode) {
    let jsCode = vbsCode;

    // Replace VBScript Dim with JavaScript let
    jsCode = jsCode.replace(/Dim/g, 'let');

    // Replace VBScript MsgBox function with JavaScript confirm function
    jsCode = jsCode.replace(/MsgBox\("Testing screen", vbOKCancel, "Title"\)/g, 'window.confirm("Testing screen")');

    // Replace VBScript If...Then...Else structure with JavaScript if...else structure
    jsCode = jsCode.replace(/If result = vbOK Then/g, 'if (result) {');
    jsCode = jsCode.replace(/Else/g, '} else {');
    jsCode = jsCode.replace(/End If/g, '}');

    // Replace VBScript MsgBox function with JavaScript alert function
    jsCode = jsCode.replace(/MsgBox "You clicked OK!"/g, 'alert("You clicked OK!")');

    // Replace VBScript WScript.Quit function with JavaScript window.close function
    jsCode = jsCode.replace(/WScript.Quit/g, 'window.close()');

    // Replace VBScript comments with JavaScript comments
    jsCode = jsCode.replace(/' Code to execute when the OK button is clicked/g, '// Code to execute when the OK button is clicked');

    return jsCode;
}

function translateJsToVbs(jsCode) {
    let vbsCode = jsCode;

    // Replace JavaScript let with VBScript Dim
    vbsCode = vbsCode.replace(/let/g, 'Dim');

    // Replace JavaScript confirm function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/window.confirm\("Testing screen"\)/g, 'MsgBox("Testing screen", vbOKCancel, "Title")');

    // Replace JavaScript if...else structure with VBScript If...Then...Else structure
    vbsCode = vbsCode.replace(/if \(result\) \{/g, 'If result = vbOK Then');
    vbsCode = vbsCode.replace(/\} else \{/g, 'Else');
    vbsCode = vbsCode.replace(/\}/g, 'End If');

    // Replace JavaScript alert function with VBScript MsgBox function
    vbsCode = vbsCode.replace(/alert\("You clicked OK!"\)/g, 'MsgBox "You clicked OK!"');

    // Replace JavaScript window.close function with VBScript WScript.Quit function
    vbsCode = vbsCode.replace(/window.close\(\)/g, 'WScript.Quit');

    // Replace JavaScript comments with VBScript comments
    vbsCode = vbsCode.replace(/\/\/ Code to execute when the OK button is clicked/g, "' Code to execute when the OK button is clicked");

    return vbsCode;
}
