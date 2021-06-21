# ** Note for backup **

## 1. Force .NET application to run as administrator

This works on Visual Studio 2008 and higher: Project + Add New Item, select "Application Manifest File". Change the &lt;equestedExecutionLevel &gt; element to:

```XML
<requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
```

The user gets the UAC prompt when they start the program. Use wisely; their patience can wear out quickly.

## 2. Javascript check exist of array of objects.

use function <strong><em>some</em></strong>.

```Javascript
arr.some(el => el.username === name);
// for IE, not support Lambda, use function(x)
arr.some(function(el) { return el.username === name});

var f = (a) => {a.some1(); this.some2();};
// for IE
var f = function(a) {a.some1(); this.some2();}.bind(this);
```

## 3. execute a JavaScript function by its name as a string

```Javacript
window["functionName"](arguments);
//if have namespace
window["My"]["Namespace"]["functionName"](arguments);
```

In order to make that easier and provide some flexibility, here is a convenience function:

```Javascript
function executeFunctionByName(functionName, context /*, args */) {
  var args = Array.prototype.slice.call(arguments, 2);
  var namespaces = functionName.split(".");
  var func = namespaces.pop();
  for(var i = 0; i < namespaces.length; i++) {
    context = context[namespaces[i]];
  }
  return context[func].apply(context, args);
}

// can call it like this
executeFunctionByName("My.Namespace.functionName", window, arguments);

// can pass in whatever context
executeFunctionByName("Namespace.functionName", My, arguments);
```

## 4. Vsix for SSMS

### a. How

For how to create VSIX for SSMS, refer to [this](https://stackoverflow.com/questions/55661806/how-to-create-an-extension-for-ssms-2019-v18/55661807#55661807)

If we don't have administrator permission, we could install the vsix for check. In this case, we could use debug mode, we can do is add more log for check. If need update, uninstall and install again.

Following command for install and uninstall

```Dos

:: install
"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe"  "C:\TEMP\testVsix.vsix"

:: uninstall
"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe" /u: /u:"%Identifier%"

:: find the Identifier of vsix: unzip VSIX file and look in the manifest file and there should be a Identifier key like this one:
:: <Identity Id="testVsix.7c985797-3089-440a-a54c-b0125720263d" Version="1.0.0" Language="en-US" Publisher="Wei" />

:: e.g. Identifier is testVsix.7c985797-3089-440a-a54c-b0125720263d

"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe" /u: /u:"testVsix.7c985797-3089-440a-a54c-b0125720263d"

```

### b. More

The following code use to hook the AfterExecute command in SSMS.

```CSharp

//if current user without amdin permission, the DTE2 and CommandEvenet should declare as class private member.
private DTE2 _dte；
private CommandEvents cmeExecQuery;

// add to InitializeAsync of XXXPackage class.
_dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
cmeExecQuery = _dte.Events.get_CommandEvents("{52692960-56BC-4989-B5D3-94C47A513E8D}", 1);
cmeExecQuery.AfterExecute += new _dispCommandEvents_AfterExecuteEventHandler(commandEvents_AfterExecute);

//the command handler as follow.
private void commandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
{
    Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
    DTE2 dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
    TextDocument objTextDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
    EditPoint2 startPoint = objTextDoc.StartPoint.CreateEditPoint() as EditPoint2;
    EditPoint2 endPoint = objTextDoc.EndPoint.CreateEditPoint() as EditPoint2;

    string docName = dte.ActiveDocument.Name;
    string tmpCaption = dte.ActiveDocument.ActiveWindow.Caption;
    if (docName.StartsWith("~vs"))
        docName = "NoDocName";
    string execText = execText = startPoint.GetText(endPoint).Trim();
    //TODO: log the exec text to database for further check.
}

```

## 5. Doc to Markdown

Convert .doc(x) to .md by pandoc by following command:

```Batch
pandoc -s example30.docx --no-wrap --reference-links -t markdown -o example35.md
```

> "--no-wrap" for avoid 80 characters per line.
> "--reference-links" for use reference-style links, rather than inline links

More details please refer the option page: http://pandoc.org/README.html#reader-options
