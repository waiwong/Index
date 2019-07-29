# ** Note for backup **

## 1. Force .NET application to run as administrator

This works on Visual Studio 2008 and higher: Project + Add New Item, select "Application Manifest File". Change the &lt;equestedExecutionLevel &gt; element to:

```XML
<requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
```

The user gets the UAC prompt when they start the program. Use wisely; their patience can wear out quickly.

## 1. Javascript check exist of array of objects.

use function <strong><em>some</em></strong>.

``` Javascript
arr.some(el => el.username === name);
// for IE, not support Lambda, use function(x)
arr.some(function(el) { return el.username === name});

var f = (a) => {a.some1(); this.some2();}; 
// for IE
var f = function(a) {a.some1(); this.some2();}.bind(this);
```

## 1. execute a JavaScript function by its name as a string

``` Javacript
window["functionName"](arguments);
//if have namespace
window["My"]["Namespace"]["functionName"](arguments);
```

In order to make that easier and provide some flexibility, here is a convenience function:

``` Javascript
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

