# ** Note for backup **

## 1. Force .NET application to run as administrator

This works on Visual Studio 2008 and higher: Project + Add New Item, select "Application Manifest File". Change the &lt;equestedExecutionLevel &gt; element to:

```XML
<requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
```

The user gets the UAC prompt when they start the program. Use wisely; their patience can wear out quickly.
