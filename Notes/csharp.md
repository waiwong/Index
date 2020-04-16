# Note of C-Sharp

## 1. Disable Copy/Paste option from Word (Docx) document using C# and OpenXML

```csharp

DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
documentProtection.Enforcement = docProtection;

documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
documentProtection.CryptographicAlgorithmSid = 4; // SHA1
//The iteration count is unsigned
UInt32 uintVal = new UInt32();
uintVal = (uint)iterations;
documentProtection.CryptographicSpinCount = uintVal;
documentProtection.Hash = Convert.ToBase64String(generatedKey);
documentProtection.Salt = Convert.ToBase64String(arrSalt);
_objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
_objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
_objDoc.Close();

```

Ref [source](https://www.codeproject.com/Articles/1162746/Disable-Copy-Paste-option-from-Word-Docx-document)
