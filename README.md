[![Test and Release](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml/badge.svg)](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml)

# OOXML-Validator

## What is it?

The OOXML Validator is a .NET CLI package for validating Open Office XML files (.docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, or .xlam).

## How is it used?

Clone the repository, build the project and add the .dll to your project. This gives access to the `OOXMLValidator` namespace.

### OOXMLValidator.OOXML Method

Returns an IEnumerable containing the ValidationErrorInfo instances with any errors from the document.

#### OOXML(string, FormatVersion)

```csharp
public static IEnumerable<ValidationErrorInfo> OOXML(string fileName, FormatVersion? format)
```
##### Parameters
`fileName` string
A string representing the absolute path to the OOXML file to validate
`format` FormatVersion?

The `OOXMLValidator.FormatVersion` to use for validation
##### Returns
`IEnumerable<ValidationErrorInfo>`
### DocumentFormat.OpenXml.FormatVersion Enum
An enum representing the version of office to validate against
#### FormatVersion
```csharp
//
// Summary:
//     Defines the Office Open XML file format version.
[Flags]
public enum FileFormatVersions
{
    //
    // Summary:
    //     Represents an invalid value which will cause an exception.
    None = 0x0,
    //
    // Summary:
    //     Represents Microsoft Office 2007.
    Office2007 = 0x1,
    //
    // Summary:
    //     Represents Microsoft Office 2010.
    Office2010 = 0x2,
    //
    // Summary:
    //     Represents Microsoft Office 2013.
    Office2013 = 0x4,
    //
    // Summary:
    //     Represents Microsoft Office 2016.
    Office2016 = 0x8,
    //
    // Summary:
    //     Represents Microsoft Office 2019.
    Office2019 = 0x10,
    //
    // Summary:
    //     Represents Microsoft Office 2021.
    Office2021 = 0x20,
    //
    // Summary:
    //     Represents Microsoft 365.
    Microsoft365 = 0x40000000
}
```