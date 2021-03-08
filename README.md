# OOXML-Validator

## What is it?

The OOXML Validator is a .NET package for validating Open Office XML files (.docx, .pptx, .xlsx). 

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
### OOXMLValidator.FormatVersion Enum
An enum representing the version of office to validate against
#### FormatVersion
```csharp
enum FormatVersion
{
    Office2007,
    Office2010,
    Office2013,
    Office2016,
    Office2019
}
```