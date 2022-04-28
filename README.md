[![Test and Release](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml/badge.svg)](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml)

# OOXML-Validator

## What is it?

The OOXML Validator is a .NET CLI package which validates Open Office XML files (.docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, or .xlam) and returns the validation errors as JSON or XML.

## How is it used?

The OOXML Validator CLI accepts 1 required and 3 optional parameters. The first argument must be the file or directory path, the order of the other 3 arguments doesn't matter.

Argument Order | Type | Value | Required/Optional
---|---|---|---
First | string | The absolute path to the file or folder to validate | Required
Second | string | The version of Office to validate against* | Optional
Third | string | If the value is `--xml`, cli will return xml | Optional
Fourth | string | If the value is `--recursive` or `-r` validates files recursively through all folders\** | Optional

\* Must be one of these: `Office2007`, `Office2010`, `Office2013`, `Office2016`, `Office2019`, `Office2021`, `Microsoft365`. Defaults to `Microsoft365`

\** Only valid if the path passed is a directory and not a file, otherwise it is ignored

## XML vs JSON

XML and JSON both return a list of validation errors. In addition to the validation error list, the XML data has 2 attributes: `FilePath` and `IsStrict`. FilePath is the path to the file that was validated and IsStrict is true if the document is saved in strict mode and false otherwise.

## Validating Directories

If the first argument passed is a directory path, then all Office Open XML files in the directory are validated. Non-OOXML files are ignored. If the `-r` or `--recursive` flags are passed then all directories are validated recursively.
