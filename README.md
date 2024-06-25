[![Test and Release](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml/badge.svg)](https://github.com/mikeebowen/OOXML-Validator/actions/workflows/dotnet.yml)

# OOXML-Validator

## What is it?

The OOXML Validator is a .NET CLI package which validates Open Office XML files (.docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, or .xlam) and returns the validation errors as JSON or XML.

## How is it used?

The OOXML Validator CLI accepts 1 required and 4 optional parameters. The first argument must be the file or directory path, the order of the other 4 arguments doesn't matter.

Argument Order | Type | Value | Required/Optional
---|---|---|---
First (Required) | string | The absolute path to the file or folder to validate | Required
Any | string | The version of Office to validate against* | Optional
Any | string | If the value is `--xml` or `-x`, cli will return xml | Optional
Any | string | If the value is `--recursive` or `-r` validates files recursively through all folders\** | Optional
Any | string | If the value is `--all` or `-a` files without any errors are included in the list. | Optional

\* Must be one of these (case sensitive): `Office2007`, `Office2010`, `Office2013`, `Office2016`, `Office2019`, `Office2021`, `Microsoft365`. Defaults to `Microsoft365`

\** Only valid if the path passed is a directory and not a file, otherwise it is ignored

## XML vs JSON

XML and JSON both return a list of validation errors. In addition to the validation error list, the XML data has 2 attributes: `FilePath` and `IsStrict`. FilePath is the path to the file that was validated and IsStrict is true if the document is saved in strict mode and false otherwise.

## Validating Directories

If the first argument passed is a directory path, then all Office Open XML files in the directory are validated. Non-OOXML files are ignored. If the `-r` or `--recursive` flags are passed then all directories are validated recursively.


## Development
You can run some development scripts with

```bash
./dev.sh help
# ./dev.sh <command> [build_env]
# 
#   help                 show this help message
#   docker               docker container for development
#   build <build_env>    build
#   envs                 show build envs
#   run <build_env>      run the command line
# 
```

To build run the following, replacing `linux-x64` with your environment

```bash
./dev.sh build linux-x64
```

Run the executable with

```bash
./dev.sh run linux-x64 ./test.docx
```

If you don't want to install `dotnet` just run the following to start a container shell

```bash
./dev.sh docker
# root@8b0f1aad055c:/code# [RUN YOUR COMMANDS HERE]
```