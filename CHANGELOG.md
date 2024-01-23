# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [2.1.2] - 2024-01-23

## Updated
- Build for dotnet 6, 7, and 8

## [2.1.1] - 2024-01-22

## Updated

- Updated to OOXML SDK 3.0
- Updated to dotnet 8

## [2.1.1] - 2022-10-27

## Fixed

- Update Path property on ValidationErrorInfoInternal to have type XmlPath

## [2.1.0] - 2022-09-29

## Added

- Updated Open XML SDK Version

## [2.0.0] - 2022-07-27

## Added

- Returned XML data returns list of `<File />` elements with a child `<ValidationErrorInfoList />` element, instead of list of `<ValidationErrorInfoList />` elements.

## Fixed

- If a file cannot be opened by the Validator, a `<ValidationErrorInfoInternal />` element is added with the error message.

## [1.2.0] - 2022-04-28

### Added

- Validation can be done on all OOXML files in a single directory or recursively through all child directories

## [1.1.0] - 2022-04-11

### Added

- Validation errors can be returned as XML or JSON

## [1.0.0] - 2021-09-02

### Added

- Validates OOXML files and returns JSON string of validation errors.
