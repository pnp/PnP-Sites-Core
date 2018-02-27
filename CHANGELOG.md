# OfficeDevPnP.Sites.Core Changelog
*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [2.24.1803.0 - Unreleased]

### Added
- Added SetDefaultContentType extension methods on List objects. Notice that this method behaves different from the deprecated SetDefaultContentTypeToList method. See the Deprecated section.

### Changed
- Added support for updating owners and members of an Office 365 Group.

### Deprecated
- Marked SetDefaultContentTypeToList extensions methods on List and Web objects as deprecated. This method has some flaws. It was possible to use the ID of a content type at site level to set as a default content type in the list, IF a content type in that list was inheriting from the parent content type. The new method requires you to specify the actual content type that is associated with the list. It will not work to specify a parent content type id.
