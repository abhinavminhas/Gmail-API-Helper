# Changelog
All notable changes to this project documented here.

## [Released]
## [1.8.5](https://www.nuget.org/packages/GmailHelper/1.8.5) - 2022-07-04
### Changed
- MimeKitLite dependency update from ('3.2.0' -> '3.3.0').

## [1.8.4](https://www.nuget.org/packages/GmailHelper/1.8.4) - 2022-07-04
### Changed
- Gmail API dependency update ('1.57.0.2622' -> '1.57.0.2650').

## [1.8.3](https://www.nuget.org/packages/GmailHelper/1.8.3) - 2022-07-01
### Changed
- Gmail API dependency update ('1.56.0.2622' -> '1.57.0.2622').
- MimeKitLite dependency update ('3.1.1' -> '3.2.0').

## [1.8.2](https://www.nuget.org/packages/GmailHelper/1.8.2) - 2022-06-29
### Changed
- Gmail API dependency update ('1.56.0.2510' -> '1.56.0.2622').

## [1.8.1](https://www.nuget.org/packages/GmailHelper/1.8.1) - 2022-06-29
### Changed
- Gmail API dependency update ('1.55.0.2510' -> '1.56.0.2510').

## [1.8.0](https://www.nuget.org/packages/GmailHelper/1.8.0) - 2022-02-13
### Added
- Addition of additional optional parameter to current available extension methods that allows to keep Gmail service connection alive or dispose it (default value of this parameter is set to dispose the service connection in the method call).

### Changed
- MimeKitLite version update (3.1.0 -> 3.1.1).
- Slight refactoring & documentation updates.

## [1.7.1](https://www.nuget.org/packages/GmailHelper/1.7.1) - 2022-01-31
### Changed
- Gmail API dependency update ('1.55.0.2356' -> '1.55.0.2510').

## [1.7.0](https://www.nuget.org/packages/GmailHelper/1.7.0) - 2022-01-18
### Added
- Send email messages with attachments.

### Changed
- SendMessage() - Changed thrown exception from 'Exception' to 'FormatException' for invalid email ids.

## [1.6.0](https://www.nuget.org/packages/GmailHelper/1.6.0) - 2022-01-09
### Added
- Download message/messages attachments.

## [1.5.0](https://www.nuget.org/packages/GmailHelper/1.5.0) - 2021-12-27
### Added
- Create Gmail user labels.
- Update Gmail user labels.
- Delete Gmail user labels.
- List Gmail user labels. 

## [1.4.0](https://www.nuget.org/packages/GmailHelper/1.4.0) - 2021-12-23
### Added
- Spam/Unspam email message/messages based on query search.
- Mark email message/messages read/unread based on query search.

### Changed
- Gmail Helper documentation updates.

## [1.3.0](https://www.nuget.org/packages/GmailHelper/1.3.0) - 2021-12-19
### Added
- Untrash email message/messages and move them to inbox.

## [1.2.0](https://www.nuget.org/packages/GmailHelper/1.2.0) - 2021-12-14
### Added
- Dispose Gmail service method.

### Changed
- Improvement - Dispose Gmail service connection utilized in different extension methods.

## [1.1.0](https://www.nuget.org/packages/GmailHelper/1.1.0) - 2021-12-12
### Added
- Modify message/messages labels addition.

## [1.0.0](https://www.nuget.org/packages/GmailHelper/1.0.0) - 2021-11-01
### Added
- Gmail API helper release with certain useful extension methods to manipulate gmail messages.