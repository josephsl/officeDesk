# Contributing to Office Desk NVDA add-on

Thank you for your interest in contributing to Office Desk NVDA add-on. The purpose of this document is to outline the overall add-on contribution requirements, development process, and offer tips when contributing.

## Overview

Office Desk is a collection of modules designed to improve support for Microsoft Office applications.

## Ways to contribute

You can contribute to Office Desk add-on in a number of ways:

1. Testing: provided that you are using supported Office releases (including latest Microsoft 365 release) with latest NVDA and the add-on installed, you can help improve the add-on by testing and reporting feedback.
2. Localization: you can translate add-on documentation and messages via NVDA translations workflow.
3. Issues and pull requests: the add-on developers welcomes issues and pull requests submitted via GitHub (see the section on issues and pull requests for more information).

## Contribution requirements

You must:

1. Be running the latest supported version of Office (as of January 2022, Office 2021 or Microsoft 365 are supported).
2. Be running the latest stable version of NVDA or later (as of January 2022, this means NVDA 2021.3 or latest alpha release).
3. If you wish to offer pull requests, you must be running latest NVDA stable version or later.

## Contribution process

### Testing the add-on

Testing the add-on simply involves using NVDA as usual. If you do encounter issues:

1. Restart NVDA with add-ons disabled.
2. Enable one add-on at a time to make sure the issue is not related to add-ons other than Office Desk.
3. If the issue occurs after enabling Office Desk, note steps to reproduce the issue.
4. Use GitHub and submit a new issue (https://github.com/josephsl/officedesk/issues/new). Be sure to include NVDA version, add-on version, Windows version, and steps to reproduce the problem.
5. Sometimes the author will ask for a debug log. If so, restart NvDA with debug logging enabled, try reproducing the issue, then either attach the debug log as part of the GitHub issue or copy and paste the relevant log fragment from the log viewer (NVDA+F1).

### Localizing add-on documentation and messages

You must do this through NVDA translations workflow, not through pull requests.

### Offering pull requests

#### Coding style

Office Desk follows NVDA's own coding style (tabs for indentation, camel case, etc.). Although you don't have to, please align closely to NVDA's own coding style.

#### Pull request process

1. If you want, create a new issue on GitHub proposing specific changes. This is so that more people can discuss changes.
2. Create an accompanying pull request via GitHub (be sure to fork the add-on source code before doing so). Unless otherwise noted, pull request base branch should be "main" and each pull request must be done from a different branch.
3. In the pull request comment, describe the pull request, including applicable Office releases (if any) such as whether the pull request addresses issues with a specific build range.
4. It can take up to 24 hours for the pull request to be reviewed and a decision is made. Depending on the severity of the issue, it can be included in the next version of the add-on or delayed for the next milestone release.
