# spfx-scheduler

## Summary

Project for 30 Days of Graph - uses /me/findMeetingTimes endpoint to find upcoming meeting slots for all attendees.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution       | Author(s)                                               |
| -------------- | ------------------------------------------------------- |
| spfx-scheduler | Corrie Haffly @corriespondent                           |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | November 22, 2022 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp bundle**
  - **gulp package-solution**
- Install the solution into your SharePoint app catalog
- Approve API permissions
- Add the web part to a page

## Features

This web part allows the user to add attendees, optionally select a block of time to look for meeting times, and optionally specify the length of the meeting time. The MS Graph /me/findMeetingTimes endpoint is used to find available slots of time when everyone is available for a meeting and displays a list of times.

This project illustrates the following concepts:

- Using MS Graph in an SPFx solution
- Using the /me/findMeetingTimes endpoint to retrieve suggested meeting slots
- Using PeoplePicker and DateTimePicker controls from @pnp/sp-dev-fx-controls-react 

## References

- [#30DaysOfMSGraph](https://microsoft.github.io/30daysof/docs/roadmaps/microsoft-graph/)
- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
