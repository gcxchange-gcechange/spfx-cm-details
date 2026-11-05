# spfx-cm-details

## Summary

Provide a view of the opportunity in the Career MarketPlace. From there, user can apply to the opportunity which is a mailto.

## Prerequisites

This web part connects to the opportunity SharePoint list.

## API permission
List of api permission that need to be approve by a sharepoint admin.

TermStore.Read.All - Office 365 SharePoint Online

## Version 

Used SharePoint Framework Webpart or Sharepoint Framework Extension 

![SPFX](https://img.shields.io/badge/SPFX-1.18.2-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-v18.17+-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Version history

Version|Date|Comments
-------|----|--------
1.0|Dec 23, 2024|Initial release
2.0|Nov 05, 2026|SPFx 1.22.2 update

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - **npm install**
  - **npm start**
- To debug in the front end:
  - go to the `serve.json` file and update `initialPage` to `https://domain-name.sharepoint.com/_layouts/15/workbench.aspx`
  - Run the command **npm start**
- To deploy: in the command-line run
  - **npx heft clean**
  - **npx heft build --production**
  - **npx heft package-solution --production**
- Add the webpart to your tenant app store
- Approve the web API permissions

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**