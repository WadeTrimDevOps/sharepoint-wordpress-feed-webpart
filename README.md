# SharePoint Framework (SPFx) WordPress News Feed Webpart

This SharePoint Framework (SPFx) Webpart allows you to fetch and display posts from a WordPress site as a highly configurable feed using the WordPress REST API. This webpart supports fetching either news-style posts or portfolio-style project highlights. Users can filter posts by tags, categories, post title patterns, and date ranges. This component is designed for embedding dynamic blog content from the wadetrim.com site, but can be used to render a news feed from other WordPress sites as well.

## Features

- Connect to any WordPress site URL
- Filter posts by tags and categories
- Limit posts by date range (e.g., last 7 days), or display up to years of information
- Refresh feed and clear feed settings with one click
- Fully integrated with Fluent UI for a responsive SharePoint look and feel
- Configure colors of nearly every UI element by choosing from the theme color palette
- Configure font styles for heading text

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

### For using: 
- Have a Wordpress site in mind that allows access to the /wp-json path for post consumption
- Have a SharePoint Online site

### For developing:  

- [Node.js (LTS version)](https://nodejs.org/)
- [Yeoman and SPFx Generator](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- SharePoint Online Developer site (or compatible local workbench)

## Solution

| Solution    | Authors                                               |
| ----------- | ------------------------------------------------------- |
| sharepoint-wordpress-news-feed | WadeTrimDevOps
| | repo initially forked from Austin Hunt, austinjhunt.com 

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- | 
| 1.0.0.0 | April 2025 | Initial release |
| 2.1.0.1 | July 2026 | Migrate to Heft + WT Customization |
| 2.1.0.3 | October 2026 | Add Project Highlight Fetch + Additional Customization Options |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---
 
## Further Development
- Clone this repository
- Ensure that you are at the solution folder
- in the command line run:
  - **npm install**
  - **heft start** 
