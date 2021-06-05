# Me solution for Teams

## Supported Versions of Office 365
Commercial|GCC|GCC High|GCC DoD
-|-|-|-
![Supported](assets/supported.png)|![Supported](assets/supported.png)|![Unknown](assets/supported.png)|![Unknown](assets/unknown-supported.png)


## SharePoint Framework Version
![version](https://img.shields.io/badge/version-1.11-green.svg)

## Overview

The Me solution for Teams will provide users a place to view all their 
personal information with the click of button. The Me solution consists of several custom SharePoint Framework Web parts (SPFX) along with out of the box SharePoint webparts to allow an organization to 
create a SharePoint page for users. 


## Minimal Path to Awesome
1. Upload the [my-email.sppkg](./solution/my-email.sppkg) to your tenant's SharePoint App Catalog.
2. Upload the [my-calendar.sppkg](./solution/my-calendar.sppkg) to your tenant's SharePoint App Catalog.
3. Upload the [my-todo.sppkg](./solution/my-todo.sppkg) to your tenant's SharePoint App Catalog.
    1. In each of the **Do you trust solution** dialog
        1. Make sure **Make this site available to all in the organization** is checked
        1. Click the deploy button
4. Create a SharePoint Page and configure with Me-Hub webparts.
5. Download the sample [Teams Manifest](./solution/Me-Hub%20Team%20Manifest/manifest.json)
6. Update manifest file links and domains to appropriate domain and pages. 
7. Create zip file with Teams manifest, [Color.png](./solution/Me-Hub%20Team%20Manifest/color.png) and [outline.png](./solution/Me-Hub%20Team%20Manifest/outline.png)
8. Upload zip file to Teams to install the app



## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
