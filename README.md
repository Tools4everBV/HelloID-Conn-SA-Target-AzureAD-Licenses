<!-- Description -->
## Description
This connector contains tasks that will query the available licenses form an Azure tenant.
The information we can query from Azure can then be processed and when a set threshold is reached we can act on this.

## Versioning
| Version | Description | Date |
| - | - | - |
| 1.0.0   | Initial release | 2021/12/20  |

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Table of Contents](#table-of-contents)
- [Introduction](#introduction)
- [Getting the Azure AD graph API access](#getting-the-azure-ad-graph-api-access)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Authentication and Authorization](#authentication-and-authorization)
  - [Connection settings](#connection-settings)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)

<!-- GETTING STARTED -->
## Getting the Azure AD graph API access

By using this connector you will have the ability to manage Azure AD Guest accounts.

### Application Registration
The first step to connect to Graph API and make requests, is to register a new <b>Azure Active Directory Application</b>. The application is used to connect to the API and to manage permissions.

* Navigate to <b>App Registrations</b> in Azure, and select “New Registration” (<b>Azure Portal > Azure Active Directory > App Registration > New Application Registration</b>).
* Next, give the application a name. In this example we are using “<b>HelloID PowerShell</b>” as application name.
* Specify who can use this application (<b>Accounts in this organizational directory only</b>).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “<b>Register</b>” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

To assign your application the right permissions, navigate to <b>Azure Portal > Azure Active Directory >App Registrations</b>.
Select the application we created before, and select “<b>API Permissions</b>” or “<b>View API Permissions</b>”.
To assign a new permission to your application, click the “<b>Add a permission</b>” button.
From the “<b>Request API Permissions</b>” screen click “<b>Microsoft Graph</b>”.
For this connector the following permissions are used as <b>Application permissions</b>:
*	Read and Write data to an organization’s directory by using <b><i>Directory.Read.All</i></b>

Some high-privilege permissions can be set to admin-restricted and require an administrators consent to be granted.

To grant admin consent to our application press the “<b>Grant admin consent for TENANT</b>” button.

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the <b>Client ID</b>, go to the <b>Azure Portal > Azure Active Directory > App Registrations</b>.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a <b>Client Secret</b>.
*	From the Azure Portal, go to <b>Azure Active Directory > App Registrations</b>.
*	Select the application we have created before, and select "<b>Certificates and Secrets</b>". 
*	Under “Client Secrets” click on the “<b>New Client Secret</b>” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get is the <b>Tenant ID</b>. This can be found in the Azure Portal by going to <b>Azure Active Directory > Custom Domain Names</b>, and then finding the .onmicrosoft.com domain.


## Script variables
| Variable name | Description   | Example value |
| -| -  | - |
| AADtenantID | Id of the Azure tenant | 	12ab345c-0c41-4cde-9908-dabf3cad26b6   |
| AADAppId  | Id of the Azure app  |   12ab123c-fe99-4bdc-8d2e-87405fdb2379   |
| AADAppSecret   |  Secret of the Azure app  |   AB01C~DeFgHijkLMN.k-11AVdZSRzVnltkPqr   |
| licensesToCheck  | Hashtable containing all the licenses to check and their individueal thresholds  | See the provided script for an example |
| thresholdPercentageDefault  | Default threshold percentage value (percentage of licenses in use) to use, any license provided in 'licensesToCheck' without a threshold will use the default |  See the provided script for an example   |
| thresholdValueDefault  | Default threshold value (amount of licenses left) to use, any license provided in 'licensesToCheck' without a threshold will use the default |  See the provided script for an example   |
| csvPath  | path to the csv, which has to be the (latest) download of 'Product names and service plan identifiers for licensing.csv' from Microsoft (https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference. This is needed because the Graph API itself does not provide friendly names) |  C:\_Data\Scripting\PowerShell\AzureAD\Product names and service plan identifiers for licensing.csv   |

## Azure license check.ps1
This script only queries the subscription data from Azure and shows the reached threshold in PS only.

## Azure license check - send mail.ps1
This script only queries the subscription data from Azure and sends an e-mail containing a table showing the reached threshold.

## Azure license check - create topdesk incident.ps1
This script only queries the subscription data from Azure and creates a TOPdesk ticket containing a table showing the reached threshold.

## Azure license check - create topdesk incident - ids.ps1
This script only queries the subscription data from Azure and creates a TOPdesk change containing a table showing the reached threshold.
    > This differs from 'create topdesk incident.ps1' in that you can only provide the ids of the TOPdesk fields and not the friendly name.

## Azure license check - create topdesk change.ps1
This script only queries the subscription data from Azure and creates a TOPdesk change containing a table showing the reached threshold.

## Remarks
- The 4 examples which actually act on the threshold each have variables extra on top of the above specified, however they are not included in the Readme, since they are self explanatory.
    > If you need any help, feel free to contact Toosl4ever.

## Getting help
_If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/service-automation/679-helloid-sa-azuread-licenses-overview)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
