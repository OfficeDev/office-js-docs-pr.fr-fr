---
title: Élément WebApplicationInfo dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1328dc40e98c321c9c4b7d3d692da8c8bdd29492
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389196"
---
# <a name="webapplicationinfo-element"></a>Élément WebApplicationInfo

Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :

- En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.
- Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.

> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge en préversion pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.  

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Id**    |  Oui   |  **ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.|
|  **Resource**  |  Oui   |  Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.|
|  [Scopes](scopes.md)                |  Non  |  Spécifie les autorisations dont le complément a besoin pour Microsoft Graph.  |

> [!NOTE] 
> À l’heure actuelle, il est nécessaire que les ressources de votre complément correspondent à son hôte. Office ne demandera pas un jeton pour un complément à moins de pouvoir prouver qu’il en est le propriétaire ; à l’heure actuelle, ceci s’effectue en hébergeant le complément sous le nom de domaine complet de la ressource.

## <a name="webapplicationinfo-example"></a>Exemple pour WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
