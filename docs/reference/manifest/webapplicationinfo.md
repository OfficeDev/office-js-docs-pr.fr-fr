---
title: Élément WebApplicationInfo dans le fichier manifeste
description: Documentation de référence de l’élément VersionOverrides pour les fichiers manifeste des compléments Office (XML).
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1bbb3cc9b3db792b2d24ab2fd4003be6093fa837
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604478"
---
# <a name="webapplicationinfo-element"></a>Élément WebApplicationInfo

Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :

- En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.
- Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.

> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge en préversion pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](../requirement-sets/identity-api-requirement-sets.md). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.  

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **Id**    |  Oui   |  **ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.|
|  **MsaId**    |  Non   |  ID client de l’application Web de votre complément pour MSA, tel qu’inscrit dans msm.live.com.|
|  **Resource**  |  Oui   |  Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.|
|  [Scopes](scopes.md)                |  Oui  |  Spécifie les autorisations dont le complément a besoin pour une ressource, telle que Microsoft Graph.  |
|  [Autorisations](authorizations.md)  |  Non   | Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.|

## <a name="webapplicationinfo-example"></a>Exemple pour WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
