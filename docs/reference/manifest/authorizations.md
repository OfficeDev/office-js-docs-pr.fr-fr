---
title: Authorization, élément dans le fichier manifeste
description: Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 7ae0b9d0ec32a20846142a9fc89c48fe9cdf8053
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720658"
---
# <a name="authorizations-element"></a>Authorizations, élément

Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.

**Authorizations** est un élément enfant de l’élément [WebApplicationInfo](webapplicationinfo.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Autorisation](authorization.md)                |  Oui     |   Identifie une ressource externe dont l’application Web du complément a besoin d’autorisation, ainsi que les étendues (autorisations) dont elle a besoin. |

## <a name="example"></a>Exemple

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
