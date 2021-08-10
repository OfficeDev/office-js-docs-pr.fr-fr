---
title: Élément Authorizations dans le fichier manifeste
description: Spécifie les ressources externes à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 068e6753e2e8e947e5e6e3c0885e7cd006165660862a37346eea114abb81a9b8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092500"
---
# <a name="authorizations-element"></a>Élément Authorizations

Spécifie les ressources externes à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.

**Authorizations est** un élément enfant de [l’élément WebApplicationInfo](webapplicationinfo.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Autorisation](authorization.md)                |  Oui     |   Identifie une ressource externe à qui l’application web du add-in a besoin d’une autorisation, ainsi que les étendues (autorisations) dont elle a besoin. |

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
