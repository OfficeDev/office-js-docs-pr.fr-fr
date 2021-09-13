---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le add-in a besoin pour se connecter à une ressource externe.
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 346a143fdba35a153229b00052a463f726fd9056
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153608"
---
# <a name="scopes-element"></a>Élément Scopes

Contient les autorisations dont le add-in a besoin pour une ressource externe, telle que Microsoft Graph. Lorsque Microsoft Graph la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

**Scopes est** un élément enfant des éléments [WebApplicationInfo](webapplicationinfo.md) et [Authorization](authorization.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  Oui     |   Nom d’une autorisation ; par exemple, Files.Read.All ou profil. |

## <a name="example"></a>Exemple

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
