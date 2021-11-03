---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le add-in a besoin pour se connecter à une ressource externe.
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 16e8a19a7aa73efa6aac00c915fde8d2b8647bad
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681534"
---
# <a name="scopes-element"></a>Élément Scopes

Contient les autorisations dont le add-in a besoin pour une ressource externe, telle que Microsoft Graph. Lorsque Microsoft Graph la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

**Scopes est** un élément enfant de [l’élément WebApplicationInfo](webapplicationinfo.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
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
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
