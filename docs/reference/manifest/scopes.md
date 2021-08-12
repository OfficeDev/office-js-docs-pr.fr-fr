---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le add-in a besoin pour se connecter à une ressource externe.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 05582ae05c13fae8e2272de3fe6111c5ff639f938a817fd0b50ad22e4234d033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087255"
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
