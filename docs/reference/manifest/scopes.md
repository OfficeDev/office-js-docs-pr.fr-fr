---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le complément a besoin pour se connecter à une ressource externe.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612240"
---
# <a name="scopes-element"></a>Élément Scopes

Contient les autorisations dont le complément a besoin pour une ressource externe, telle que Microsoft Graph. Lorsque Microsoft Graph est la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

**Scopes** est un élément enfant des éléments [WebApplicationInfo](webapplicationinfo.md) et [authorization](authorization.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  Oui     |   Nom d’une autorisation ; par exemple, files. Read. All ou Profile. |

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
