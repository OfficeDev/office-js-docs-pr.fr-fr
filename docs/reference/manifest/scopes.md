---
title: Élément Scopes dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771785"
---
# <a name="scopes-element"></a>Élément Scopes

Contient des autorisations Microsoft Graph requises par le complément. AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Type  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Nom d’une autorisation Microsoft Graph ; par exemple, Files.Read.All. |

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
