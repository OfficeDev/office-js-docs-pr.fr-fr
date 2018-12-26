---
title: Élément Scopes dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432640"
---
# <a name="scopes-element"></a>Élément Scopes

Contient des autorisations Microsoft Graph requises par le complément. L’Office Store se sert de l’élément Scope pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

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
