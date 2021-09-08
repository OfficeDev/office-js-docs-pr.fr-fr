---
title: Élément Authorization dans le fichier manifeste
description: Spécifie une ressource externe à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938187"
---
# <a name="authorization-element"></a>Élément Authorization

Spécifie les ressources externes à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.

**L’autorisation** est un élément enfant de [l’élément Authorizations](authorizations.md) dans le manifeste.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Resource**  |  Oui   |  Spécifie l’URL de la ressource externe.|
|  [Scopes](scopes.md)                |  Oui  |  Spécifie les autorisations dont le add-in a besoin pour la ressource.  |

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
