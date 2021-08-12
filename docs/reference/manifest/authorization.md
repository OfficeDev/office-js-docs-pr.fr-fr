---
title: Élément Authorization dans le fichier manifeste
description: Spécifie une ressource externe à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: af40a47c4ae30b6d18d3457704487027ff18ac92da2a3ae23cf1afe5c1e9b46a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087710"
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
