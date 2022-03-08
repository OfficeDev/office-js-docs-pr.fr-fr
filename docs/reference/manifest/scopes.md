---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le add-in a besoin pour se connecter à une ressource externe.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 883a1e318df7262bf8cdbd9d97b9d02d201066d8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340399"
---
# <a name="scopes-element"></a>Élément Scopes

Contient les autorisations dont le add-in a besoin pour une ressource externe, telle que Microsoft Graph. Lorsque Microsoft Graph la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

**Type de add-in :** Volet De tâches, Courrier, Contenu

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Contenu 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)

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
