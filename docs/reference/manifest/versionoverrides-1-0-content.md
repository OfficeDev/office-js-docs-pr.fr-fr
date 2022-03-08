---
title: Élément VersionOverrides 1.0 dans le fichier manifeste pour un add-in de contenu
description: Documentation de référence de l’élément VersionOverrides (contenu) Office fichiers manifeste (XML) des add-ins.
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ef083ef5df322c230292625576e36db8923d00c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341050"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>Élément VersionOverrides 1.0 dans le fichier manifeste pour un add-in de contenu

Cet élément contient des informations sur les fonctionnalités qui ne sont pas pris en charge dans le manifeste de base.

> [!NOTE]
> Cet article suppose que vous connaissez la vue d’ensemble de l’élément [VersionOverrides](versionoverrides.md), qui contient des informations importantes sur les attributs et les variantes de l’élément.

## <a name="child-elements"></a>Éléments enfants

Le tableau suivant s’applique uniquement à la version 1.0 des éléments **VersionOverrides** et uniquement aux modules de contenu.

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Non  | Actuellement insaisissable dans VersionOverrides 1.0 pour les add-ins de contenu. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Spécifie des détails sur l’inscription du add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0. |

## <a name="example"></a>Exemple

Voici un exemple simple. Pour obtenir des exemples plus complexes, consultez les manifestes des exemples de Office des [exemples de code de modules.](https://github.com/OfficeDev/PnP-OfficeAddins)

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
