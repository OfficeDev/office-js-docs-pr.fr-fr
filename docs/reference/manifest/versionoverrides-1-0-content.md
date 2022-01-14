---
title: Élément VersionOverrides 1.0 dans le fichier manifeste pour un add-in de contenu
description: Documentation de référence de l’élément VersionOverrides (contenu) pour Office de manifeste des modules (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a9cd431f0e8fb4a7abe49103522e04900d9bcfd
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042173"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>Élément VersionOverrides 1.0 dans le fichier manifeste pour un add-in de contenu

Cet élément contient des informations pour les fonctionnalités qui ne sont pas pris en charge dans le manifeste de base.

> [!NOTE]
> Cet article suppose que vous connaissez la vue d’ensemble de l’élément [VersionOverrides,](versionoverrides.md)qui contient des informations importantes sur les attributs et les variantes de l’élément.

## <a name="child-elements"></a>Éléments enfants

Le tableau suivant s’applique uniquement à la version 1.0 des éléments **VersionOverrides** et uniquement aux modules de contenu.

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Non  | Actuellement insaisissable dans VersionOverrides 1.0 pour les add-ins de contenu. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Spécifie des détails sur l’inscription du add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0. |

## <a name="example"></a>Exemple

Voici un exemple simple. Pour obtenir des exemples plus complets, consultez les manifestes des exemples de Office exemples de [code de la version de l’exemple.](https://github.com/OfficeDev/PnP-OfficeAddins)

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
