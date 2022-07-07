---
title: Utiliser des remplacements étendus du manifeste
description: Découvrez comment configurer des fonctionnalités d’extensibilité avec des remplacements étendus du manifeste.
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43e9820f54f2812130f7f86529c52b20b92811a0
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659950"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Utiliser des remplacements étendus du manifeste

Certaines fonctionnalités d’extensibilité des compléments Office sont configurées avec des fichiers JSON hébergés sur votre serveur, et non avec le manifeste XML du complément.

> [!NOTE]
> Cet article part du principe que vous êtes familiarisé avec les manifestes de complément Office et leur rôle dans les compléments. Veuillez lire le [manifeste XML des compléments Office](add-in-manifests.md), si ce n’est pas le cas récemment.

Le tableau suivant spécifie les fonctionnalités d’extensibilité qui nécessitent un remplacement étendu, ainsi que des liens vers la documentation de la fonctionnalité.

| Fonctionnalité | Instructions de développement |
| :----- | :----- |
| Raccourcis clavier | [Ajouter des raccourcis clavier personnalisés à vos compléments Office](../design/keyboard-shortcuts.md) |

Le schéma qui définit le format JSON est [un schéma de manifeste étendu](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!TIP]
> Cet article est un peu abstrait. Envisagez de lire l’un des articles du tableau pour clarifier les concepts.

## <a name="tell-office-where-to-find-the-json-file"></a>Indiquer à Office où trouver le fichier JSON

Utilisez le manifeste pour indiquer à Office où trouver le fichier JSON. Immédiatement *en dessous* (pas à l’intérieur) de l’élément **\<VersionOverrides\>** dans le manifeste, ajoutez un élément [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Définissez l’attribut `Url` sur l’URL complète d’un fichier JSON. Voici un exemple de l’élément le plus simple possible **\<ExtendedOverrides\>** .

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

Voici un exemple d’un fichier JSON de remplacements étendus très simple. Il affecte le raccourci clavier Ctrl+Maj+A à une fonction (définie ailleurs) qui ouvre le volet Office du complément.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a>Localiser le fichier de remplacements étendus

Si votre complément prend en charge plusieurs paramètres régionaux, vous pouvez utiliser l’attribut `ResourceUrl` de l’élément **\<ExtendedOverrides\>** pour pointer Office vers un fichier de ressources localisées. Voici un exemple.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Pour plus d’informations sur la création et l’utilisation du fichier de ressources, sur la façon de faire référence à ses ressources dans le fichier de remplacements étendus et sur les options supplémentaires qui ne sont pas abordées ici, consultez [Localiser les remplacements étendus](localization.md#localize-extended-overrides).
