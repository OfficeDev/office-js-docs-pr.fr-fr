---
title: Utilisation des substitutions étendues du manifeste
description: Découvrez comment configurer des fonctionnalités d’extensibilité avec des substitutions étendues du manifeste.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 09ced571f4b7d72a3479984582a8f58a0cb440bb2a3e62afe3f90329f2cd1be3
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080669"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Utilisation des substitutions étendues du manifeste

Certaines fonctionnalités d’extensibilité des Office Sont configurées avec des fichiers JSON hébergés sur votre serveur, et non avec le manifeste XML du module.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les manifestes de Office et leur rôle dans les add-ins. Veuillez lire Office manifeste XML des [add-ins,](add-in-manifests.md)si ce n’est pas le cas récemment.

Le tableau suivant spécifie les fonctionnalités d’extensibilité qui nécessitent une substitution étendue, ainsi que des liens vers la documentation de la fonctionnalité.

| Fonctionnalité | Instructions de développement |
| :----- | :----- |
| Raccourcis clavier | [Ajouter des raccourcis clavier personnalisés à vos Office de commande](../design/keyboard-shortcuts.md) |

Le schéma qui définit le format JSON est [un schéma de manifeste étendu.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!TIP]
> Cet article est quelque peu abstrait. Envisagez de lire l’un des articles du tableau pour clarifier les concepts.

## <a name="tell-office-where-to-find-the-json-file"></a>Indiquer Office où trouver le fichier JSON

Utilisez le manifeste pour indiquer Office où trouver le fichier JSON. Juste *en dessous* (pas à l’intérieur) de l’élément dans le manifeste, ajoutez un élément `<VersionOverrides>` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Définissez `Url` l’attribut sur l’URL complète d’un fichier JSON. Voici un exemple de l’élément le plus `<ExtendedOverrides>` simple possible.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

Voici un exemple de fichier JSON de remplacements étendu très simple. Il affecte le raccourci clavier Ctrl+Shift+A à une fonction (définie ailleurs) qui ouvre le volet Des tâches du module.

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

## <a name="localize-the-extended-overrides-file"></a>Localiser le fichier de remplacements étendu

Si votre add-in prend en charge plusieurs paramètres régionaux, vous pouvez utiliser l’attribut de l’élément pour pointer Office vers un `ResourceUrl` `<ExtendedOverrides>` fichier de ressources localisées. Voici un exemple.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Pour plus d’informations sur la création et l’utilisation du fichier de ressources, sur la façon de faire référence à ses ressources dans le fichier de remplacements étendu et pour les options supplémentaires non abordées ici, voir [Localize extended overrides](localization.md#localize-extended-overrides).
