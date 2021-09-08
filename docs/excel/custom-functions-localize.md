---
ms.date: 11/06/2020
description: Localisez vos Excel personnalisées.
title: Localiser les fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: b393cbb76e4993eb77df8ddbe60247c8af74c580
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936935"
---
# <a name="localize-custom-functions"></a>Localiser les fonctions personnalisées

Vous pouvez localiser à la fois vos noms de fonctions personnalisées et de votre add-in. Pour ce faire, fournissez des noms de fonctions localisées dans le fichier JSON des fonctions et des informations de paramètres régionaux dans le fichier manifeste XML.

>[!IMPORTANT]
> Les métadonnées auto-genrées ne fonctionnent pas pour la localisation. Vous devez donc mettre à jour le fichier JSON manuellement. Pour savoir comment faire, voir Créer manuellement des [métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Localiser les noms des fonctions

Pour localiser vos fonctions personnalisées, créez un fichier de métadonnées JSON pour chaque langue. Dans chaque fichier JSON de langue, créez `name` et créez `description` les propriétés dans la langue cible. Le fichier par défaut pour l’anglais est **nomméfunctions.jssur**. Utilisez les paramètres régionaux dans le nom de fichier pour chaque fichier JSON supplémentaire, par exemple **functions-de.jspour** les identifier.

Les `name` et `description` apparaissent dans Excel et sont localisées. Toutefois, `id` la fonction de chaque fonction n’est pas localisée. La propriété est comment Excel votre fonction comme unique et ne doit pas être modifiée une fois `id` qu’elle est définie.

Le JSON suivant montre comment définir une fonction avec la `id` propriété « MULTIPLY ». La `name` propriété et la propriété de la fonction sont `description` localisées pour l’allemand. Chaque paramètre `name` est également localisée pour `description` l’allemand.

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

Comparez le JSON précédent avec le JSON suivant pour l’anglais.

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a>Localiser votre add-in

Après avoir créé un fichier JSON pour chaque langue, mettez à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre local qui spécifie l’URL de chaque fichier de métadonnées JSON. Le manifeste XML suivant affiche les paramètres régionaux par défaut avec une URL de fichier JSON de substitution `en-us` `de-de` pour (Allemagne). Le **functions-de.jssur** le fichier contient les ID et les noms de fonctions allemands localisées.

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

Pour plus d’informations sur le processus de localisation d’un Office, consultez La localisation [des modules complémentaires.](../develop/localization.md#control-localization-from-the-manifest)

## <a name="next-steps"></a>Prochaines étapes
Découvrez les [conventions d’attribution de noms pour les fonctions personnalisées](custom-functions-naming.md) ou découvrez les meilleures [pratiques de gestion des erreurs.](custom-functions-errors.md)

## <a name="see-also"></a>Voir aussi

* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
