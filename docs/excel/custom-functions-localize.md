---
ms.date: 04/29/2020
description: Localisez vos fonctions personnalisées Excel.
title: Localiser des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 427bff029c5e85caa216f628df450525ee187c17
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609295"
---
# <a name="localize-custom-functions"></a>Localiser des fonctions personnalisées

Vous pouvez localiser votre complément et vos noms de fonctions personnalisées. Pour ce faire, fournissez des noms de fonction localisés dans le fichier JSON des fonctions et des informations de paramètres régionaux dans le fichier manifeste XML.

>[!IMPORTANT]
> Les métadonnées générées automatiquement ne fonctionnent pas pour la localisation, c’est pourquoi vous devez mettre à jour le fichier JSON manuellement. Pour savoir comment procéder, consultez la rubrique [Metadata for Custom Functions in Excel](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Noms des fonctions de localisation

Pour localiser vos fonctions personnalisées, créez un nouveau fichier de métadonnées JSON pour chaque langue. Dans chaque fichier JSON de langue, créez `name` et des `description` Propriétés dans la langue cible. Le fichier par défaut pour l’anglais est nommé **functions. JSON**. Utilisez les paramètres régionaux dans le nom de fichier de tous les fichiers JSON supplémentaires, tels que les **fonctions-de-JSON** pour les identifier.

Le `name` et `description` s’affichent dans Excel et sont localisés. Toutefois, la `id` de chaque fonction n’est pas localisée. La `id` propriété indique comment Excel identifie votre fonction comme étant unique et ne doit pas être modifiée une fois qu’elle a été définie.

Le code JSON suivant montre comment définir une fonction avec la `id` propriété « Multiply ». La `name` `description` propriété et de la fonction est localisée pour l’allemand. Chaque paramètre `name` et `description` est également localisé pour l’allemand.

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

## <a name="localize-your-add-in"></a>Localiser votre complément

Après avoir créé un fichier JSON pour chaque langue, mettez à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre régional qui spécifie l’URL de chaque fichier de métadonnées JSON. Le code XML de manifeste suivant affiche les `en-us` paramètres régionaux par défaut avec une URL de fichier JSON de remplacement pour `de-de` (Allemagne). Le fichier **Functions-de. JSON** contient les noms et les ID des fonctions localisées en allemand.

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

Pour plus d’informations sur le processus de localisation d’un complément, reportez-vous à la rubrique [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Étapes suivantes
Découvrez [les conventions d’affectation de noms pour les fonctions personnalisées](custom-functions-naming.md) ou découvrir les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
