---
ms.date: 11/06/2020
description: Localisez vos Excel personnalisées.
title: Localiser les fonctions personnalisées
ms.localizationpriority: medium
ms.openlocfilehash: 7219c838cfd5a6c827b74b5d04442280be7ebac7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744506"
---
# <a name="localize-custom-functions"></a>Localiser les fonctions personnalisées

Vous pouvez localiser à la fois vos noms de fonctions personnalisées et de votre add-in. Pour ce faire, fournissez des noms de fonctions localisées dans le fichier JSON des fonctions et des informations de paramètres régionaux dans le fichier manifeste XML.

>[!IMPORTANT]
> Les métadonnées auto-genrées ne fonctionnent pas pour la localisation. Vous devez donc mettre à jour le fichier JSON manuellement. Pour savoir comment faire, voir [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Localiser les noms des fonctions

Pour localiser vos fonctions personnalisées, créez un fichier de métadonnées JSON pour chaque langue. Dans chaque fichier JSON de langue, créez et créez `name` `description` les propriétés dans la langue cible. Le fichier par défaut pour l’anglais est **nommé functions.json**. Utilisez les paramètres régionaux dans le nom de fichier pour chaque fichier JSON supplémentaire, tel que **functions-de.json** , pour les identifier.

Les `name` et apparaissent `description` dans Excel et sont localisées. Toutefois, la `id` fonction de chaque fonction n’est pas localisée. La `id` propriété est comment Excel identifie votre fonction comme unique et ne doit pas être modifiée une fois définie.

Le JSON suivant montre comment définir une fonction avec la `id` propriété « MULTIPLY ». La `name` propriété `description` et la propriété de la fonction sont localisées pour l’allemand. Chaque paramètre `name` est `description` également localisée pour l’allemand.

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

Après avoir créé un fichier JSON pour chaque langue, mettez à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre local qui spécifie l’URL de chaque fichier de métadonnées JSON. Le manifeste XML suivant affiche les paramètres `en-us` régionaux par défaut avec une URL de fichier JSON de substitution pour `de-de` (Allemagne). Le **fichier functions-de.json** contient les ID et les noms de fonctions allemands localisées.

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

Pour plus d’informations sur le processus de localisation d’un Office, consultez La localisation des [modules complémentaires](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Prochaines étapes
Découvrez les [conventions d’attribution de noms pour les fonctions personnalisées](custom-functions-naming.md) ou découvrez les meilleures [pratiques de gestion des erreurs](custom-functions-errors.md).

## <a name="see-also"></a>Voir aussi

* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
