---
ms.date: 04/30/2019
description: Localisez vos fonctions personnalisées Excel.
title: Localiser des fonctions personnalisées (aperçu)
localization_priority: Normal
ms.openlocfilehash: 1c7fba297996c8cf050eb23b34823debf87b4e88
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527308"
---
# <a name="localize-custom-functions"></a>Localiser des fonctions personnalisées

Pour faire fonctionner vos fonctions personnalisées dans le monde entier, localisez-les dans différentes langues. Pour localiser des fonctions personnalisées, vous devez fournir des noms de fonctions localisés dans le fichier JSON des fonctions et fournir des informations de paramètres régionaux dans le fichier manifeste XML. Les métadonnées générées automatiquement ne fonctionnent pas pour la localisation, c’est pourquoi vous devez mettre à jour le fichier JSON manuellement.

## <a name="localize-function-names"></a>Noms des fonctions de localisation

Pour localiser vos fonctions personnalisées, créez un nouveau fichier de métadonnées JSON pour chaque langue. Dans chaque fichier JSON de langue, `name` créez `description` et des propriétés dans la langue cible. Le fichier par défaut pour l’anglais est nommé **functions. JSON**. Il est recommandé d’utiliser les paramètres régionaux dans le nom de fichier de chaque fichier JSON supplémentaire, comme les **fonctions-de-JSON** pour les identifier. 

Le `name` et `description` s’affichent dans Excel et sont localisés. Toutefois, la `id` de chaque fonction n’est pas localisée. La `id` propriété indique comment Excel identifie votre fonction comme étant unique et ne doit pas être modifiée une fois qu’elle a été définie.

Le code JSON suivant montre comment définir une fonction avec la `id` propriété «Multiply». La `name` propriété `description` et de la fonction est localisée pour l’allemand. Chaque paramètre `name` et `description` est également localisé pour l’allemand.

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

Après avoir créé un fichier JSON pour chaque langue, vous devez mettre à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre régional qui spécifie l’URL de chaque fichier de métadonnées JSON. Le code XML de manifeste suivant affiche `en-us` les paramètres régionaux par défaut avec une URL de fichier `de-de` JSON de remplacement pour (Allemagne). Le fichier **Functions-de. JSON** contient les noms et les ID des fonctions localisées en allemand.

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

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
