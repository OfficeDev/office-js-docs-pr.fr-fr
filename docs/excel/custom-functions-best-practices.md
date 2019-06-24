---
ms.date: 06/18/2019
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques pour l’utilisation des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 7c836119a783f5cc7e1e7f4f52f1d21b86091bfe
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127932"
---
# <a name="custom-functions-best-practices"></a>Meilleures pratiques pour l’utilisation des fonctions personnalisées

Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Comme décrit dans l’article[vue d’ensemble de fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier de script (JavaScript ou machine à écrire) pour former une fonction complète. Si vous utilisez `yo office` les métadonnées JSON, vous pouvez les générer à partir des commentaires de code. Dans le cas contraire, vous devez générer le fichier de métadonnées JSON manuellement.

Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de `id` la fonction à l’implémentation JavaScript. Vérifiez qu’il existe une association, sinon la fonction ne sera pas appelée. L’exemple de code suivant montre comment effectuer l’Association à l' `CustomFunctions.associate()` aide de la méthode. L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Le code JSON suivant illustre les métadonnées JSON associées au code JavaScript de fonction personnalisée précédent.

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier. Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.

* Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante. Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.

* Dans le fichier JavaScript, spécifiez une association de fonctions `CustomFunctions.associate` personnalisées à l’aide de after each.

L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript. Les `id` valeurs `name` de la propriété et sont en majuscules, ce qui est recommandé lors de la description de vos fonctions personnalisées. Vous n’avez besoin d’ajouter ce JSON que si vous préparez votre propre fichier JSON manuellement et non à l’aide de la génération automatique. Pour plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="additional-considerations"></a>Considérations supplémentaires

Évitez d’accéder directement ou indirectement au modèle DOM (Document Object Model) (par exemple, à l’aide de jQuery) à partir de votre fonction personnalisée. Dans Excel sur Windows, où les fonctions personnalisées utilisent le [Runtime JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [effectuer des requêtes Web avec des fonctions personnalisées](custom-functions-web-reqs.md).

## <a name="see-also"></a>Voir aussi

* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
