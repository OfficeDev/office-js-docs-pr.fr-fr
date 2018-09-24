---
ms.date: 09/20/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062143"
---
# <a name="custom-functions-metadata"></a>Métadonnées des fonctions personnalisées

Lorsque vous définissez des[fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour inscrire les fonctions personnalisées et de les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.

> [!NOTE]
> Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md#learn-the-basics).

## <a name="example-metadata"></a>Métadonnées d’exemple

L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "increment",
                    "description": "the number to be added each time",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "stream": true,
                "cancelable": true
            }
        },
        {
            "id": "SECONDHIGHEST",
            "name": "SECONDHIGHEST", 
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "range",
                    "description": "the input range",
                    "type": "number",
                    "dimensionality": "matrix"
                }
            ]
        }
    ]
}
```

> [!NOTE]
> Un fichier d’exemple JSON complet est disponible dans le [référentiel GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

La propriété `functions` est un tableau d’objets de fonctions personnalisées. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non  |  Une description de la fonction apparaissant dans l’interface utilisateur Excel. Par exemple, **Convertit une valeur Celsius en Fahrenheit**. |
|  `helpUrl`  |  string  |   Non  |  L’URL où vos utilisateurs peuvent obtenir de l’aide sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Oui | ID unique de la fonction. Cet ID ne doit pas être modifié après sa définition. |
|  `name`  |  string  |  Oui  |  Le nom de la fonction telle qu'elle apparaîtra (préfixée d'un espace de nom) dans l'interface utilisateur Excel lorsqu'un utilisateur sélectionne une fonction. Il n’a pas besoin d’être le même que le nom de la fonction telle que définie dans le JavaScript. |
|  `options`  |  object  |  Non  |  Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand. Voir [objet options](#options-object) pour plus de détails. |
|  `parameters`  |  array  |  Oui  |  Tableau qui définit les paramètres d’entrée de la fonction. Voir[tableau parameters](#parameters-array) pour plus de détails. |
|  `result`  |  object  |  Oui  |  Objet qui définit le type de l’information renvoyée par la fonction. Voir [objet result](#result-object) pour plus de détails. |

## <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand. Le tableau suivant répertorie les propriétés de l’objet `options`.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Non, la valeur par défaut est `false`.  |  Lorsqu’`true`Excel appelle le `onCanceled` gestionnaire au moment où l'utilisateur prend une action visant par exemple à annuler la fonction, le déclenchement manuel du recalcul ou la modification d’une cellule est référencée par cette fonction. Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`. Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Non, la valeur par défaut est `false`.  |  Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu'elle n'est invoquée qu'une seule fois. Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action. Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). La fonction ne devrait pas avoir de `return` déclaration. Au lieu de cela, la valeur du résultat est passée comme argument à la méthode de rappel `caller.setResult`. Pour plus d’informations, voir [Fonctions de flux](custom-functions-overview.md#streamed-functions). |

## <a name="parameters"></a>parameters

La propriété `parameters` est un tableau d’objets parameter. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Une description du paramètre.  |
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).  |
|  `name`  |  string  |  Oui  |  Nom du paramètre. Ce nom est affiché dans l’IntelliSense d’Excel.  |
|  `type`  |  string  |  Non  |  Le type de données du paramètre. Doit être **boolean**, **number** ou **string**.  |

## <a name="result"></a>result

L’objet `results` définit le type de l’information renvoyée par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions). |
|  `type`  |  string  |  Oui  |  Le type de données du paramètre. Doit être **boolean**, **number** ou **string**.  |

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)