---
ms.date: 09/27/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: 025be277a5e436a1ce2885815e9b8cbf9b206799
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348134"
---
# <a name="custom-functions-metadata-preview"></a>Métadonnées des fonctions personnalisées (aperçu)

Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.

Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Métadonnées d’exemple

L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
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
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
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
|  `description`  |  string  |  Non  |  Description de la fonction que les utilisateurs voient dans Excel. Par exemple, **Convertit une valeur en Celsius en Fahrenheit**. |
|  `helpUrl`  |  string  |   Non  |  URL qui fournit des informations sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Oui | ID unique de la fonction. Cet ID ne doit pas être modifié après sa définition. |
|  `name`  |  string  |  Oui  |  Nom de la fonction que les utilisateurs voient dans Excel. Dans Excel, ce nom de fonction sera préfixé par l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML. |
|  `options`  |  object  |  Non  |  Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand. Voir [objet options](#options-object) pour plus de détails. |
|  `parameters`  |  array  |  Oui  |  Tableau qui définit les paramètres d’entrée de la fonction. Voir[tableau parameters](#parameters-array) pour plus de détails. |
|  `result`  |  object  |  Oui  |  Objet qui définit le type de l’information renvoyée par la fonction. Voir [objet result](#result-object) pour plus de détails. |

## <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand. Le tableau suivant répertorie les propriétés de l’objet `options`.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`. Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu’elle n’est invoquée qu’une seule fois. Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action. Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). La fonction ne devrait pas avoir d’instruction `return`. Au lieu de cela, la valeur du résultat est passée comme argument à la méthode de rappel `caller.setResult`. Pour plus d’informations, voir [Fonctions de flux](custom-functions-overview.md#streamed-functions). |

## <a name="parameters"></a>parameters

La propriété `parameters` est un tableau d’objets parameter. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre.  |
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).  |
|  `name`  |  string  |  Oui  |  Nom du paramètre. Ce nom est affiché dans l’IntelliSense d’Excel.  |
|  `type`  |  string  |  Non  |  Type de données du paramètre. Doit être **boolean**, **number** ou **string**.  |

## <a name="result"></a>result

L’objet `results` définit le type de l’information renvoyée par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions). |
|  `type`  |  string  |  Oui  |  Type de données du paramètre. Doit être **boolean**, **number** ou **string**.  |

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)