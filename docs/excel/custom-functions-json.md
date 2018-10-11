---
ms.date: 09/27/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459237"
---
# <a name="custom-functions-metadata-preview"></a>Métadonnées des fonctions personnalisées (aperçu)

Lorsque vous définissez des[fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.

Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Métadonnées d’exemple

L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit les fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.

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
> Un fichier d’exemple JSON complet est disponible dans le référentiel GitHub [ OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>fonctions 

La propriété `functions` est un tableau d’objets de fonctions personnalisées.. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non  |  Description de la fonction que les utilisateurs voient dans Excel. Par exemple, **Convertit une valeur Celsius en Fahrenheit**. |
|  `helpUrl`  |  string  |   Non  |  URL qui fournit des informations sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Oui | ID unique de la fonction. Cet ID ne doit pas être modifié après sa définition. |
|  `name`  |  string  |  Oui  |  Nom de la fonction que l’utilisateur final voit dans Excel. Dans Excel, ce nom de fonction aura pour préfixe l’espace de noms des fonctions personnalisées qui est spécifié dans le fichier manifeste XML. |
|  `options`  |  object  |  Non  |  Permet de personnaliser en partie comment et quand Excel exécute la fonction. Voir l' [objet options](#options-object) pour plus d’informations. |
|  `parameters`  |  array  |  Oui  |  Tableau qui définit les paramètres d’entrée de la fonction. Consultez [Tableau de paramètres](#parameters-array) pour plus d’informations. |
|  `result`  |  objet  |  Oui  |  Objet qui définit le type d’informations renvoyées par la fonction. Voir l' [Objet de résultat](#result-object) pour plus d’informations. |

## <a name="options"></a>options

L’objet `options` vous permet de personnaliser en partie comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l'objet  `options`.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`. Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si `true`, la fonction peut déclencher le recalcul d'une cellule de manière répétée, même lorsqu’elle est appelée une seule fois. Cette option est utile pour les sources de données qui évoluent rapidement, telles que des actions. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). La fonction ne devrait pas avoir de déclaration `return`. Au lieu de cela, la valeur du résultat est transmise en tant que motif de la méthode de rappel  `caller.setResult`. Pour plus d’informations, voir [Diffusion en continu d’une fonction](custom-functions-overview.md#streaming-functions). |

## <a name="parameters"></a>parameters

La propriété `parameters` est un tableau de paramètres d'objets. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre.  |
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).  |
|  `name`  |  string  |  Oui  |  Le nom du paramètre. Ce nom est affiché dans intelliSense d'Excel.  |
|  `type`  |  string  |  Non  |  Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.  |

## <a name="result"></a>result

L'objet `results` définit le type d’informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l'objet `result` .

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Non  |  Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions). |
|  `type`  |  string  |  Oui  |  Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.  |

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)