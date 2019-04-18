---
ms.date: 03/29/2019
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: 3703699348e99fd076fe0e3affac88038e3aaf59
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914255"
---
# <a name="custom-functions-metadata-preview"></a>Métadonnées de fonctions personnalisées (aperçu)

Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément inclut un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux. Ce fichier est généré:

- par vous-même, dans un fichier JSON manuscrit
- à partir des commentaires JSDoc que vous entrez au début de votre fonction

Les fonctions personnalisées sont inscrites lorsque l'utilisateur exécute le complément pour la première fois et après qu'il est disponible pour le même utilisateur dans tous les classeurs.

Cet article décrit le format du fichier de métadonnées JSON, en supposant que vous l'écrivez manuellement. Pour plus d'informations sur la génération de fichiers JSON de commentaire JSDoc, voir [generate JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.

## <a name="example-metadata"></a>Exemple de métadonnées

L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.

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
> Un exemple de fichier JSON complet est disponible dans le dépôt GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).

## <a name="functions"></a>fonctions 

La propriété `functions` est un tableau d’objets de fonction personnalisés. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non  |  Description de la fonction que voient les utilisateurs finaux dans Excel. Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**. |
|  `helpUrl`  |  string  |   Non  |  URL fournissant des informations sur la fonction (elle est affichée dans un volet des tâches). Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Oui | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini. |
|  `name`  |  string  |  Oui  |  Nom de la fonction que voient les utilisateurs finaux dans Excel. Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML. |
|  `options`  |  object  |  Non  |  Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Reportez-vous aux [options](#options) pour plus d’informations. |
|  `parameters`  |  tableau  |  Oui  |  Tableau qui définit les paramètres d’entrée de la fonction. Reportez-vous aux [paramètres](#parameters) pour plus d’informations. |
|  `result`  |  objet  |  Oui  |  Objet qui définit le type d’informations renvoyées par la fonction. Reportez-vous au [résultat](#result) pour plus d’informations. |

## <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l’objet `options`.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si la valeur est `true`, Excel appelle le gestionnaire `onCanceled` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire (n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être attribué au membre `caller.onCanceled`. Pour plus d’informations, voir [Annuler une fonction](custom-functions-web-reqs.md#canceling-a-function). |
|  `requiresAddress`  | boolean | Non <br/><br/>La valeur par défaut est `false`. | <br /><br /> Si la valeur est true, votre fonction personnalisée peut accéder à l'adresse de la cellule qui a appelé votre fonction personnalisée. Pour obtenir l'adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée. Pour plus d’informations, voir[Déterminer quelle cellule a appelé votre fonction personnalisée](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function). Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress. Lorsque vous utilisez cette option, le paramètre «invocationContext» doit être le dernier paramètre passé dans options. |
|  `stream`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique. Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire (n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`). La fonction ne doit pas utiliser d’instruction `return`. Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `caller.setResult`. Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#streaming-functions). |
|  `volatile`  | boolean | Non <br/><br/>La valeur par défaut est `false`. | <br /><br /> Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées. Une fonction ne peut pas être à la fois diffusée en continu et volatile. Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée. |

## <a name="parameters"></a>paramètres

La propriété `parameters` est un tableau d’objets paramètre. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre. S’affiche dans intelliSense d’Excel.  |
|  `dimensionality`  |  string  |  Non  |  Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).  |
|  `name`  |  string  |  Oui  |  Le nom du paramètre. Ce nom s’affiche dans intelliSense d’Excel.  |
|  `type`  |  string  |  Non  |  Type de données du paramètre. Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents. Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**. |
|  `optional`  | boolean | Non | Si la valeur est `true`, le paramètre est facultatif. |

>[!NOTE]
> Si la propriété `type` d’un paramètre facultatif n’est pas spécifiée ou est définie sur `any`, vous remarquerez peut-être des problèmes tels que des erreurs de linting dans votre IDE et des paramètres facultatifs non affichés lorsque la fonction est saisie dans une cellule Excel. Ces problèmes seront résolus en décembre 2018.

## <a name="result"></a>résultat

L’objet `result` définit le type des informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Non  |  Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel). |
|  `type`  |  string  |  Oui  |  Type de données du paramètre. Doit être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents. |

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
