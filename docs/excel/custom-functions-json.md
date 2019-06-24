---
ms.date: 06/20/2019
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: f97a339972a8ac134bd30c87b86c4701cb4b5fc4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127869"
---
# <a name="custom-functions-metadata"></a>Métadonnées des fonctions personnalisées

Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément inclut un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ce fichier est généré:

- Par vous-même, dans un fichier JSON manuscrit
- À partir des commentaires JSDoc que vous entrez au début de votre fonction

Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après qu’il est disponible pour le même utilisateur dans tous les classeurs.

Cet article décrit le format du fichier de métadonnées JSON, en supposant que vous l’écrivez manuellement. Pour plus d’informations sur la génération de fichiers JSON de commentaire JSDoc, voir [generate JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

Les paramètres serveur sur le serveur qui héberge le fichier JSON doivent avoir [cors](https://developer.mozilla.org/docs/Web/HTTP/CORS) activé afin que les fonctions personnalisées fonctionnent correctement dans Excel sur le Web.

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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> Un exemple de fichier JSON complet est disponible dans l’historique de validation du référentiel [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) github. Lorsque le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit est uniquement disponible dans les versions précédentes du projet.

## <a name="functions"></a>fonctions 

La propriété `functions` est un tableau d’objets de fonction personnalisés. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non  |  Description de la fonction que voient les utilisateurs finaux dans Excel. Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**. |
|  `helpUrl`  |  string  |   Non  |  URL fournissant des informations sur la fonction (elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`. |
| `id`     | string | Oui | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini. |
|  `name`  |  string  |  Oui  |  Nom de la fonction que voient les utilisateurs finaux dans Excel. Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML. |
|  `options`  |  object  |  Non  |  Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Reportez-vous aux [options](#options) pour plus d’informations. |
|  `parameters`  |  tableau  |  Oui  |  Tableau qui définit les paramètres d’entrée de la fonction. Reportez-vous aux [paramètres](#parameters) pour plus d’informations. |
|  `result`  |  objet  |  Oui  |  Objet qui définit le type d’informations renvoyées par la fonction. Reportez-vous au [résultat](#result) pour plus d’informations. |

## <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l’objet `options`.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction. Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui renvoient un seul résultat et doivent gérer l’annulation d’une demande de données. Une fonction ne peut pas être à la fois en continu et annulable. Pour plus d’informations, reportez-vous à la remarque à la fin de la [création d’une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function). |
|  `requiresAddress`  | boolean | Non <br/><br/>La valeur par défaut est `false`. | <br /><br /> Si la valeur est true, votre fonction personnalisée peut accéder à l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée. Pour plus d’informations, voir[Déterminer quelle cellule a appelé votre fonction personnalisée](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function). Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress. Lorsque vous utilisez cette option, le paramètre «invocation» doit être le dernier paramètre passé dans options. |
|  `stream`  |  boolean  |  Non<br/><br/>La valeur par défaut est `false`.  |  Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique. Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières. La fonction ne doit pas utiliser d’instruction `return`. Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`. Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#make-a-streaming-function). |
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

## <a name="result"></a>résultat

L’objet `result` définit le type des informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Non  |  Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel). |

## <a name="next-steps"></a>Étapes suivantes
Découvrez les [meilleures pratiques de dénomination de votre fonction](custom-functions-naming.md) ou Découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.

## <a name="see-also"></a>Voir aussi

* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)