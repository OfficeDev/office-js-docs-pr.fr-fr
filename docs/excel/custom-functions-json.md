---
ms.date: 01/14/2020
description: Définissez des métadonnées JSON pour les fonctions personnalisées dans Excel et associez vos ID de fonction et propriétés de nom.
title: Métadonnées pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 679087336fc7aea741c98d0104514ab96068ffbf
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719461"
---
# <a name="custom-functions-metadata"></a>Métadonnées des fonctions personnalisées

Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md) , un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier script (JavaScript ou machine à écriture) pour enregistrer une fonction, le rendant ainsi disponible. Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après qu’il est disponible pour le même utilisateur dans tous les classeurs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Il est recommandé d’utiliser la génération automatique JSON dans la mesure du possible `yo office` , à l’aide des fichiers de l’échafaudage, de la même manière que le processus illustré dans le [didacticiel de fonction personnalisée Excel](../tutorials/excel-tutorial-create-custom-functions.md) , car ce processus est plus facile et moins sujet aux erreurs de l’utilisateur. Pour plus d’informations sur le processus de génération de fichiers JSON de commentaire JSDoc, voir [génération de métadonnées JSON pour les fonctions personnalisées](custom-functions-json-autogeneration.md).

Toutefois, vous pouvez créer un projet de fonctions personnalisées de a à z ; pour ce faire, vous devez :

- Écrire votre fichier JSON manuellement
- Vérifier que votre fichier manifeste est connecté à votre fichier JSON créé manuellement
- Associez les fonctions `id` et `name` les propriétés dans le fichier de script pour enregistrer vos fonctions

Cet article vous explique comment effectuer ces trois étapes.

L’image suivante explique les différences entre l' `yo office` utilisation de fichiers de structure et l’écriture de JSON à partir de zéro.
![Image des différences entre l’utilisation de yo Office et l’écriture de votre propre JSON](../images/custom-functions-json.png)

> [!NOTE]
> Contrairement aux fichiers de `yo office` l’échafaudage, vous devez connecter votre manifeste au fichier JSON que vous créez, via la `<Resources>` section de votre fichier manifeste XML. Notez que [cors](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour les paramètres serveur sur le serveur qui héberge le fichier JSON afin que les fonctions personnalisées fonctionnent correctement dans Excel sur le Web.

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Création de métadonnées et connexion au manifeste

Vous devez créer un fichier JSON dans votre projet et fournir toutes les informations sur les fonctions qu’il contient, telles que les paramètres de la fonction. Consultez l' [exemple de métadonnées suivant](#json-metadata-example) et [la référence de métadonnées](#metadata-reference) pour obtenir la liste complète des propriétés de fonction.

Vous devez également vous assurer que votre fichier manifeste XML fait référence à votre fichier JSON `<Resources>` dans la section, comme dans l’exemple suivant.

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a>Exemple de métadonnées JSON

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
      "description": "Count up from zero",
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
      "description": "Get the second highest number from a range",
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

## <a name="metadata-reference"></a>Référence de métadonnées

### <a name="functions"></a>fonctions

La propriété `functions` est un tableau d’objets de fonction personnalisés. Le tableau suivant répertorie les propriétés de chaque objet.

| Propriété      | Type de données | Requis | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Non       | Description de la fonction que voient les utilisateurs finaux dans Excel. Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.                                                            |
| `helpUrl`     | string    | Non       | URL fournissant des informations sur la fonction (elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Oui      | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.                                            |
| `name`        | string    | Oui      | Nom de la fonction que voient les utilisateurs finaux dans Excel. Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML. |
| `options`     | object    | Non       | Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Reportez-vous aux [options](#options) pour plus d’informations.                                                          |
| `parameters`  | tableau     | Oui      | Tableau qui définit les paramètres d’entrée de la fonction. Pour plus d’informations, consultez la rubrique [paramètres](#parameters) .                                                                             |
| `result`      | objet    | Oui      | Objet qui définit le type d’informations renvoyées par la fonction. Reportez-vous au [résultat](#result) pour plus d’informations.                                                                 |

### <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l’objet `options`.

| Propriété          | Type de données | Requis                               | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | boolean   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction. Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui renvoient un seul résultat et doivent gérer l’annulation d’une demande de données. Une fonction ne peut pas être à la fois en continu et annulable. Pour plus d’informations, reportez-vous à la remarque à la fin de la [création d’une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function). |
| `requiresAddress` | boolean   | Non <br/><br/>La valeur par défaut est `false`. | Si `true`votre fonction personnalisée peut accéder à l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée. Pour plus d’informations, consultez la rubrique relative au [paramètre context de la cellule Addressing](../excel/custom-functions-parameter-options.md#addressing-cells-context-parameter). Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress. Lorsque vous utilisez cette option, le paramètre « invocation » doit être le dernier paramètre passé dans options.                                              |
| `stream`          | boolean   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique. Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières. La fonction ne doit pas utiliser d’instruction `return`. Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`. Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#make-a-streaming-function).                                                                                                                                                                |
| `volatile`        | boolean   | Non <br/><br/>La valeur par défaut est `false`. | <br /><br /> Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées. Une fonction ne peut pas être à la fois diffusée en continu et volatile. Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a>paramètres

La propriété `parameters` est un tableau d’objets paramètre. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre. S’affiche dans intelliSense d’Excel.  |
|  `dimensionality`  |  string  |  Non  |  Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).  |
|  `name`  |  string  |  Oui  |  Le nom du paramètre. Ce nom s’affiche dans intelliSense d’Excel.  |
|  `type`  |  string  |  Non  |  Type de données du paramètre. Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents. Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**. |
|  `optional`  | boolean | Non | Si la valeur est `true`, le paramètre est facultatif. |
|`repeating`| boolean | Non | Si `true`, les paramètres sont renseignés à partir d’un tableau spécifié. Notez que les fonctions de tous les paramètres répétitifs sont considérées comme des paramètres facultatifs par définition.  |

### <a name="result"></a>résultat

L’objet `result` définit le type des informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

| Propriété         | Type de données | Requis | Description                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Non       | Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel). |

## <a name="associating-function-names-with-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de `id` la fonction à l’implémentation JavaScript. Assurez-vous qu’il existe une association, sinon la fonction n’est pas inscrite et n’est pas utilisable dans Excel. L’exemple de code suivant montre comment effectuer l’Association à l' `CustomFunctions.associate()` aide de la méthode. L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.

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
    }
  ]
}
```

N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.

- Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.

- Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier. Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.

- Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante. Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.

- Dans le fichier JavaScript, spécifiez une association de fonctions `CustomFunctions.associate` personnalisées à l’aide de after each.

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

## <a name="next-steps"></a>Étapes suivantes

Découvrez les [meilleures pratiques de dénomination de votre fonction](custom-functions-naming.md) ou Découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.

## <a name="see-also"></a>Voir aussi

- [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
