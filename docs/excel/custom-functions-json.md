---
ms.date: 12/22/2020
description: Définissez les métadonnées JSON pour les fonctions personnalisées Excel et associez votre ID de fonction et vos propriétés de nom.
title: Créer manuellement des métadonnées JSON pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: c03238d46e8d861307ba0db3d03dafea81aeca51
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349629"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Créer manuellement des métadonnées JSON pour les fonctions personnalisées

Comme décrit dans l’article de vue d’ensemble des fonctions [personnalisées,](custom-functions-overview.md) un projet de fonctions personnalisées doit inclure à la fois un fichier de métadonnées JSON et un fichier de script (JavaScript ou TypeScript) pour inscrire une fonction, ce qui le rend disponible pour utilisation. Les fonctions personnalisées sont enregistrées lorsque l’utilisateur exécute le add-in pour la première fois et après cela sont disponibles pour le même utilisateur dans tous les workbooks.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Nous vous recommandons d’utiliser la génération automatique JSON lorsque cela est possible au lieu de créer votre propre fichier JSON. La génération automatique est moins sujette aux erreurs de l’utilisateur et les fichiers `yo office` échafaudés l’incluent déjà. Pour plus d’informations sur les balises JSDoc et le processus de génération automatique JSON, voir métadonnées JSON de génération automatique [pour les fonctions personnalisées.](custom-functions-json-autogeneration.md)

Toutefois, vous pouvez créer un projet de fonctions personnalisées à partir de zéro. Ce processus nécessite que vous :

- Écrivez votre fichier JSON.
- Vérifiez que votre fichier manifeste est connecté à votre fichier JSON.
- Associez les propriétés et les fonctions de vos fonctions dans `id` le fichier de script afin `name` d’inscrire vos fonctions.

L’image suivante explique les différences entre l’utilisation de fichiers de la `yo office` échafaudage et l’écriture de JSON à partir de zéro.

![Image des différences entre l’utilisation de Yo Office et l’écriture de votre propre JSON.](../images/custom-functions-json.png)

> [!NOTE]
> N’oubliez pas de connecter votre manifeste au fichier JSON que vous créez, via la section de votre fichier manifeste XML si vous `<Resources>` n’utilisez pas le `yo office` générateur.

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Créer des métadonnées et se connecter au manifeste

Créez un fichier JSON dans votre projet et fournissez tous les détails sur vos fonctions, telles que les paramètres de la fonction. Consultez [l’exemple de métadonnées suivant](#json-metadata-example) [et la référence des métadonnées](#metadata-reference) pour obtenir la liste complète des propriétés de la fonction.

Assurez-vous que votre fichier manifeste XML fait référence à votre fichier JSON dans la `<Resources>` section, comme dans l’exemple suivant.

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
> Un exemple complet de fichier JSON est disponible dans [officeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub’historique de validation du référentiel. Comme le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit n’est disponible que dans les versions précédentes du projet.

## <a name="metadata-reference"></a>Référence des métadonnées

### <a name="functions"></a>fonctions

La propriété `functions` est un tableau d’objets de fonction personnalisés. Le tableau suivant répertorie les propriétés de chaque objet.

| Propriété      | Type de données | Requis | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Non       | Description de la fonction que voient les utilisateurs finaux dans Excel. Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.                                                            |
| `helpUrl`     | string    | Non       | URL fournissant des informations sur la fonction (elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Oui      | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.                                            |
| `name`        | string    | Oui      | Nom de la fonction que voient les utilisateurs finaux dans Excel. Dans Excel, ce nom de fonction est précédé de l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML. |
| `options`     | object    | Non       | Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Reportez-vous aux [options](#options) pour plus d’informations.                                                          |
| `parameters`  | tableau     | Oui      | Tableau qui définit les paramètres d’entrée de la fonction. Pour [plus d’informations,](#parameters) voir paramètres.                                                                             |
| `result`      | objet    | Oui      | Objet qui définit le type d’informations renvoyées par la fonction. Reportez-vous au [résultat](#result) pour plus d’informations.                                                                 |

### <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l’objet `options`.

| Propriété          | Type de données | Requis                               | Description |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction. Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui retournent un résultat unique et qui doivent gérer l’annulation d’une demande de données. Une fonction ne peut pas utiliser les `stream` propriétés et les `cancelable` propriétés. |
| `requiresAddress` | boolean   | Non <br/><br/>La valeur par défaut est `false`. | Si `true` , votre fonction personnalisée peut accéder à l’adresse de la cellule qui l’a appelé. La `address` propriété du paramètre [d’appel](custom-functions-parameter-options.md#invocation-parameter) contient l’adresse de la cellule qui a appelé votre fonction personnalisée. Une fonction ne peut pas utiliser les `stream` propriétés et les `requiresAddress` propriétés. |
| `requiresParameterAddresses` | boolean   | Non <br/><br/>La valeur par défaut est `false`. | Si `true` , votre fonction personnalisée peut accéder aux adresses des paramètres d’entrée de la fonction. Cette propriété doit être utilisée en association avec la propriété de l’objet de résultat et doit `dimensionality` être définie sur [](#result) `dimensionality` `matrix` . Pour [plus d’informations, voir Détecter l’adresse d’un](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) paramètre. |
| `stream`          | boolean   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique. Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières. La fonction ne doit pas utiliser d’instruction `return`. Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`. Pour plus d’informations, [voir Faire une fonction de diffusion en continu.](custom-functions-web-reqs.md#make-a-streaming-function) |
| `volatile`        | boolean   | Non <br/><br/>La valeur par défaut est `false`. | Si , la fonction recalcule à chaque Excel recalcul, et non uniquement lorsque les valeurs dépendantes de la `true` formule ont changé. Une fonction ne peut pas utiliser les `stream` propriétés et les `volatile` propriétés. Si les `stream` `volatile` propriétés et les propriétés sont définies sur , la propriété `true` volatile est ignorée. |

### <a name="parameters"></a>paramètres

La propriété `parameters` est un tableau d’objets paramètre. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Requis  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre. Elle s’affiche dans Excel’IntelliSense.  |
|  `dimensionality`  |  string  |  Non  |  Doit être `scalar` (une valeur autre qu’un tableau) ou (un tableau `matrix` à 2 dimensions).  |
|  `name`  |  string  |  Oui  |  Le nom du paramètre. Ce nom s’affiche dans Excel’IntelliSense.  |
|  `type`  |  string  |  Non  |  Type de données du paramètre. Peut être , ou , qui vous permet d’utiliser l’un des trois `boolean` `number` types `string` `any` précédents. Si cette propriété n’est pas spécifiée, le type de données est par défaut `any` . |
|  `optional`  | boolean | Non | Si la valeur est `true`, le paramètre est facultatif. |
|`repeating`| boolean | Non | Si `true` , les paramètres sont remplis à partir d’un tableau spécifié. Notez que, par définition, tous les paramètres exexionnels sont considérés comme des paramètres facultatifs.  |

### <a name="result"></a>résultat

L’objet `result` définit le type des informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

| Propriété         | Type de données | Requis | Description                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Non       | Doit être `scalar` (une valeur autre qu’un tableau) ou (un tableau `matrix` à 2 dimensions). |
| `type` | string    | Non       | Type de données du résultat. Peut être , ou (ce qui vous permet d’utiliser l’un des trois `boolean` `number` types `string` `any` précédents). Si cette propriété n’est pas spécifiée, le type de données est par défaut `any` . |

## <a name="associating-function-names-with-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de la fonction `id` à l’implémentation JavaScript. Assurez-vous qu’il existe une association, sinon la fonction ne sera pas enregistrée et ne peut pas être Excel. L’exemple de code suivant montre comment faire en sorte que l’association utilise la `CustomFunctions.associate()` méthode. L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est **AJOUTER**.

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

Le code JSON suivant présente les métadonnées JSON associées au code JavaScript de la fonction personnalisée précédente.

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

- Dans le fichier JavaScript, spécifiez une association de fonction personnalisée à l’aide `CustomFunctions.associate` d’après chaque fonction.

L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans l’exemple de code JavaScript précédent. Les valeurs de propriété et les valeurs sont en minuscules, ce qui est une meilleure pratique lorsque vous `id` `name` décrivez vos fonctions personnalisées. Vous devez ajouter ce JSON uniquement si vous préparez manuellement votre propre fichier JSON sans utiliser la génération automatique. Pour plus d’informations sur la génération automatique, voir [autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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

Découvrez les [meilleures pratiques pour nommer](custom-functions-naming.md) votre [](custom-functions-localize.md) fonction ou découvrir comment la localiser à l’aide de la méthode JSON manuscrite précédemment décrite.

## <a name="see-also"></a>Voir aussi

- [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Options des paramètres de fonctions personnalisées](custom-functions-parameter-options.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
