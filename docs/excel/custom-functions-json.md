---
title: Créer manuellement des métadonnées JSON pour des fonctions personnalisées dans Excel
description: Définissez les métadonnées JSON pour les fonctions personnalisées dans Excel et associez vos propriétés d’ID de fonction et de nom.
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4bc9139b3e46bc64749a58537737db2f048ee82
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68540997"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Créer manuellement des métadonnées JSON pour des fonctions personnalisées

Comme décrit dans l’article de [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md) , un projet de fonctions personnalisées doit inclure à la fois un fichier de métadonnées JSON et un fichier de script (JavaScript ou TypeScript) pour inscrire une fonction, ce qui le rend disponible pour une utilisation. Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après cela sont disponibles pour le même utilisateur dans tous les classeurs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Nous vous recommandons d’utiliser la génération automatique JSON dans la mesure du possible au lieu de créer votre propre fichier JSON. La génération automatique est moins sujette aux erreurs de l’utilisateur et les fichiers générés automatiquement l’incluent `yo office` déjà. Pour plus d’informations sur les balises JSDoc et le processus de génération automatique JSON, consultez [Métadonnées JSON de génération automatique pour les fonctions personnalisées](custom-functions-json-autogeneration.md).

Toutefois, vous pouvez créer un projet de fonctions personnalisées à partir de zéro. Ce processus vous oblige à :

- Écrivez votre fichier JSON.
- Vérifiez que votre fichier manifeste est connecté à votre fichier JSON.
- Associez vos fonctions `id` et `name` propriétés dans le fichier de script afin d’inscrire vos fonctions.

L’image suivante explique les différences entre l’utilisation `yo office` de fichiers automatiques et l’écriture de JSON à partir de zéro.

![Image des différences entre l’utilisation du générateur Yeoman pour les compléments Office et l’écriture de votre propre code JSON.](../images/custom-functions-json.png)

> [!NOTE]
> N’oubliez pas de connecter votre manifeste au fichier JSON que vous créez, via la **\<Resources\>** section de votre fichier manifeste XML, si vous n’utilisez pas le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Création de métadonnées et connexion au manifeste

Créez un fichier JSON dans votre projet et fournissez tous les détails sur vos fonctions, telles que les paramètres de la fonction. Consultez [l’exemple de métadonnées suivant](#json-metadata-example) et [la référence des métadonnées](#metadata-reference) pour obtenir la liste complète des propriétés de fonction.

Vérifiez que votre fichier manifeste XML fait référence à votre fichier JSON dans la **\<Resources\>** section, comme dans l’exemple suivant.

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
  "allowCustomDataForDataTypeAny": true,
  "allowErrorForDataTypeAny": true,
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
> Un exemple complet de fichier JSON est disponible dans l’historique de validation du référentiel GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) . Comme le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit n’est disponible que dans les versions précédentes du projet.

## <a name="metadata-reference"></a>Informations de référence sur les métadonnée

### <a name="allowcustomdatafordatatypeany"></a>allowCustomDataForDataTypeAny

La `allowCustomDataForDataTypeAny` propriété est un type de données booléen. La définition de cette valeur permet à `true` une fonction personnalisée d’accepter les types de données en tant que paramètres et valeurs de retour. Pour en savoir plus, consultez [Fonctions personnalisées et types de données](custom-functions-data-types-concepts.md).

> [!NOTE]
> Contrairement à la plupart des autres propriétés de métadonnées JSON, `allowCustomDataForDataTypeAny` il s’agit d’une propriété de niveau supérieur qui ne contient aucune sous-propriété. Consultez [l’exemple de code de métadonnées JSON](#json-metadata-example) précédent pour obtenir un exemple de mise en forme de cette propriété.

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

La `allowErrorForDataTypeAny` propriété est un type de données booléen. La définition de la valeur permet à `true` une fonction personnalisée de traiter les erreurs en tant que valeurs d’entrée. Tous les paramètres avec le type `any` ou `any[][]` peuvent accepter des erreurs en tant que valeurs d’entrée lorsque `allowErrorForDataTypeAny` la valeur est définie sur `true`. La valeur par défaut `allowErrorForDataTypeAny` est `false`.

> [!NOTE]
> Contrairement aux autres propriétés de métadonnées JSON, `allowErrorForDataTypeAny` il s’agit d’une propriété de niveau supérieur qui ne contient aucune sous-propriété. Consultez [l’exemple de code de métadonnées JSON](#json-metadata-example) précédent pour obtenir un exemple de mise en forme de cette propriété.

### <a name="functions"></a>fonctions

La propriété `functions` est un tableau d’objets de fonction personnalisés. Le tableau suivant répertorie les propriétés de chaque objet.

| Propriété      | Type de données | Obligatoire | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Non       | Description de la fonction que voient les utilisateurs finaux dans Excel. Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.                                                            |
| `helpUrl`     | string    | Non       | URL fournissant des informations sur la fonction (elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Oui      | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.                                            |
| `name`        | string    | Oui      | Nom de la fonction que voient les utilisateurs finaux dans Excel. Dans Excel, ce nom de fonction est préfixé par l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML. |
| `options`     | objet    | Non       | Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Reportez-vous aux [options](#options) pour plus d’informations.                                                          |
| `parameters`  | tableau     | Oui      | Tableau qui définit les paramètres d’entrée de la fonction. Pour plus d’informations, consultez [les paramètres](#parameters) .                                                                             |
| `result`      | objet    | Oui      | Objet qui définit le type d’informations renvoyées par la fonction. Reportez-vous au [résultat](#result) pour plus d’informations.                                                                 |

### <a name="options"></a>options

L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l’objet `options`.

| Propriété          | Type de données | Obligatoire                               | Description |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction. Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui retournent un seul résultat et doivent gérer l’annulation d’une demande de données. Une fonction ne peut pas utiliser à la fois les propriétés et `cancelable` les `stream` propriétés. |
| `requiresAddress` | valeur booléenne   | Non <br/><br/>La valeur par défaut est `false`. | Si `true`, votre fonction personnalisée peut accéder à l’adresse de la cellule qui l’a appelée. La `address` propriété du [paramètre d’appel](custom-functions-parameter-options.md#invocation-parameter) contient l’adresse de la cellule qui a appelé votre fonction personnalisée. Une fonction ne peut pas utiliser à la fois les propriétés et `requiresAddress` les `stream` propriétés. |
| `requiresParameterAddresses` | valeur booléenne   | Non <br/><br/>La valeur par défaut est `false`. | Si `true`, votre fonction personnalisée peut accéder aux adresses des paramètres d’entrée de la fonction. Cette propriété doit être utilisée en combinaison avec la `dimensionality` propriété de l’objet [de résultat](#result) et `dimensionality` doit être définie sur `matrix`. Pour plus [d’informations, consultez Détecter l’adresse d’un paramètre](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) . |
| `stream`          | valeur booléenne   | Non<br/><br/>La valeur par défaut est `false`.  | Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique. Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières. La fonction ne doit pas utiliser d’instruction `return`. Au lieu de cela, la valeur du résultat est passée comme argument de la `StreamingInvocation.setResult` fonction de rappel. Pour plus d’informations, consultez [Créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function). |
| `volatile`        | valeur booléenne   | Non <br/><br/>La valeur par défaut est `false`. | Si `true`, la fonction recalcule chaque fois qu’Excel recalcule, au lieu de seulement lorsque les valeurs dépendantes de la formule ont changé. Une fonction ne peut pas utiliser à la fois les propriétés et `volatile` les `stream` propriétés. Si les propriétés et `volatile` les `stream` propriétés sont définies `true`sur , la propriété volatile est ignorée. |

### <a name="parameters"></a>paramètres

La propriété `parameters` est un tableau d’objets paramètre. Le tableau suivant répertorie les propriétés de chaque objet.

|  Propriété  |  Type de données  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Non |  Description du paramètre. Cela s’affiche dans IntelliSense d’Excel.  |
|  `dimensionality`  |  string  |  Non  |  Doit être `scalar` (valeur non matricielle) ou `matrix` (tableau à 2 dimensions).  |
|  `name`  |  string  |  Oui  |  Le nom du paramètre. Ce nom s’affiche dans IntelliSense d’Excel.  |
|  `type`  |  string  |  Non  |  Type de données du paramètre. Peut être `boolean`, `number`, `string`ou `any`, qui vous permet d’utiliser l’un des trois types précédents. Si cette propriété n’est pas spécifiée, le type de données est défini par défaut `any`sur . |
|  `optional`  | valeur booléenne | Non | Si la valeur est `true`, le paramètre est facultatif. |
|`repeating`| valeur booléenne | Non | Si `true`, les paramètres sont renseignés à partir d’un tableau spécifié. Notez que les fonctions de tous les paramètres répétitifs sont considérées comme des paramètres facultatifs par définition.  |

### <a name="result"></a>result

L’objet `result` définit le type des informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l’objet `result`.

| Propriété         | Type de données | Obligatoire | Description                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Non       | Doit être `scalar` (valeur non matricielle) ou `matrix` (tableau à 2 dimensions). |
| `type` | string    | Non       | Type de données du résultat. Peut être `boolean`, `number`, `string`ou `any` (ce qui vous permet d’utiliser l’un des trois types précédents). Si cette propriété n’est pas spécifiée, le type de données est défini par défaut `any`sur . |

## <a name="associating-function-names-with-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de `id` la fonction à l’implémentation JavaScript. Assurez-vous qu’il existe une association, sinon la fonction ne sera pas inscrite et n’est pas utilisable dans Excel. L’exemple de code suivant montre comment créer l’association à l’aide de la `CustomFunctions.associate()` fonction. L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est **AJOUTER**.

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

Le code JSON suivant montre les métadonnées JSON associées au code JavaScript de la fonction personnalisée précédente.

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

L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans l’exemple de code JavaScript précédent. Les `id` valeurs et `name` les valeurs de propriété sont en majuscules, ce qui est une bonne pratique lors de la description de vos fonctions personnalisées. Vous n’avez besoin d’ajouter ce fichier JSON que si vous préparez votre propre fichier JSON manuellement et que vous n’utilisez pas la génération automatique. Pour plus d’informations sur la génération automatique, consultez [Métadonnées JSON de génération automatique pour les fonctions personnalisées](custom-functions-json-autogeneration.md).

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

## <a name="next-steps"></a>Prochaines étapes

Découvrez les [meilleures pratiques pour nommer votre fonction](custom-functions-naming.md) ou découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.

## <a name="see-also"></a>Voir aussi

- [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Options des paramètres des fonctions personnalisées](custom-functions-parameter-options.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
