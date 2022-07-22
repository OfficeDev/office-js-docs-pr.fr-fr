---
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: da51afbcc56a86d74a9ab4edf2ebf283436196d5
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958403"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Générer automatiquement des métadonnées JSON pour des fonctions personnalisées

Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les [balises JSDoc](https://jsdoc.app/) pour la détailler en ajoutant des informations supplémentaires. Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le fichier de métadonnées JSON. L’utilisation de balises JSDoc vous évite de [modifier manuellement le fichier de métadonnées JSON](custom-functions-json.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.

Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript. Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).

## <a name="add-a-description-to-a-function"></a>Ajouter une description à une fonction

La description s’affiche pour l’utilisateur sous forme de texte d’aide lorsqu’il a besoin d’aide pour comprendre le rôle de votre fonction personnalisée. La description ne nécessite aucune balise spécifique. Il vous suffit d’entrer une brève description dans le commentaire JSDoc. En général, la description est placée au début de la section commentaires JSDoc, mais elle fonctionnera peu importe son emplacement.

Pour consulter des exemples de descriptions de fonction intégrées, ouvrez Excel, accédez à l’onglet **formules** , puis sélectionnez **insérer une fonction**. Vous pouvez ensuite parcourir toutes les descriptions de fonction et voir vos propres fonctions personnalisées répertoriées.

Dans cet exemple, la phrase «calcule le volume d’une sphère.» est la description de la fonction personnalisée.

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```

## <a name="jsdoc-tags"></a>Balises JSDoc

Les balises JSDoc suivantes sont prises en charge dans les fonctions personnalisées Excel.

- [@ annulable](#cancelable)
- [@customfunction](#customfunction) *nom* *d’ID*
- [URL @helpurl](#helpurl) 
- [@param](#param) *description* du *nom* *{type}*
- [@requièreuneadresse](#requiresAddress)
- [@requiresParameterAddresses](#requiresParameterAddresses)
- [@renvoie](#returns) *{type}*
- [@diffusionencontinu](#streaming)
- [@volatile](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a>@ annulable

Indique qu’une fonction personnalisée effectue une action lorsque la fonction est annulée.

Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`. La fonction peut affecter une fonction à la `oncanceled` propriété pour indiquer le résultat lorsque la fonction est annulée.

Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.

Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.

<a id="customfunction"></a>

### <a name="customfunction"></a>@fonctionpersonnalisée

Syntaxe: @fonctionpersonnalisée *id* *nom*

Cette balise indique que la fonction JavaScript/TypeScript est une fonction personnalisée Excel. Il est nécessaire de créer des métadonnées pour la fonction personnalisée.

Voici un exemple de cette balise.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

Identifie `id` une fonction personnalisée.

- Si`id`n’est pas fourni, le nom de la fonction JavaScript/TypeScript est converti en majuscules, et les caractères rejetés sont supprimés.
- Le `id`doit être unique pour toutes les fonctions personnalisées.
- Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).

Dans l’exemple suivant, Increments correspond à l’`id` et au `name` de la fonction.

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>name

Fournit le nom d’affichage `name`de la fonction personnalisée.

- Si aucun nom n’est fourni, l’id servira aussi de nom.
- Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).
- Doit commencer par une lettre.
- Sa longueur maximale est limitée à 128 caractères.

Dans l’exemple suivant, INC correspond à l’`id` de la fonction, tandis que `increment` correspond au `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>description

Une description s’affiche aux utilisateurs dans Excel lorsqu’ils entrent dans la fonction et spécifie ce que fait la fonction. Une description ne nécessite aucune balise spécifique. Ajoutez une description à une fonction personnalisée en ajoutant une expression pour décrire le rôle de la fonction dans le commentaire JSDoc. Par défaut, le texte non balisé dans la section commentaire JSDoc est la description de la fonction.

Dans l’exemple suivant, la phrase « A function that adds two numbers » (« Une fonction qui ajoute deux nombres ») est la description de la fonction personnalisée dont la propriété ID est `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>

### <a name="helpurl"></a>@urlaide

Syntaxe: @urlaide *url*

L’*url* fournie est affichée dans Excel.

Dans l’exemple suivant, il s’agit `www.contoso.com/weatherhelp`de `helpurl` .

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>

### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

Syntaxe JavaScript : @param *description* du *nom* {type}

- `{type}` spécifie les informations de type entre accolades. Consultez la section [Types](#types) pour savoir quels types peuvent être utilisés. Si aucun type n’est spécifié, le type `any` par défaut est utilisé.
- `name` spécifie le paramètre auquel la balise @param s’applique. C’est obligatoire.
- `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Elle est facultative.

Pour indiquer un paramètre de fonction personnalisé comme facultatif, placez des crochets autour du nom du paramètre. Par exemple : `@param {string} [text] Optional text`.

> [!NOTE]
> La valeur par défaut pour les paramètres facultatifs est `null`.

L’exemple suivant montre une fonction ADD qui ajoute deux ou trois nombres, avec le troisième nombre comme paramètre facultatif.

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

Syntaxe TypeScript : *@param description* *du nom*

- `name` spécifie le paramètre auquel la balise @param s’applique. C’est obligatoire.
- `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Elle est facultative.

Consultez la section [Types](#types) pour savoir quels types de paramètres de fonction peuvent être utilisés.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :

- Utilisez un paramètre facultatif. Par exemple : `function f(text?: string)`
- Définissez ce paramètre sur une valeur par défaut. Par exemple : `function f(text: string = "abc")`

Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> La valeur par défaut pour les paramètres facultatifs est `null`.

L’exemple suivant représente la fonction `add` qui ajoute deux nombres.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requièreuneadresse

Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.

Le dernier paramètre de fonction doit être de type `CustomFunctions.Invocation` ou dérivé à utiliser `@requiresAddress`. Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.

L’exemple suivant montre comment utiliser le `invocation` paramètre en combinaison avec `@requiresAddress` pour retourner l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour plus d’informations, consultez [le paramètre Invocation](custom-functions-parameter-options.md#invocation-parameter) .

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  const address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>

### <a name="requiresparameteraddresses"></a>@requiresParameterAddresses

Indique que la fonction doit retourner les adresses des paramètres d’entrée.

Le dernier paramètre de fonction doit être de type `CustomFunctions.Invocation` ou dérivé à utiliser  `@requiresParameterAddresses`. Le commentaire JSDoc doit également inclure une `@returns` balise spécifiant que la valeur de retour est une matrice, telle que `@returns {string[][]}` ou `@returns {number[][]}`. Pour plus d’informations, consultez [Types de matrices](#matrix-type) .

Lorsque la fonction est appelée, la `parameterAddresses` propriété contient les adresses des paramètres d’entrée.

L’exemple suivant montre comment utiliser le `invocation` paramètre en combinaison avec `@requiresParameterAddresses` pour retourner les adresses de trois paramètres d’entrée. Pour plus [d’informations, consultez Détecter l’adresse d’un paramètre](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) .

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>

### <a name="returns"></a>@renvoie :

Syntaxe: @renvoie {*type*}

Fournit le type pour la valeur renvoyée.

Si `{type}` est omis, les informations de type TypeScript seront utilisées. S’il n’existe aucune information définissant le type, ce dernier sera `any`.

L’exemple suivant représente la fonction `add` qui utilise la balise `@returns`.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>

### <a name="streaming"></a>@diffusionencontinu

Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.

Le dernier paramètre est de type `CustomFunctions.StreamingInvocation<ResultType>`.
La fonction retourne `void`.

Les fonctions de diffusion en continu ne retournent pas de valeurs directement, mais elles appellent `setResult(result: ResultType)` à l’aide du dernier paramètre.

Les exceptions levées par une fonction en continu sont ignorées. `setResult()`peut être appelée avec Error pour indiquer un résultat erroné. Si vous souhaitez consulter un exemple de fonction de diffusion en continu et obtenir d’autres informations, veuillez vous reporter à la section [Créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).

Les fonctions de diffusion en continu ne peuvent pas être marquées comme étant [@volatile](#volatile).

<a id="volatile"></a>

### <a name="volatile"></a>@volatile

Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas. À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes. C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.

Les fonctions de diffusion en continu ne peuvent pas être volatiles.

La fonction suivante est volatile et utilise la balise `@volatile`.

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>Types

En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction. Si le type est `any`, Excel n’effectue pas de conversion.

### <a name="value-types"></a>Types de valeur

Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.

### <a name="matrix-type"></a>Type matrice

Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs. Par exemple, le type `number[][]` indique une matrice de nombres et `string[][]` une matrice de chaînes.

### <a name="error-type"></a>Type d’erreur

Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.

Une fonction de diffusion en continu peut indiquer une erreur en appelant`setResult()`avec un type Error.

### <a name="promise"></a>Promise

Une fonction personnalisée peut retourner une promesse qui fournit la valeur lorsque la promesse est résolue. Si la promesse est rejetée, la fonction personnalisée génère une erreur.

### <a name="other-types"></a>Autres types

Tout autre type sera traité comme une erreur.

## <a name="next-steps"></a>Étapes suivantes

Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md). Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).

## <a name="see-also"></a>Voir aussi

- [Créer manuellement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
