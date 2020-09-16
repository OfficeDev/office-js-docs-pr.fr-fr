---
ms.date: 05/06/2020
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 8138e738188e50d2a1369c359fbca3e1574db32f
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819517"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Générer automatiquement des métadonnées JSON pour des fonctions personnalisées

Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les [balises JSDoc](https://jsdoc.app/) pour la détailler en ajoutant des informations supplémentaires. Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md). En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.

Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript. Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).

### <a name="adding-a-description-to-a-function"></a>Ajout d’une description à une fonction

La description s’affiche pour l’utilisateur sous forme de texte d’aide lorsqu’il a besoin d’aide pour comprendre le rôle de votre fonction personnalisée. La description ne nécessite aucune balise spécifique. Il vous suffit d’entrer une brève description dans le commentaire JSDoc. En général, la description est placée au début de la section commentaires JSDoc, mais elle fonctionnera peu importe son emplacement.

Pour consulter des exemples de descriptions de fonction intégrées, ouvrez Excel, accédez à l’onglet**formules** , puis sélectionnez **insérer une fonction**. Vous pouvez ensuite parcourir toutes les descriptions de fonction et voir vos propres fonctions personnalisées répertoriées.

Dans cet exemple, la phrase «calcule le volume d’une sphère.» est la description de la fonction personnalisée.

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a>Balises JSDoc
Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :
* [@ annulable](#cancelable)
* [@fonctionpersonnalisée](#customfunction)nom id
* url[@urlaide](#helpurl)
* [@param](#param) _{type}_ description nom
* [@requièreuneadresse](#requiresAddress)
* [@renvoie](#returns) _{type}_
* [@diffusionencontinu](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@ annulable
<a id="cancelable"/>

Indique qu’une fonction personnalisée effectue une action lorsque la fonction est annulée.

Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`. La fonction peut affecter une fonction à la `oncanceled` propriété pour indiquer le résultat lorsque la fonction est annulée.

Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.

Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.

---
### <a name="customfunction"></a>@fonctionpersonnalisée
<a id="customfunction"/>

Syntaxe: @fonctionpersonnalisée_id_ _nom_

Cette balise indique que la fonction JavaScript/dactylographié est une fonction personnalisée Excel. Il est nécessaire de créer des métadonnées pour la fonction personnalisée.

Voici un exemple de cette balise.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

L' `id` identifie une fonction personnalisée.

* Si`id`n’est pas fourni, le nom de la fonction JavaScript/TypeScript est converti en majuscules, et les caractères rejetés sont supprimés.
* Le `id`doit être unique pour toutes les fonctions personnalisées.
* Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).

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

* Si aucun nom n’est fourni, l’id servira aussi de nom.
* Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).
* Doit commencer par une lettre.
* Sa longueur maximale est limitée à 128 caractères.

Dans l’exemple suivant, INC correspond à l’`id` de la fonction, tandis que `increment` correspond au `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>description

Une description apparaît pour les utilisateurs dans Excel lorsqu’ils entrent dans la fonction et spécifie le rôle de la fonction. Une description ne nécessite aucune balise spécifique. Ajoutez une description à une fonction personnalisée en ajoutant une expression pour décrire le rôle de la fonction dans le commentaire JSDoc. Par défaut, le texte non balisé dans la section commentaire JSDoc est la description de la fonction.

Dans l’exemple suivant, la phrase « A function that adds two numbers » (« Une fonction qui ajoute deux nombres ») est la description de la fonction personnalisée dont la propriété ID est `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a>@urlaide
<a id="helpurl"/>

Syntaxe: @urlaide_url_

L’_url_ fournie est affichée dans Excel.

Dans l’exemple suivant, le `helpurl` est `www.contoso.com/weatherhelp` .

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

Syntaxe JavaScript : @param {type} nom_description_

* `{type}` spécifie les informations de type entre accolades. Consultez la section [Types](#types) pour savoir quels types peuvent être utilisés. Si aucun type n’est spécifié, le type par défaut est `any` utilisé.
* `name` Spécifie le paramètre auquel s’applique la balise @param. Elle est obligatoire.
* `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Elle est facultative.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif :
* Placez les crochets autour du nom du paramètre. Par exemple : `@param {string} [text] Optional text`.

> [!NOTE]
> La valeur par défaut pour les paramètres facultatifs est `null`.

L’exemple suivant montre une fonction ADD qui ajoute deux ou trois nombres, le troisième étant un paramètre facultatif.

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

Syntaxe TypeScript : nom @param_description_

* `name` Spécifie le paramètre auquel s’applique la balise @param. Elle est obligatoire.
* `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Elle est facultative.

Consultez la section [Types](#types) pour savoir quels types de paramètres de fonction peuvent être utilisés.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :
* Utilisez un paramètre facultatif. Par exemple : `function f(text?: string)`
* Définissez ce paramètre sur une valeur par défaut. Par exemple : `function f(text: string = "abc")`

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

---
### <a name="requiresaddress"></a>@requièreuneadresse
<a id="requiresAddress"/>

Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.

Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé. Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.

---
### <a name="returns"></a>@renvoie :
<a id="returns"/>

Syntaxe: @renvoie {_type_}

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

---
### <a name="streaming"></a>@diffusionencontinu
<a id="streaming"/>

Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu. 

Le dernier paramètre est de type `CustomFunctions.StreamingInvocation<ResultType>` .
La fonction renvoie `void` .

Les fonctions de diffusion en continu ne renvoient pas directement de valeurs, mais appelent `setResult(result: ResultType)` à l’aide du dernier paramètre.

Les exceptions levées par une fonction en continu sont ignorées. `setResult()`peut être appelée avec Error pour indiquer un résultat erroné. Si vous souhaitez consulter un exemple de fonction de diffusion en continu et obtenir d’autres informations, veuillez vous reporter à la section [Créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).

Les fonctions de diffusion en continu ne peuvent pas être marquées comme étant [@volatile](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas. À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes. C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.

Les fonctions de diffusion en continu ne peuvent pas être volatiles.

La fonction suivante est volatile et utilise la balise `@volatile`.

```js
/**
 * Simulates rolling a 6-sided dice.
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

Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs. Par exemple, le type `number[][]` indique une matrice de nombres. `string[][]`indique une matrice de chaînes.

### <a name="error-type"></a>Type d’erreur

Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.

Une fonction de diffusion en continu peut indiquer une erreur en appelant`setResult()`avec un type Error.

### <a name="promise"></a>Promise

Une fonction peut renvoyer une promesse, qui fournit la valeur lorsque la promesse est résolue. Si la promesse est rejetée, elle génère une erreur.

### <a name="other-types"></a>Autres types

Tout autre type sera traité comme une erreur.

## <a name="next-steps"></a>Étapes suivantes
Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md). Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
