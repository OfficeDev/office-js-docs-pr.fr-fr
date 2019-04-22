---
ms.date: 04/03/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914283"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a>Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)

Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les balises JSDoc pour la détailler en ajoutant des informations supplémentaires. Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md). En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.

Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.

Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript. Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).

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

Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.

Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`. La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.

Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.

Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.

---
### <a name="customfunction"></a>@fonctionpersonnalisée
<a id="customfunction"/>

Syntaxe: @fonctionpersonnalisée_id_ _nom_

Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.

Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.

Vous devez également insérer un appel vers`CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id 

L’id est utilisé en tant qu’identificateur invariant pour la fonction personnalisée stockée dans le document. Elle ne doit pas changer.

* Si l’id n’est pas fourni, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, et les caractères rejetés sont supprimés.
* L’id doit être unique pour toutes les fonctions personnalisées.
* Seuls les caractères alphanumériques majuscules et minuscules (A-Z, a-z, 0-9) et le point (.) sont autorisés.

#### <a name="name"></a>name

Fournit le nom d’affichage de la fonction personnalisée. 

* Si aucun nom n’est fourni, l’id servira aussi de nom.
* Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).
* Doit commencer par une lettre.
* Sa longueur maximale est limitée à 128 caractères.

---
### <a name="helpurl"></a>@urlaide
<a id="helpurl"/>

Syntaxe: @urlaide_url_

L’_url_ fournie est affichée dans Excel.

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

Syntaxe JavaScript : @param {type} nom_description_

* `{type}`doit spécifier les informations de type entre deux accolades. Consultez la section [Types](##types) pour savoir quels types peuvent être utilisés. Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.
* `name`spécifie le paramètre auquel s’applique la balise. Obligatoire.
* `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Facultatif.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif :
* Placez les crochets autour du nom du paramètre. Par exemple : `@param {string} [text] Optional text`.

#### <a name="typescript"></a>TypeScript

Syntaxe TypeScript : nom @param_description_

* `name`spécifie le paramètre auquel s’applique la balise. Obligatoire.
* `description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction. Facultatif.

Consultez la section [Types](##types) pour savoir quels types de paramètres de fonction peuvent être utilisés.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :
* Utilisez un paramètre facultatif. Par exemple : `function f(text?: string)`
* Définissez ce paramètre sur une valeur par défaut. Par exemple : `function f(text: string = "abc")`

Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](http://usejsdoc.org/tags-param.html)

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

---
### <a name="streaming"></a>@diffusionencontinu
<a id="streaming"/>

Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu. 

Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.
La fonction doit renvoyer `void`.

Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.

Les exceptions levées par une fonction en continu sont ignorées. `setResult()`peut être appelée avec Error pour indiquer un résultat erroné.

Vous ne pouvez pas utiliser les balises en diffusion en continu comme [@volatile](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas. À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes. C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.

Les fonctions de diffusion en continu ne peuvent pas être volatiles.

---

## <a name="types"></a>Types

En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction. Si le type est `any`, Excel n’effectue pas de conversion.

### <a name="value-types"></a>Types de valeur

Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.

### <a name="matrix-type"></a>Type matrice

Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs. Par exemple, le type `number[][]` indique une matrice de nombres. `string[][]`indique une matrice de chaînes. 

### <a name="error-type"></a>Type d’erreur

Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.

Une fonction de diffusion en continu peut indiquer une erreur en appelant la méthode setResult() avec un type Error.

### <a name="promise"></a>Promise

Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution. Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.

### <a name="other-types"></a>Autres types

Tout autre type sera traité comme une erreur.

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Débogage des fonctions personnalisées](custom-functions-debugging.md)
