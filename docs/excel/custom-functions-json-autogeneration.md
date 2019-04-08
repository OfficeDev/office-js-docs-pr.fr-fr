---
ms.date: 04/03/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478955"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a>Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)

Lorsqu’ une fonction personnalisée Excel est écrite dans JavaScript ou TypeScript, les balises JSDoc servent à fournir des informations supplémentaires sur la fonction personnalisée. Les balises JSDoc sont ensuite utilisées au moment de build pour créer le [fichier de métadonnées JSON](custom-functions-json.md). Utiliser des balises JSDoc vous évite des efforts pour modifier manuellement le fichier de métadonnées JSON.

Ajouter la`@customfunction` balise dans les commentaires du code d’une fonction JavaScript ou TypeScript pour la marquer comme une fonction personnalisée.

La fonction types de paramètre peut être fournie à l’aide de la [@param ](#param) balise dans JavaScript, ou via la [fonction type](http://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript. Pour plus d’informations, consultez la [@param](#param) balise et la section[Types](#Types).

## <a name="jsdoc-tags"></a>Balises JSDoc
Les balises JSDoc suivants sont prises en charge dans les fonctions personnalisées Excel :
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) nom d’ID
* [@helpurl](#helpurl)URL
* [@param](#param) _{type}_ description nom
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.

Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`. La fonction peut affecter une fonction à la `oncanceled` propriété pour désigner l’action à effectuer lorsque la fonction est annulée.

Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, sera considéré comme `@cancelable` même si la balise n’apparaît pas.

Une fonction ne peut pas contenir les deux balises`@cancelable` et `@streaming`.

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

Syntaxe : @customfunction _id_ _nom_

Spécifier cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction personnalisée Excel.

Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.

Il doit également être un appel au `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id 

L’id est utilisé en tant qu’identificateur indifférent pour la fonction personnalisée stockée dans le document. Elle ne doit pas changer.

* Si l’id n’est pas inclus, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, les caractères rejetés sont supprimés.
* L’id doit être unique pour toutes les fonctions personnalisées.
* Les caractères autorisés sont limités aux : A-Z, a-z, 0-9 et point (.).

#### <a name="name"></a>nom

Fournit le nom d’affichage pour la fonction personnalisée. 

* Si aucun nom n’est fourni, l’id est également utilisé comme le nom.
* Caractères autorisés : lettres [caractère Unicode alphabétique](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), nombres, point (.) et un trait de soulignement (\_).
* Doit commencer par une lettre.
* La longueur maximale est de 128 caractères.

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

Syntaxe: @helpurl_url_

L’_url_ fournie est affichée dans Excel.

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

Syntaxe JavaScript : @param nom {type} _description_

* `{type}` doit spécifier les informations de type au sein des accolades. Voir les[Types](##types) pour plus d’informations sur les types qui peuvent être utilisés. Facultatif: Si aucun serveur n'est spécifié, le type`any` sera utilisé.
* `name` spécifie le paramètre auquel la@parambalise s’applique. Obligatoire.
* `description` fournit la description qui s’affiche dans Excel pour le paramètre de fonction. Facultatif.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif:
* Placez les crochets autour du paramètre de nom. Par exemple : `@param {string} [text] Optional text`.

#### <a name="typescript"></a>TypeScript

Syntaxe JavaScript : @paramnom_description_

* `name` spécifie le paramètre auquel la@parambalise s’applique. Obligatoire.
* `description` fournit la description qui s’affiche dans Excel pour le paramètre de fonction. Facultatif.

Voir les[Types](##types) pour plus d’informations sur les types de paramètre de fonction qui peuvent être utilisés.

Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes:
* Utilisez un paramètre facultatif. Par exemple : `function f(text?: string)`
* Donne une valeur par défaut au paramètre. Par exemple : `function f(text: string = "abc")`

Pour une description détaillée du @paramvoir:[JSDoc](http://usejsdoc.org/tags-param.html)

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie. 

Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé. Lorsque la fonction est appelée, la`address` propriété contiendra l’adresse.

---
### <a name="returns"></a>@returns
<a id="returns"/>

Syntaxe : @returns {_type_}

Fournit le type pour la valeur de retour.

Si `{type}` est omis, les informations de type TypeScript seront utilisées. S’il n’existe aucune information type, le type sera `any`.

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu. 

Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.
La fonction doit retourner `void`.

Les fonctions de diffusion en continu ne retournent pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.

Les exceptions levées par une fonction en continu sont ignorées. `setResult()` peut être appelée avec l’erreur pour indiquer un résultat de l’erreur.

Les fonctions en continu ne peuvent pas être marquées comme étant [ @volatile ](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Une fonction volatile est une dont le résultat ne peut pas être considéré comme le même à partir d’un moment à l’autre même si elle ne prend aucun argument ou les arguments n’ont pas changé. Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes les cellules dépendantes, chaque fois qu’il effectue un recalcul. C’est pourquoi trop de dépendance des fonctions volatiles peut ralentir le recalcul.

Les fonctions en continu ne peuvent pas être volatiles.

---

## <a name="types"></a>Types

En spécifiant un type de paramètre, Excel convertit les valeurs dans ce type avant d’appeler la fonction. Si le type est`any`, aucune opération de conversion n’est effectuée.

### <a name="value-types"></a>Types de valeur

Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number`, `string`.

### <a name="matrix-type"></a>Type matrice

Utilisez une matrice à deux dimensions pour lesquels le paramètre ou la valeur de retour peut être une matrice de valeurs. Par exemple, le type `number[][]` indique une matrice de nombres. `string[][]` Indique une matrice de chaînes. 

### <a name="error-type"></a>Type d’erreur

Une fonction en non continu peut indiquer une erreur en retournant un type d’erreur.

Une fonction en continu peut indiquer une erreur en retournant un type d’erreur().

### <a name="promise"></a>Promesse

Une fonction peut renvoyer une promesse qui fournit la valeur lorsque la promesse aura été résolue. Si la promesse est refusée, alors elle est considérée comme une erreur.

### <a name="other-types"></a>Autres types

Un autre type sera traité comme une erreur.

## <a name="see-also"></a>Voir aussi

* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Débogage des métadonnées des fonctions personnalisées](custom-functions-debugging.md)
