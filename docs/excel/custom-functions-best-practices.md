---
ms.date: 09/20/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068815"
---
# <a name="custom-functions-best-practices"></a>Meilleures pratiques pour les fonctions personnalisées

Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.

## <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="error-logging"></a>Journalisation des erreurs

Vous pouvez activer la journalisation des erreurs pour votre complément de fonctions personnalisées de plusieurs façons, telles que : 

- [Utiliser la journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) pour le débogage du fichier manifeste XML de votre complément. 

- Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.

> [!NOTE]
> La fonctionnalité de journalisation runtime n'est actuellement disponible que pour Office 2016 bureau.

## <a name="debugging"></a>Débogage

Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à utiliser [Excel Online](https://www.office.com/launch/excel) et l’outil de débogage F12 natif de votre navigateur. Des outils de débogage spécifiques pour les fonctions personnalisées pourraient être disponibles à l’avenir.

## <a name="mapping-names"></a>Mappage de noms

Par défaut, le nom d’une fonction personnalisée dans votre fichier JavaScript est déclaré généralement à l’aide de lettres toutes en majuscule et correspond exactement au nom de la fonction que l'utilisateur final voit dans Excel. Toutefois, vous pouvez modifier ce mappage à l’aide de l'objet `CustomFunctionsMappings` pour mapper un ou plusieurs noms de fonction à partir du fichier JavaScript à des valeurs différentes que les utilisateurs finaux verront s’afficher comme noms de fonction dans Excel. Bien que vous ne soyez pas obligé d’utiliser `CustomFunctionsMapping`, il peut être utile si vous utilisez un uglifier, un webpack ou une syntaxe d'importation - qui tous ont des difficultés avec les noms de fonctions en majuscules.
  
L’exemple de code suivant définit une seule paire clé-valeur qui mappe le nom de la fonction JavaScript `plusFortyTwo` au nom de la fonction `ADD42` dans l’interface utilisateur d’Excel. Lorsque l’utilisateur final choisit la fonction `ADD42` dans Excel, la fonction JavaScript `plusFortyTwo` s’exécute.

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

L’exemple de code suivant définit deux paires clé-valeur. La première paire mappe le nom de la fonction JavaScript `plusFifty` au nom de la fonction `ADD50` dans l’interface utilisateur d’Excel et la seconde paire mappe le nom de la fonction JavaScript `plusOneHundred` au nom de la fonction `ADD100` dans l’interface utilisateur d’Excel. Lorsque l’utilisateur final choisit la fonction `ADD50` dans Excel, la fonction JavaScript `plusFifty` s’exécute. Lorsque l’utilisateur final choisit la fonction `ADD100` dans Excel, la fonction JavaScript `plusOneHundred` s’exécute.

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)