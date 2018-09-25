---
ms.date: 09/20/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004909"
---
# <a name="custom-functions-best-practices"></a>Meilleures pratiques pour les fonctions personnalisées

Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.

## <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
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

## <a name="debugging"></a>Débogage
Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à premier [sideload](../testing/sideload-office-add-ins-for-testing.md) votre complément dans **Excel Online**. Ensuite, vous pouvez déboguer vos fonctions personnalisées à l’aide de l’[outil de débogage F12 natif de votre navigateur](../testing/debug-add-ins-in-office-online.md). Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.

Si votre complément ne parvient pas à s’enregistrer, [vérifiez que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web qui héberge votre application de complément.

Si vous testez votre complément dans Office 2016 bureau, vous pouvez activer la [journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) pour résoudre les problèmes du fichier manifeste XML de votre complément, ainsi que plusieurs conditions d’installation et d’exécution. 


## <a name="mapping-names"></a>Mappage de noms

Par défaut, le nom d’une fonction personnalisée dans votre fichier JavaScript est déclaré généralement à l’aide de lettres toutes en majuscule et correspond exactement au nom de la fonction que l'utilisateur final voit dans Excel. Toutefois, vous pouvez modifier ce mappage à l’aide de l'objet `CustomFunctionsMappings` pour mapper un ou plusieurs noms de fonction à partir du fichier JavaScript à des valeurs différentes que les utilisateurs finaux verront s’afficher comme noms de fonction dans Excel. Cela peut être utile si vous utilisez un uglifier, un webpack ou une syntaxe d’importation - qui ont tous des difficultés avec les noms de fonctions en majuscules. `CustomFunctionsMappings` Il est éventuellement facultatif pour les projets utilisant JavaScript, mais vous devez vous en servir si votre projet utilise des caractères dactylographiés.  
  
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

- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Métadonnées des fonctions personnalisées](custom-functions-json.md)
- [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md)
