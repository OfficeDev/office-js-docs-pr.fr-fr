---
ms.date: 10/03/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: f6781de97f912df70800532032162187ae9f9344
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459111"
---
# <a name="custom-functions-best-practices-preview"></a>Meilleures pratiques pour les fonctions personnalisées (aperçu)

Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
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

Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à premier [sideload](../testing/sideload-office-add-ins-for-testing.md) votre complément dans **Excel Online**. Ensuite, vous pouvez déboguer vos fonctions personnalisées à l’aide de l’[outil de débogage F12 natif de votre navigateur](../testing/debug-add-ins-in-office-online.md) en combinaison avec les techniques suivantes :

- Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.

- Utilisez les instructions `debugger;` au sein de votre code des fonctions personnalisées pour spécifier les points d’arrêt où l’exécution s’interrompra lorsque la fenêtre F12 est ouverte. Par exemple, si la fonction suivante s’exécute alors que la fenêtre F12 est ouverte, l’exécution s’interrompra sur l’instruction `debugger;`, ce qui vous permettra d’inspecter manuellement les valeurs de paramètre avant le retour de la fonction.L’instruction `debugger;` n’a aucun effet dans Excel Online lorsque la fenêtre F12 n’est pas ouverte. Actuellement, les instructions `debugger;` n’ont aucun effet dans Excel pour Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

Si votre complément ne parvient pas à s’enregistrer, [vérifiez que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web qui héberge votre application de complément.

Si vous testez votre complément dans Office sur le bureau Windows, vous pouvez activer la [journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) pour résoudre les problèmes du fichier manifeste XML de votre complément, ainsi que plusieurs conditions d’installation et d’exécution.

## <a name="mapping-function-names-to-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires à Excel pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux. En outre, dans le fichier JavaScript qui définit vos fonctions personnalisées, vous devez fournir les informations pour spécifier l’objet de fonction dans le fichier de métadonnées JSON correspondant à chaque fonction personnalisée dans le fichier JavaScript.

Par exemple, l’exemple de code suivant définit la fonction personnalisée `add`, puis spécifie que la fonction `add` correspond à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **ADD**.

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

Gardez à l’esprit les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et en spécifiant les informations correspondantes dans le fichier de métadonnées JSON.

* Dans le fichier JavaScript, spécifiez les noms de fonction en casse mixte. Par exemple, le nom de la fonction `addTenToInput` est écrit en casse mixte : le premier mot dans le nom commence par une lettre minuscule, et chaque mot suivant dans le nom commence par une lettre majuscule.

* Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque propriété `name` en majuscules. La propriété `name`  définit le nom de la fonction que les utilisateurs finaux verront s’afficher dans Excel. L’utilisation de lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonctions intégrées sont en majuscules.

* Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque propriété `id` en majuscules. Ainsi, il est évident quelle partie de l’instruction `CustomFunctionMappings`  dans votre code JavaScript correspond à la propriété `id`    dans le fichier de métadonnées JSON (à condition que votre nom de la fonction utilise CamelCase, comme indiqué précédemment).

* Dans le fichier de métadonnées JSON, assurez-vous que la valeur de chaque propriété `id` est unique dans l’étendue du fichier. Autrement dit, deux objets fonctions dans le fichier de métadonnées ne doivent pas avoir la même valeur `id`. En outre, ne spécifiez pas deux valeurs `id`  dans le fichier de métadonnées qui diffèrent uniquement par la casse. Par exemple, ne définissez pas un objet fonction avec une valeur `id`  de **add** et un autre objet fonction avec une valeur `id`  de **ADD**.

* Ne modifiez pas la valeur d’une propriété `id` dans le fichier de métadonnées JSON après qu’il a été mappé à un nom de fonction JavaScript correspondant. Vous pouvez modifier le nom de la fonction que les utilisateurs voient dans Excel en mettant à jour la propriété `name`  dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une propriété `id`  une fois établie.

* Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées au même endroit. Par exemple, l’exemple de code suivant définit deux fonctions personnalisées puis spécifie les informations de mappage pour les deux fonctions.

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans cet exemple de code JavaScript.

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

## <a name="additional-considerations"></a>Considérations supplémentaires

Pour créer un complément qui s’exécute sur plusieurs plates-formes (l’un des principaux clients des compléments Office), vous ne devez pas accéder au DOM (Document Object Model) dans des fonctions personnalisées ni utiliser des bibliothèques telles que jQuery qui s’appuient sur le modèle DOM. Dans Excel pour Windows, où les fonctions personnalisées utilisent le [runtime JavaScript](custom-functions-runtime.md), des fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Runtime de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)
