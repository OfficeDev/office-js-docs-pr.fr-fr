---
ms.date: 10/24/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions Excel personnalisées.
title: Meilleures pratiques de fonctions personnalisées
ms.openlocfilehash: 0408318227e1f89726ed7c0e4dfbb8e6340abef4
ms.sourcegitcommit: 52d18dd8a60e0cec1938394669d577570700e61e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/26/2018
ms.locfileid: "25797398"
---
# <a name="custom-functions-best-practices-preview"></a>Meilleures pratiques de fonctions personnalisées (aperçu)

Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

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

## <a name="troubleshooting"></a>Résolution des problèmes

Si vous testez votre complément dans Office sur Windows, vous devez autoriser la ** [connexion d’exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) ** à résoudre les problèmes XML du fichier manifeste de votre complément, ainsi que plusieurs conditions d’installation et exécution. La connexion d’exécution écrit les`console.log`instructions vers un fichier journal pour vous aider à découvrir des problèmes.

Pour signaler des commentaires à l’équipe Excel des fonctions personnalisées sur cette méthode de résolution des problèmes, envoyez des commentaires à l’équipe. Pour ce faire, sélectionnez **Fichier | Commentaires | Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez. 

## <a name="debugging"></a>Débogage

Pour l’instant, la méthode optimale pour le débogage de fonctions personnalisées Excel consiste à [charger](../testing/sideload-office-add-ins-for-testing.md) votre complément au sein d’**Excel Online**. Vous pouvez ensuite déboguer vos fonctions personnalisées à l’aide de l’ [outil natif F12 de débogage de votre navigateur](../testing/debug-add-ins-in-office-online.md) en combinaison avec les techniques suivantes :

- Utilisez les`console.log` instructions au sein de votre code de fonctions personnalisées pour envoyer la sortie à la console en temps réel.

- Utilisez les `debugger;` instructions au sein de votre code de fonctions personnalisées pour spécifier les points d'arrêt où l’exécution sera suspendue lorsque la fenêtre F12 est ouverte. Par exemple, si la fonction suivante s’exécute lorsque la fenêtre F12 est ouverte, l’exécution sera suspendue sur la`debugger;` déclaration, vous permettant d’inspecter manuellement les valeurs de paramètres avant le retour de la fonction. L’`debugger;` instruction n’a aucun effet dans Excel Online lorsque la fenêtre F12 n’est pas ouverte. Pour l’instant, l’`debugger;` instruction n’a aucun effet dans Excel pour Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

Si votre complément ne parvient pas à s’enregistrer, [vérifier que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web hébergeant votre application complément.

## <a name="mapping-function-names-to-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux. Par ailleurs, dans le fichier JavaScript qui définit vos fonctions personnalisées, vous devez fournir des informations pour spécifier quel objet fonction dans le fichier de métadonnées JSON correspond à chaque fonction personnalisée dans le fichier JavaScript.

Par exemple, l’exemple de code suivant définit la fonction personnalisée `add` et puis indique que la fonction `add` correspond à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **Ajouter**.

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.

* Dans le fichier JavaScript, spécifiez les noms de fonction dans camelCase. Par exemple, le nom de fonction `addTenToInput` écrit dans camelCase : le premier mot dans le nom commence par une lettre en minuscule et chaque mot suivant dans le nom commence par une lettre en majuscule.

* Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque `name` propriété en majuscules. La `name` propriété définit le nom de la fonction que les utilisateurs finaux verront dans Excel. Utiliser des lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonction intégrée sont en majuscules.

* Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque `id` propriété en majuscules. Cette opération souligne quelle partie de l’`CustomFunctionMappings` instruction dans votre code JavaScript correspond à la `id` propriété dans le fichier métadonnées JSON (à condition que votre nom de fonction utilise camelCase, comme recommandé précédemment).

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques. 

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier. Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit avoir la même`id` valeur. En outre, n’indiquez pas deux `id` valeurs dans le fichier de métadonnées qui diffèrent uniquement par la casse. Par exemple, ne définissez pas un objet fonction avec une `id` valeur **ajouter** et un autre objet fonction avec une`id` valeur de **AJOUTER**.

* Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante. Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.

* Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées dans le même emplacement. Par exemple, le code suivant définit deux fonctions personnalisées et indique ensuite les informations de mappage pour les deux fonctions.

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

    L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.

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

Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM. Sur Excel pour Windows, où les fonctions personnalisées utilisent l’[exécution JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
