---
ms.date: 01/08/2019
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques de fonctions personnalisées (aperçu)
localization_priority: Normal
ms.openlocfilehash: 4efcd0ba5efb0dc7450192694e8f0750de43b8a8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448608"
---
# <a name="custom-functions-best-practices-preview"></a>Meilleures pratiques de fonctions personnalisées (aperçu)

Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a>Résolution des problèmes

1. Si vous testez votre complément dans Office sur Windows, vous devez autoriser la ** [connexion d’exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) ** à résoudre les problèmes XML du fichier manifeste de votre complément, ainsi que plusieurs conditions d’installation et exécution. La connexion d’exécution écrit les`console.log`instructions vers un fichier journal pour vous aider à découvrir des problèmes.

2. Votre complément ne se charge pas si une ou plusieurs fonctions personnalisées sont en conflit avec les fonctions personnalisées d'un complément enregistré précédemment. Dans ce cas, vous pouvez supprimer le complément existant ou, si vous rencontrez cette erreur lors du développement d'un complément, vous pouvez spécifier un autre nom d'espace de noms dans votre manifeste.

3. Pour signaler des commentaires à l’équipe Excel des fonctions personnalisées sur cette méthode de résolution des problèmes, envoyez des commentaires à l’équipe. Pour ce faire, sélectionnez **Fichier | Commentaires | Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.

## <a name="associating-function-names-with-json-metadata"></a>Mappage des noms de fonction aux métadonnées JSON

Comme décrit dans l’article[vue d’ensemble de fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier de script (JavaScript ou machine à écrire) pour former une fonction complète. Pour qu'une fonction fonctionne correctement, vous devez associer l'ID à l'implémentation JavaScript. Vérifiez qu'il existe une association, sinon la fonction ne sera pas appelée.

L’exemple de code suivant montre comment procéder à cette association. L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.

* Utilisez uniquement des lettres majuscules d’une fonction `name` et `id` dans le fichier de métadonnées JSON. N’utilisez pas un mélange de cas ou uniquement des lettres minuscules. Si vous le faites, vous risquez de finir avec deux valeurs différentes uniquement par la casse ,cela entraînera un remplacement involontaire de vos fonctions. Par exemple, un objet de fonction à une `id` valeur**ajouter** peut être remplacé par déclaration plus loin dans le fichier d’objet de fonction avec une`id` valeur**AJOUTER**. De plus, la `name` propriété définit le nom de la fonction que les utilisateurs finaux verront dans Excel. Utiliser des lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonction intégrée sont en majuscules.

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.

* Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier. Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur. 

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

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript. Notez que les propriétés`id` et `name`sont en majuscules dans ce fichier. 

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

## <a name="declaring-optional-parameters"></a>Déclarer des paramètres facultatifs 

Dans Excel pour Windows (version 1812 ou version ultérieure), vous pouvez déclarer des paramètres facultatifs pour vos fonctions personnalisées. Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Par exemple, une fonction `FOO` avec un paramètre obligatoire appelé`parameter1` et un autre paramètre facultatif appelé `parameter2` apparaîtra sous la forme `=FOO(parameter1, [parameter2])` dans Excel.

Pour rendre un paramètre facultatif, ajouter `"optional": true` au paramètre dans le fichier de métadonnées JSON qui définit la fonction. L’exemple suivant montre comment cela peut se présenter pour la fonction `=ADD(first, second, [third])`. Vous pouvez remarquer que le paramètre facultatif `[third]` suit deux paramètres requis. Les paramètres obligatoires apparaissent en premier dans l’interface utilisateur formule d’Excel.

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
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
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le paramètre`zipCode` n’est pas défini, la valeur par défaut est définie sur 98052. Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a>Considérations supplémentaires

Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM. Sur Excel pour Windows, où les fonctions personnalisées utilisent l’[exécution JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.

## <a name="see-also"></a>Voir aussi

* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
