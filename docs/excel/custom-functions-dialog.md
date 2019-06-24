---
ms.date: 06/18/2019
description: Créez une boîte de dialogue via des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Afficher la boîte de dialogue d’une fonction personnalisée
localization_priority: Priority
ms.openlocfilehash: e513aedd46f129371a5c858e84f7e230f8d7ae11
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127925"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a>Afficher la boîte de dialogue d’une fonction personnalisée

Si votre fonction personnalisée doit interagir avec l’utilisateur, vous pouvez créer une boîte de dialogue à l’aide de l’[objet`Office.Dialog`](/javascript/api/office-runtime/officeruntime.dialog?view=office-js). Un scénario classique pour l’utilisation de la boîte de dialogue consiste à authentifier un utilisateur afin que votre fonction personnalisée puisse accéder à un service web. Pour plus d’informations sur l’authentification de fonctions personnalisées, voir[authentification des fonctions personnalisées](./custom-functions-authentication.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> L’objet `Office.Dialog` fait partie de l’exécution de fonctions personnalisées. Les volets Office n’utilisent pas l’objet `Dialog`. Pour créer une boîte de dialogue à partir d’un volet de tâches, consultez [API de boîte de dialogue](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## <a name="dialog-box-api-example"></a>Exemple d’API de boîte de dialogue

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la fonction `Dialog` de l’API `displayWebDialogOptions` pour afficher une boîte de dialogue.

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      Office.displayWebDialogOptions(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md).

## <a name="see-also"></a>Voir aussi

* [Authentification des fonctions personnalisées](custom-functions-authentication.md)
* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
