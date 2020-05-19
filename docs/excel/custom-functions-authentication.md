---
ms.date: 05/17/2020
description: Authentifier les utilisateurs à l’aide de fonctions personnalisées dans Excel qui n’utilisent pas le volet Office.
title: Authentification pour les fonctions personnalisées sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: 93073fb23f3f4d30c36faf4927a3aebdafbc887d
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278377"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Authentification pour les fonctions personnalisées sans interface utilisateur

Dans certains scénarios, votre fonction personnalisée qui n’utilise pas de volet de tâches ou d’autres éléments de l’interface utilisateur (fonction personnalisée sans interface utilisateur) doit authentifier l’utilisateur afin d’accéder aux ressources protégées. N’oubliez pas que les fonctions personnalisées sans interface utilisateur s’exécutent dans un Runtime JavaScript uniquement. Pour cette raison, vous devez transmettre les données entre le runtime JavaScript uniquement et le runtime du moteur de navigateur standard utilisé par la plupart des compléments à l’aide de l' `OfficeRuntime.storage` objet et de l’API de dialogue.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objet OfficeRuntime.storage

Le runtime JavaScript uniquement utilisé par des fonctions personnalisées sans interface utilisateur ne dispose pas d’un `localStorage` objet disponible dans la fenêtre globale, dans laquelle vous stockez généralement les données. Au lieu de cela, vous devez partager des données entre des fonctions personnalisées sans interface utilisateur et des volets de tâches à l’aide de [OfficeRuntime. Storage](/javascript/api/office-runtime/officeruntime.storage) pour définir et obtenir des données.

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous devez vous authentifier à partir d’une fonction personnalisée sans interface utilisateur, vérifiez `storage` si le jeton d’accès a déjà été acquis. Si ce n’est pas le cas, utilisez l’API de boîte de dialogue pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton dans `storage`pour une utilisation ultérieure.

## <a name="dialog-api"></a>API de boîte de dialogue

Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter. Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké dans `storage`.

> [!NOTE]
> Le runtime JavaScript uniquement utilise un objet Dialog qui est légèrement différent de l’objet Dialog dans le runtime du moteur du navigateur utilisé par les volets des tâches. Ils sont tous deux appelés « API de dialogue », mais utilisent `OfficeRuntime.Dialog` pour authentifier les utilisateurs dans le runtime JavaScript uniquement.

Le diagramme suivant décrit ce processus de base. La ligne pointillée indique que les fonctions personnalisées sans interface utilisateur et le volet Office de votre complément font partie de votre complément dans son intégralité, même s’ils utilisent des runtimes distincts.

1. Vous émettez un appel de fonction personnalisée sans interface utilisateur à partir d’une cellule dans un classeur Excel.
2. La fonction personnalisée sans interface utilisateur utilise `Dialog` pour transmettre vos informations d’identification d’utilisateur à un site Web.
3. Ce site Web renvoie ensuite un jeton d’accès à la fonction personnalisée sans interface utilisateur.
4. Votre fonction personnalisée sans interface utilisateur définit ensuite le jeton d’accès sur `storage` .
5. Le volet de tâches de votre complément accède au jeton à partir de`storage`.

![Diagramme de la fonction personnalisée à l’aide de l’API de boîte de dialogue pour obtenir le jeton d’accès, puis partager le jeton avec le volet de tâches via l’API OfficeRuntime. Storage.](../images/authentication-diagram.png "Diagramme d’authentification.")

## <a name="storing-the-token"></a>Stockage du jeton

Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Pour obtenir un exemple complet de partage de données entre des fonctions personnalisées sans interface utilisateur et le volet Office, reportez-vous à cet exemple de code.

Si la fonction personnalisée sans interface utilisateur s’authentifie, elle reçoit le jeton d’accès et doit le stocker dans `storage` . L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur. La `storeValue` fonction est une fonction personnalisée sans interface utilisateur qui, par exemple, stocke une valeur de l’utilisateur. Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

Lorsque le volet de tâches a besoin du jeton d’accès, il peut récupérer le jeton à partir de `storage`. L’exemple de code suivant montre comment utiliser la méthode`storage.getItem` pour récupérer le jeton.

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>Instructions générales

Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web. Il n’existe pas de modèle ni de méthode particulier à respecter pour implémenter votre propre authentification avec des fonctions personnalisées sans interface utilisateur. Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](../develop/auth-external-add-ins.md).  

Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées :  

- `localStorage`: Les fonctions personnalisées sans interface utilisateur n’ont pas accès à l' `window` objet global et, par conséquent, n’ont pas accès aux données stockées dans `localStorage` .
- `Office.context.document.settings`: Cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.

## <a name="dialog-box-api-example"></a>Exemple d’API de boîte de dialogue

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la `Dialog` fonction de l’API `displayWebDialogOptions` pour afficher une boîte de dialogue. Cet exemple est fourni pour afficher les fonctionnalités de l' `Dialog` objet, ne pas montrer comment s’authentifier.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
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
      OfficeRuntime.displayWebDialog(url, {
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
Découvrez comment [Déboguer des fonctions personnalisées sans interface utilisateur](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

* [Runtime pour les fonctions personnalisées Excel sans interface utilisateur](custom-functions-runtime.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
