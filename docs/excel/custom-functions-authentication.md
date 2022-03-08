---
ms.date: 05/17/2020
description: Authentifier les utilisateurs à l’aide de fonctions Excel qui n’utilisent pas le volet Des tâches.
title: Authentification pour les fonctions personnalisées sans interface utilisateur
ms.localizationpriority: medium
ms.openlocfilehash: 946800cf884f903e0794702d32ffb7e1075e1ca3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340273"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Authentification pour les fonctions personnalisées sans interface utilisateur

Dans certains scénarios, votre fonction personnalisée qui n’utilise pas de volet de tâches ou d’autres éléments d’interface utilisateur (fonction personnalisée sans interface utilisateur) devra authentifier l’utilisateur pour accéder aux ressources protégées. N’ignorez pas que les fonctions personnalisées sans interface utilisateur s’exécutent dans un runtime JavaScript uniquement. Pour cette raison, vous devez transmettre des données entre le runtime JavaScript uniquement et le runtime `OfficeRuntime.storage` de moteur de navigateur standard utilisé par la plupart des applications à l’aide de l’objet et de l’API de dialogue.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objet OfficeRuntime.storage

Le runtime JavaScript uniquement utilisé par les fonctions personnalisées sans interface utilisateur n’a `localStorage` pas d’objet disponible dans la fenêtre globale, où vous stockez généralement des données. Au lieu de cela, vous devez partager des données entre des fonctions personnalisées sans interface utilisateur et des volets Office à l’aide [d’OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) pour définir et obtenir des données.

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous devez vous authentifier à partir d’une fonction personnalisée sans interface utilisateur, `storage` vérifiez si le jeton d’accès a déjà été acquis. Si ce n’est pas le cas, utilisez l’API de boîte de dialogue pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton dans `storage`pour une utilisation ultérieure.

## <a name="dialog-api"></a>API de boîte de dialogue

Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter. Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké dans `storage`.

> [!NOTE]
> Le runtime JavaScript uniquement utilise un objet Dialog légèrement différent de l’objet Dialog dans le runtime du moteur de navigateur utilisé par les volets Des tâches. Ils sont tous deux appelés « API `OfficeRuntime.Dialog` de boîte de dialogue », mais utilisés pour authentifier les utilisateurs dans le runtime JavaScript uniquement.

Le diagramme suivant décrit ce processus de base. La ligne en pointillés indique que les fonctions personnalisées sans interface utilisateur et le volet Des tâches de votre add-in font tous deux partie de votre module, bien qu’elles utilisent des runtimes distincts.

1. Vous émettrez un appel de fonction personnalisée sans interface utilisateur à partir d’une cellule dans un Excel de travail.
2. La fonction personnalisée sans interface utilisateur utilise `Dialog` pour transmettre vos informations d’identification utilisateur à un site web.
3. Ce site web renvoie ensuite un jeton d’accès à la fonction personnalisée sans interface utilisateur.
4. Votre fonction personnalisée sans interface utilisateur définit ensuite ce jeton d’accès sur le `storage`.
5. Le volet de tâches de votre complément accède au jeton à partir de`storage`.

![Diagramme de la fonction personnalisée à l’aide de l’API de boîte de dialogue pour obtenir un jeton d’accès, puis partagez le jeton avec le volet Office via l’API OfficeRuntime.storage.](../images/authentication-diagram.png "Diagramme d’authentification.")

## <a name="storing-the-token"></a>Stockage du jeton

Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage). Reportez-vous à cet exemple de code pour obtenir un exemple complet de partage de données entre des fonctions personnalisées sans interface utilisateur et le volet Des tâches.

Si la fonction personnalisée sans interface utilisateur s’authentifiera, elle reçoit le jeton d’accès et devra le stocker dans `storage`. L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur. La `storeValue` fonction est une fonction personnalisée sans interface utilisateur qui, par exemple, stocke une valeur de l’utilisateur. Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.

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

Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web. Il n’existe aucun modèle ou méthode particulier que vous devez suivre pour implémenter votre propre authentification avec des fonctions personnalisées sans interface utilisateur. Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](../develop/auth-external-add-ins.md).  

Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées : .

- `localStorage`: les fonctions personnalisées sans interface utilisateur n’ont pas accès à l’objet global `window` et, par conséquent, n’ont pas accès aux données stockées dans `localStorage`.
- `Office.context.document.settings`: Cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.

## <a name="dialog-box-api-example"></a>Exemple d’API de boîte de dialogue

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la `Dialog` fonction de l’API `displayWebDialogOptions` pour afficher une boîte de dialogue. Cet exemple est fourni pour montrer les fonctionnalités de l’objet `Dialog` , et non pour montrer comment s’authentifier.

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

Découvrez comment [déboguer des fonctions personnalisées sans interface utilisateur](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

* [Runtime pour les fonctions personnalisées sans interface Excel’interface utilisateur](custom-functions-runtime.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)