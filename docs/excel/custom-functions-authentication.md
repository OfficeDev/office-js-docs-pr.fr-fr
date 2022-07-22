---
title: Authentification pour les fonctions personnalisées sans runtime partagé
description: Authentifier les utilisateurs à l’aide de fonctions personnalisées qui n’utilisent pas de runtime partagé.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7ff7b1dca67e9e25f14ef07bd1c088608f254427
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958424"
---
# <a name="authentication-for-custom-functions-without-a-shared-runtime"></a>Authentification pour les fonctions personnalisées sans runtime partagé

Dans certains scénarios, une fonction personnalisée qui n’utilise pas de runtime partagé doit authentifier l’utilisateur pour accéder aux ressources protégées. Fonctions personnalisées qui n’utilisent pas d’exécution partagée dans un runtime JavaScript uniquement. Pour cette raison, si le complément a un volet Office, vous devez passer des données entre le runtime JavaScript uniquement et le runtime html utilisé par le volet Office. Pour ce faire, utilisez l’objet [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) et une API de dialogue spéciale.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objet OfficeRuntime.storage

Le runtime JavaScript uniquement n’a pas d’objet `localStorage` disponible dans la fenêtre globale, où vous stockez généralement des données. Au lieu de cela, votre code doit partager des données entre des fonctions personnalisées et des volets Office à l’aide de l’utilisation `OfficeRuntime.storage` pour définir et obtenir des données.

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous devez vous authentifier à partir d’un complément de fonction personnalisée qui n’utilise pas de runtime partagé, votre code doit vérifier `OfficeRuntime.storage` si le jeton d’accès a déjà été acquis. Si ce n’est pas le cas, utilisez [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton pour `OfficeRuntime.storage` une utilisation ultérieure.

## <a name="dialog-api"></a>API de boîte de dialogue

S’il n’existe pas de jeton, vous devez utiliser l’API `OfficeRuntime.dialog` pour demander à l’utilisateur de se connecter. Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké en tant qu’élément dans `OfficeRuntime.storage`.

> [!NOTE]
> Le runtime JavaScript uniquement utilise un objet de boîte de dialogue légèrement différent de l’objet de dialogue dans le runtime du moteur de navigateur utilisé par les volets office. Ils sont tous deux appelés « API de dialogue », mais utilisent [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) pour authentifier les utilisateurs dans le runtime JavaScript uniquement, *et non* [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)).

Le diagramme suivant décrit ce processus de base. La ligne en pointillés indique que les fonctions personnalisées et le volet Office de votre complément font tous les deux partie de votre complément dans son ensemble, bien qu’ils utilisent des runtimes distincts.

1. Vous émettez un appel de fonction personnalisée à partir d’une cellule d’un classeur Excel.
2. La fonction personnalisée utilise `OfficeRuntime.dialog` pour transmettre vos informations d’identification d’utilisateur à un site Web.
3. Ce site Web renvoie ensuite un jeton d’accès à la fonction personnalisée.
4. Votre fonction personnalisée définit ensuite ce jeton d’accès sur un élément dans le `OfficeRuntime.storage`.
5. Le volet de tâches de votre complément accède au jeton à partir de`OfficeRuntime.storage`.

![Diagramme de fonction personnalisée utilisant l’API de boîte de dialogue pour obtenir le jeton d’accès, puis partagez le jeton avec le volet Office via l’API OfficeRuntime.storage.](../images/authentication-diagram.png "Diagramme d’authentification.")

## <a name="storing-the-token"></a>Stockage du jeton

Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage). Reportez-vous à cet exemple de code pour obtenir un exemple complet de partage de données entre des fonctions personnalisées et le volet Office dans les compléments qui n’utilisent pas de runtime partagé.

Si la fonction personnalisée s’authentifie, elle reçoit le jeton d’accès et doit la stocker dans `OfficeRuntime.storage`. L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur. La `storeValue` fonction est une fonction personnalisée qui stocke une valeur de l’utilisateur. Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.

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

Lorsque le volet Office a besoin du jeton d’accès, il peut récupérer le jeton à partir de l’élément `OfficeRuntime.storage` . L’exemple de code suivant montre comment utiliser la méthode`storage.getItem` pour récupérer le jeton.

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  const key = "token";
  const tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>Instructions générales

Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web. Il n’existe pas de modèle ou de méthode spécifique que vous devez suivre pour implémenter votre propre authentification avec des fonctions personnalisées. Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](../develop/auth-external-add-ins.md).  

Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées :

- `localStorage`: les fonctions personnalisées qui n’utilisent pas de runtime partagé n’ont pas accès à l’objet global `window` et n’ont donc pas accès aux données stockées dans `localStorage`.
- `Office.context.document.settings`: cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.

## <a name="dialog-box-api-example"></a>Exemple d’API de boîte de dialogue

Dans l’exemple de code suivant, la fonction `getTokenViaDialog` utilise la `OfficeRuntime.displayWebDialog` fonction pour afficher une boîte de dialogue. Cet exemple est fourni pour montrer les fonctionnalités de la méthode, et non pour montrer comment s’authentifier.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this isn't a sufficient example of authentication but is intended to show the capabilities of the displayWebDialog method.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      const intervalId = setInterval(function () {
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

Découvrez comment [déboguer des fonctions personnalisées](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

- [Runtime JavaScript uniquement pour les fonctions personnalisées](custom-functions-runtime.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
