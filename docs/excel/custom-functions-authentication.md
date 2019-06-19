---
ms.date: 06/17/2019
description: Authentifiez les utilisateurs à l’aide de fonctions personnalisées dans Excel.
title: Authentification des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: 30ff1b91db8bf7f0183a44f1e7e078a6308c1351
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059740"
---
# <a name="authentication-for-custom-functions"></a>Authentification des fonctions personnalisées

Dans certains scénarios, votre fonction personnalisée doit authentifier l’utilisateur pour accéder aux ressources protégées. Bien que les fonctions personnalisées ne nécessitent pas de méthode spécifique d’authentification, sachez que les fonctions personnalisées s’exécutent dans un autre temps d’exécution, à partir du volet Office et d’autres éléments d’interface utilisateur de votre complément. Pour cette raison, vous devez transférer les données entre les deux exécutions à l’aide de l'objet`OfficeRuntime.storage` et de l’API de boîte de dialogue.

## <a name="officeruntimestorage-object"></a>Objet OfficeRuntime.storage

L’exécution des fonctions personnalisées n'a pas d’objet`localStorage`disponible dans la fenêtre globale, dans laquelle vous pouvez généralement stocker des données. Au lieu de cela, vous devez partager les données entre les fonctions personnalisées et les volets de tâches à l’aide de [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) pour configurer et obtenir les données.

De plus, l’utilisation de l'objet `storage`est avantageuse ; il utilise un environnement sandbox sécurisé pour que vos données ne soient pas accessibles aux autres compléments.

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous devez vous authentifier à partir du volet de tâches ou d’une fonction personnalisée, vérifiez `storage` pour voir si le jeton d’accès a déjà été acquis. Si ce n’est pas le cas, utilisez l’API de boîte de dialogue pour authentifier l’utilisateur, récupérer le jeton d’accès, puis stocker le jeton dans `storage`pour une utilisation ultérieure.

## <a name="dialog-api"></a>API de boîte de dialogue

Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter. Une fois qu’un utilisateur a entré ses informations d’identification, le jeton d’accès résultant peut être stocké dans `storage`.

> [!NOTE]
> Le runtime des fonctions personnalisées utilise un objet de boîte de dialogue qui est légèrement différent de l’objet de boîte de dialogue dans le moteur d’exécution du moteur d’exploration utilisé par les volets de tâches. Ils sont tous deux appelés «API de boîte de dialogue», mais utilisent `OfficeRuntime.Dialog`pour authentifier les utilisateurs dans le runtime de fonctions personnalisées.

Pour plus d’informations sur l’utilisation de l’objet `Dialog`, voir [boîte de dialogue fonctions personnalisées](/office/dev/add-ins/excel/custom-functions-dialog).

Lorsque vous envisagez l’intégralité du processus d’authentification, il peut être utile de considérer les éléments du volet de tâches et de l’interface utilisateur de votre complément, ainsi que les fonctions personnalisées de votre complément en tant qu’entités distinctes pouvant communiquer entre eux via`OfficeRuntime.storage`.

Le diagramme suivant décrit ce processus de base. Notez que la ligne pointillée indique qu’en effectuant des actions distinctes, les fonctions personnalisées et le volet de tâches de votre complément sont tous deux inclus dans votre complément.

1. Vous émettez un appel de fonction personnalisée à partir d’une cellule d’un classeur Excel.
2. La fonction personnalisée utilise `Dialog` pour transmettre vos informations d’identification d’utilisateur à un site Web.
3. Ce site Web renvoie ensuite un jeton d’accès à la fonction personnalisée.
4. Votre fonction personnalisée définit ensuite ce jeton d’accès sur `storage`.
5. Le volet de tâches de votre complément accède au jeton à partir de`storage`.

![Diagramme de la fonction personnalisée à l’aide de l’API de boîte de dialogue pour obtenir un jeton d’accès, puis partagez le jeton avec le volet de tâches via l’API OfficeRuntime.storage.](../images/authentication-diagram.png " Diagramme d’authentification.")

## <a name="storing-the-token"></a>Stockage du jeton

Les exemples suivants s’appliquent à partir de l’exemple de code[utilisation d’OfficeRuntime.storage dans les fonctions personnalisées](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Pour obtenir un exemple complet de partage de données entre les fonctions personnalisées et le volet de tâches, consultez cet exemple de code.

Si la fonction personnalisée s’authentifie, elle reçoit le jeton d’accès et doit la stocker dans `storage`. L’exemple de code suivant montre comment appeler la méthode`storage.setItem` pour stocker une valeur. La fonction `storeValue`est une fonction personnalisée qui, à titre d’exemple, stocke une valeur de l’utilisateur. Vous pouvez modifier cette valeur pour stocker les valeurs de jeton dont vous avez besoin.

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

CustomFunctions.associate("STOREVALUE", storeValue);
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
CustomFunctions.associate("GETTOKEN", receiveTokenFromCustomFunction);

```

## <a name="general-guidance"></a>Instructions générales

Les compléments Office sont basés sur le Web et vous pouvez utiliser n’importe quelle technique d’authentification Web. Il n’existe pas de modèle ou de méthode spécifique que vous devez suivre pour implémenter votre propre authentification avec des fonctions personnalisées. Vous pouvez consulter la documentation relative à différents modèles d’authentification, en commençant par[cet article sur l’autorisation d’accès via les services externes](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Évitez d’utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées :  

- `localStorage`: Les fonctions personnalisées n’ont pas accès à l’objet `window`global et n’ont par conséquent aucun accès aux données stockées dans`localStorage`.
- `Office.context.document.settings`: Cet emplacement n’est pas sécurisé et les informations peuvent être extraites par toute personne utilisant le complément.

## <a name="next-steps"></a>Étapes suivantes
En savoir plus sur[l’API de boîte de dialogue pour les fonctions personnalisées](custom-functions-dialog.md).

## <a name="see-also"></a>Voir aussi

* [Architecture des fonctions personnalisées](custom-functions-architecture.md)
* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
