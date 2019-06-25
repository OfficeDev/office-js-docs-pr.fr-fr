---
ms.date: 06/20/2019
description: Utilisez `OfficeRuntime.storage` pour enregistrer l’état des fonctions personnalisées.
title: Enregistrer et partager l’état des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: c6689393e5d118c779b7b261b0de04ead56aff83
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127834"
---
# <a name="save-and-share-state-in-custom-functions"></a>Enregistrer et partager l’état des fonctions personnalisées

Utilisez l’objet `OfficeRuntime.storage` pour enregistrer l’état lié aux fonctions personnalisées ou au volet Office dans votre complément. L’espace de stockage est limité à 10 Mo par domaine (avec possibilité de partage entre plusieurs compléments). Dans Excel sur Windows, l’objet `storage` correspond à un emplacement dans l’exécution de fonctions personnalisées, mais pour Excel sur le web et Mac, l’objet `storage` est le même que l’objet `localStorage` du navigateur.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Il existe plusieurs façons d’utiliser `storage` à des fins de gestion de l’état :

- Vous pouvez stocker les valeurs par défaut des fonctions personnalisées à utiliser lorsque vous êtes en mode hors connexion et dans l’impossibilité d’accéder à une ressource web.
- Vous pouvez enregistrer les valeurs des fonctions personnalisées à utiliser pour éviter d’appeler plusieurs fois une ressource web.
- Vous pouvez enregistrer des valeurs à partir de votre fonction personnalisée.
- Vous pouvez stocker les valeurs à partir de votre volet Office.

L’exemple de code suivant montre comment stocker un élément dans `storage` et le récupérer.

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

[Un exemple de code plus détaillé sur GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) montre comment transmettre ces informations au volet Office.

>[!NOTE]
> L’objet `storage` remplace l’objet de stockage précédent nommé `AsyncStorage`, et désormais déconseillé. Si vous utilisez l’objet `AsyncStorage` dans votre code de fonctions personnalisées en cours, mettez-le à jour de manière à utiliser l’objet `storage`.

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [générer automatiquement les métadonnées JSON pour vos fonctions personnalisées](custom-functions-json-autogeneration.md). 

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Débogage des fonctions personnalisées](custom-functions-debugging.md)
