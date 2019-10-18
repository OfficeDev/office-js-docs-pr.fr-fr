---
ms.date: 07/10/2019
description: Utilisez `OfficeRuntime.storage` pour enregistrer l’état des fonctions personnalisées.
title: Enregistrer et partager l’état des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: a1b70433ef0c00d507175b32fc12603ff3de1e3f
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771588"
---
# <a name="save-and-share-state-in-custom-functions"></a><span data-ttu-id="9f7e4-103">Enregistrer et partager l’état des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="9f7e4-103">Save and share state in custom functions</span></span>

<span data-ttu-id="9f7e4-104">Utilisez l’objet `OfficeRuntime.storage` pour enregistrer l’état lié aux fonctions personnalisées ou au volet Office dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-104">Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in.</span></span> <span data-ttu-id="9f7e4-105">L’espace de stockage est limité à 10 Mo par domaine (avec possibilité de partage entre plusieurs compléments).</span><span class="sxs-lookup"><span data-stu-id="9f7e4-105">Storage is limited to 10 MB per domain (which may be shared across multiple add-ins).</span></span> <span data-ttu-id="9f7e4-106">Dans Excel sur Windows, l’objet `storage` correspond à un emplacement dans l’exécution de fonctions personnalisées, mais pour Excel sur le web et Mac, l’objet `storage` est le même que l’objet `localStorage` du navigateur.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-106">In Excel on Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="9f7e4-107">Il existe plusieurs façons d’utiliser `storage` à des fins de gestion de l’état :</span><span class="sxs-lookup"><span data-stu-id="9f7e4-107">There are multiple ways to use `storage` for state management:</span></span>

- <span data-ttu-id="9f7e4-108">Vous pouvez stocker les valeurs par défaut des fonctions personnalisées à utiliser lorsque vous êtes en mode hors connexion et dans l’impossibilité d’accéder à une ressource web.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-108">You can store default values for custom functions to use when you are offline and unable to reach a web resource.</span></span>
- <span data-ttu-id="9f7e4-109">Vous pouvez enregistrer les valeurs des fonctions personnalisées à utiliser pour éviter d’appeler plusieurs fois une ressource web.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-109">You can save values for custom functions to use to avoid making additional calls to a web resource.</span></span>
- <span data-ttu-id="9f7e4-110">Vous pouvez enregistrer des valeurs à partir de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-110">You can save values from your custom function.</span></span>
- <span data-ttu-id="9f7e4-111">Vous pouvez stocker les valeurs à partir de votre volet Office.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-111">You can store values from your task pane.</span></span>

<span data-ttu-id="9f7e4-112">L’exemple de code suivant montre comment stocker un élément dans `storage` et le récupérer.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-112">The following code sample illustrates how to store an item into `storage` and retrieve it.</span></span>

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
```

<span data-ttu-id="9f7e4-113">[Un exemple de code plus détaillé sur GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) montre comment transmettre ces informations au volet Office.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-113">[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.</span></span>

>[!NOTE]
> <span data-ttu-id="9f7e4-114">L’objet `storage` remplace l’objet de stockage précédent nommé `AsyncStorage`, et désormais déconseillé.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-114">The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated.</span></span> <span data-ttu-id="9f7e4-115">Si vous utilisez l’objet `AsyncStorage` dans votre code de fonctions personnalisées en cours, mettez-le à jour de manière à utiliser l’objet `storage`.</span><span class="sxs-lookup"><span data-stu-id="9f7e4-115">If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.</span></span>

## <a name="next-steps"></a><span data-ttu-id="9f7e4-116">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="9f7e4-116">Next steps</span></span>
<span data-ttu-id="9f7e4-117">Découvrez comment [générer automatiquement les métadonnées JSON pour vos fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="9f7e4-117">Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).</span></span> 

## <a name="see-also"></a><span data-ttu-id="9f7e4-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9f7e4-118">See also</span></span>

* [<span data-ttu-id="9f7e4-119">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="9f7e4-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="9f7e4-120">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="9f7e4-120">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="9f7e4-121">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="9f7e4-121">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="9f7e4-122">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="9f7e4-122">Custom functions debugging</span></span>](custom-functions-debugging.md)
