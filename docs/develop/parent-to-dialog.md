---
title: Autres méthodes de transmission de messages à une boîte de dialogue à partir de sa page hôte
description: Découvrez les solutions de contournement à utiliser lorsque la méthode messageChild n’est pas prise en charge.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8da6bc3e1231bc6296a16fa153dc0e4ba1bd102b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349776"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Autres méthodes de transmission de messages à une boîte de dialogue à partir de sa page hôte

La méthode recommandée pour transmettre des données et des messages à partir d’une page parent à une boîte de dialogue enfant est d’utiliser la méthode décrite dans l’API de boîte de dialogue Office dans vos Office. `messageChild` [](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) Si votre module complémentaire est en cours d’exécution sur une plateforme ou un hôte qui ne prend pas en charge l’ensemble de conditions [requises DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)vous pouvez transmettre des informations à la boîte de dialogue de deux autres façons :

- ajouter des paramètres de requête à l’URL qui est transmise à `displayDialogAsync` ;
- stocker les informations à un emplacement auquel à la fois la fenêtre hôte et la boîte de dialogue ont accès. Les deux fenêtres ne partagent pas un stockage de session commun (propriété [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) mais si elles ont le même domaine *(y* compris le numéro de port, le cas Stockage [.](https://www.w3schools.com/html/html5_webstorage.asp)\*


> [!NOTE]
> \* Un bogue peut affecter votre stratégie de gestion des jetons. Si le complément s’exécute dans **Office sur le web** dans le navigateur Safari ou Edge, la boîte de dialogue et le volet des tâches Office ne partagent pas le même stockage local. Il ne peut donc pas être utilisé pour communiquer entre eux.

## <a name="use-local-storage"></a>Utilisation du stockage local

Pour utiliser le stockage local, appelez la méthode de l’objet dans la page hôte avant `setItem` `window.localStorage` `displayDialogAsync` l’appel, comme dans l’exemple suivant.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Le code de la boîte de dialogue lit l’élément lorsqu’il est nécessaire, comme dans l’exemple suivant.

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Utiliser les paramètres de requête

L’exemple suivant montre comment transmettre des données à l’aide d’un paramètre de requête.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Pour obtenir un exemple qui utilise cette technique, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Le code dans votre boîte de dialogue peut analyser l’URL et lire la valeur du paramètre.

> [!IMPORTANT]
> Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. (Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est pas ajouté à toutes les autres URL auxquelles la boîte de dialogue accède.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer entièrement, à l’avenir, donc votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue (propriété [Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.
