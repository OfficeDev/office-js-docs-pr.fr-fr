---
title: Autres façons de transmettre des messages à une boîte de dialogue à partir de sa page hôte
description: Découvrez les solutions de contournement à utiliser lorsque la méthode messageChild n’est pas prise en charge.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f42a549a3c39866516cfd5395dd7589a890b0956
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889414"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Autres façons de transmettre des messages à une boîte de dialogue à partir de sa page hôte

La méthode recommandée pour transmettre des données et des messages d’une page parente à une boîte de dialogue enfant consiste à utiliser la `messageChild` méthode décrite dans [Utiliser l’API de boîte de dialogue Office dans vos compléments Office](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). Si votre complément s’exécute sur une plateforme ou un hôte qui ne prend pas en charge [l’ensemble de conditions requises DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), il existe deux autres façons de transmettre des informations à la boîte de dialogue.

- ajouter des paramètres de requête à l’URL qui est transmise à `displayDialogAsync` ;
- stocker les informations à un emplacement auquel à la fois la fenêtre hôte et la boîte de dialogue ont accès. Les deux fenêtres ne partagent pas de stockage de session commun (propriété [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), mais *si elles ont le même domaine* (y compris le numéro de port, le cas échéant), elles partagent un [stockage local](https://www.w3schools.com/html/html5_webstorage.asp) commun.\*

> [!NOTE]
> \* Un bogue peut affecter votre stratégie de gestion des jetons. Si le complément s’exécute dans **Office sur le web** dans le navigateur Safari ou Edge, la boîte de dialogue et le volet des tâches Office ne partagent pas le même stockage local. Il ne peut donc pas être utilisé pour communiquer entre eux.

## <a name="use-local-storage"></a>Utilisation du stockage local

Pour utiliser le stockage local, appelez la `setItem` méthode de l’objet `window.localStorage` dans la page hôte avant l’appel `displayDialogAsync` , comme dans l’exemple suivant.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Le code de la boîte de dialogue lit l’élément quand il est nécessaire, comme dans l’exemple suivant.

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Utiliser les paramètres de requête

L’exemple suivant montre comment transmettre des données à l’aide d’un paramètre de requête.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Pour obtenir un exemple qui utilise cette technique, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Le code dans votre boîte de dialogue peut analyser l’URL et lire la valeur du paramètre.

> [!IMPORTANT]
> Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. (Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est pas ajouté à toutes les autres URL auxquelles la boîte de dialogue accède.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer entièrement, à l’avenir, donc votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue (propriété [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.
