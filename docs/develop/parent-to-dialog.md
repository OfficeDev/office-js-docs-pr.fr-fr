---
title: Autres méthodes de transmission de messages à une boîte de dialogue à partir de sa page hôte
description: Découvrez les solutions de contournement à utiliser lorsque la méthode messageChild n’est pas prise en charge.
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: b516896d28979f439f3065f9ff036ff21c2c0997
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293176"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="2bbd5-103">Autres méthodes de transmission de messages à une boîte de dialogue à partir de sa page hôte</span><span class="sxs-lookup"><span data-stu-id="2bbd5-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="2bbd5-104">Pour transmettre les données et les messages d’une page parent à une boîte de dialogue enfant, il est recommandé d' `messageChild` utiliser la méthode décrite dans la rubrique [use the Office Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). Si votre complément est exécuté sur une plateforme ou un hôte qui ne prend pas en charge l' [ensemble de conditions requises DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), il existe deux autres façons de transmettre des informations à la boîte de dialogue :</span><span class="sxs-lookup"><span data-stu-id="2bbd5-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="2bbd5-105">ajouter des paramètres de requête à l’URL qui est transmise à `displayDialogAsync` ;</span><span class="sxs-lookup"><span data-stu-id="2bbd5-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="2bbd5-106">stocker les informations à un emplacement auquel à la fois la fenêtre hôte et la boîte de dialogue ont accès.</span><span class="sxs-lookup"><span data-stu-id="2bbd5-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="2bbd5-107">Les deux fenêtres ne partagent pas un stockage de session commun, mais *si elles ont le même domaine* (y compris le même numéro de port, le cas échéant), elles utilisent un [Stockage local](https://www.w3schools.com/html/html5_webstorage.asp) commun.\*</span><span class="sxs-lookup"><span data-stu-id="2bbd5-107">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="2bbd5-108">\* Un bogue peut affecter votre stratégie de gestion des jetons.</span><span class="sxs-lookup"><span data-stu-id="2bbd5-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="2bbd5-109">Si le complément s’exécute dans **Office sur le web** dans le navigateur Safari ou Edge, la boîte de dialogue et le volet des tâches Office ne partagent pas le même stockage local. Il ne peut donc pas être utilisé pour communiquer entre eux.</span><span class="sxs-lookup"><span data-stu-id="2bbd5-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="2bbd5-110">Utilisation du stockage local</span><span class="sxs-lookup"><span data-stu-id="2bbd5-110">Use local storage</span></span>

<span data-ttu-id="2bbd5-111">Pour utiliser le stockage local, appelez la `setItem` méthode de l' `window.localStorage` objet dans la page hôte avant l' `displayDialogAsync` appel, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="2bbd5-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="2bbd5-112">Le code dans la boîte de dialogue qui lit l’élément lorsqu’il est nécessaire, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="2bbd5-112">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="2bbd5-113">Utiliser les paramètres de requête</span><span class="sxs-lookup"><span data-stu-id="2bbd5-113">Use query parameters</span></span>

<span data-ttu-id="2bbd5-114">L’exemple suivant montre comment transmettre des données à l’aide d’un paramètre de requête :</span><span class="sxs-lookup"><span data-stu-id="2bbd5-114">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="2bbd5-115">Pour obtenir un exemple qui utilise cette technique, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="2bbd5-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="2bbd5-116">Le code dans votre boîte de dialogue peut analyser l’URL et lire la valeur du paramètre.</span><span class="sxs-lookup"><span data-stu-id="2bbd5-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2bbd5-p103">Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. (Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est pas ajouté à toutes les autres URL auxquelles la boîte de dialogue accède.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer entièrement, à l’avenir, donc votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue. Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.</span><span class="sxs-lookup"><span data-stu-id="2bbd5-p103">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>
