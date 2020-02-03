---
title: Utiliser l’API de boîte de dialogue Office dans vos compléments Office
description: Découvrir les notions de base relatives à la création d’une boîte de dialogue dans un complément Office
ms.date: 01/29/2020
localization_priority: Priority
ms.openlocfilehash: 13badafd0d3a6bb3fdf656b5caf93c9f514921d9
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650004"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="da490-103">Utiliser l’API de boîte de dialogue Office dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="da490-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="da490-104">Vous pouvez utiliser l’[API de dialogue Office](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="da490-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="da490-105">Cet article fournit des conseils concernant l’utilisation de l’API de dialogue dans votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="da490-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="da490-p102">Pour plus d’informations sur les compléments où l’API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l’API de dialogue](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). L’API de dialogue est actuellement prise en charge pour Word, Excel, PowerPoint et Outlook.</span><span class="sxs-lookup"><span data-stu-id="da490-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="da490-108">Un scénario principal pour l’API de dialogue consiste à activer l’authentification à l'aide d'une ressource telle que Google, Facebook, ou Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="da490-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="da490-109">Pour plus d’informations, voir [S’authentifier auprès de l'API de boîte de dialogue Office](auth-with-office-dialog-api.md) *une fois* que vous êtes familiarisé(e) avec cet article.</span><span class="sxs-lookup"><span data-stu-id="da490-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="da490-110">Envisagez d’ouvrir une boîte de dialogue à partir d’un volet Office, d’un complément de contenu ou d’un [complément de commande](../design/add-in-commands.md) pour effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="da490-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="da490-111">afficher les pages de connexion qui ne peuvent pas être ouvertes directement dans un volet Office ;</span><span class="sxs-lookup"><span data-stu-id="da490-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="da490-112">fournir davantage d’espace à l’écran, ou même un plein écran, pour certaines tâches exécutées dans votre complément ;</span><span class="sxs-lookup"><span data-stu-id="da490-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="da490-113">héberger une vidéo qui serait trop petite si elle était limitée à un volet Office.</span><span class="sxs-lookup"><span data-stu-id="da490-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="da490-114">Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige.</span><span class="sxs-lookup"><span data-stu-id="da490-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="da490-115">Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets.</span><span class="sxs-lookup"><span data-stu-id="da490-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="da490-116">Pour voir un exemple, consultez la rubrique relative à l’exemple [Complément Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="da490-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="da490-117">L’image suivante montre un exemple de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-117">The following image shows an example of a dialog box.</span></span>

![Commandes de complément](../images/auth-o-dialog-open.png)

<span data-ttu-id="da490-119">Notez que la boîte de dialogue s’ouvre toujours au centre de l’écran.</span><span class="sxs-lookup"><span data-stu-id="da490-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="da490-120">L’utilisateur peut la déplacer et la redimensionner.</span><span class="sxs-lookup"><span data-stu-id="da490-120">The user can move and resize it.</span></span> <span data-ttu-id="da490-121">La fenêtre est *non modale* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page dans le volet Office, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="da490-121">The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="da490-122">Ouvrir une boîte de dialogue à partir d’une page hôte</span><span class="sxs-lookup"><span data-stu-id="da490-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="da490-123">Les API JavaScript Office incluent un objet [Dialog](/javascript/api/office/office.dialog) et deux fonctions dans l’[espace de noms Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="da490-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="da490-124">Pour ouvrir une boîte de dialogue, généralement une page dans un volet des tâches, votre code appelle la méthode [displayDialogAsync](/javascript/api/office/office.ui) et lui transmet l’URL de la ressource que vous voulez ouvrir.</span><span class="sxs-lookup"><span data-stu-id="da490-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="da490-125">La page sur laquelle cette méthode est appelée est connue sous le nom de « page hôte ».</span><span class="sxs-lookup"><span data-stu-id="da490-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="da490-126">Par exemple, si vous appelez cette méthode dans le script sur index.html d'un volet de tâches, la page index.html correspond à la page hôte de la boîte de dialogue ouverte par la méthode.</span><span class="sxs-lookup"><span data-stu-id="da490-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="da490-127">La ressource ouverte dans la boîte de dialogue correspond généralement à une page, mais ce peut être une méthode du contrôleur dans une application MVC, un itinéraire, une méthode de service web ou toute autre ressource.</span><span class="sxs-lookup"><span data-stu-id="da490-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="da490-128">Dans cet article, les termes « page » ou « site web » font référence à la ressource dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="da490-129">Le code suivant est un exemple simple :</span><span class="sxs-lookup"><span data-stu-id="da490-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="da490-130">L’URL utilise le protocole HTTP**S**.</span><span class="sxs-lookup"><span data-stu-id="da490-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="da490-131">Ceci est obligatoire pour toutes les pages chargées dans une boîte de dialogue, pas seulement la première page chargée.</span><span class="sxs-lookup"><span data-stu-id="da490-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="da490-132">Le domaine de la boîte de dialogue est le même que celui de la page hôte, qui peut être la page d’un volet Office ou le [fichier de fonctions](/office/dev/add-ins/reference/manifest/functionfile) d’une commande de complément.</span><span class="sxs-lookup"><span data-stu-id="da490-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command.</span></span> <span data-ttu-id="da490-133">Obligatoire : la page, la méthode du contrôleur ou toute autre ressource qui est transmise à la méthode `displayDialogAsync` doit se trouver dans le même domaine que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="da490-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="da490-134">La page hôte et les ressources s'ouvrant dans la boîte de dialogue doivent avoir le même domaine complet.</span><span class="sxs-lookup"><span data-stu-id="da490-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="da490-135">Si vous tentez de transmettre `displayDialogAsync` à un sous-domaine du domaine du complément, cela ne fonctionnera pas.</span><span class="sxs-lookup"><span data-stu-id="da490-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="da490-136">Le domaine complet et tous les sous-domaines doivent être exactement les mêmes.</span><span class="sxs-lookup"><span data-stu-id="da490-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="da490-137">Une fois que la première page (ou toute autre ressource) est chargée, un utilisateur peut utiliser des liens ou une autre interface utilisateur pour accéder à n’importe quel site web (ou n’importe quelle autre ressource) qui utilise le protocole HTTPS.</span><span class="sxs-lookup"><span data-stu-id="da490-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="da490-138">Vous pouvez également concevoir la première page de façon à ce que l’utilisateur soit immédiatement redirigé vers un autre site.</span><span class="sxs-lookup"><span data-stu-id="da490-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="da490-139">Par défaut, la boîte de dialogue occupera 80 % de la hauteur et de la largeur de l’écran de l’appareil, mais vous pouvez définir des pourcentages différents en transmettant un objet de configuration à la méthode, comme indiqué dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="da490-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="da490-140">Pour voir un exemple de complément qui effectue ce type d’action, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="da490-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="da490-p112">Définissez les deux valeurs sur 100 % pour bénéficier d’une réelle d’expérience de plein écran. (Le maximum réel est de 99,5 %, et la fenêtre peut toujours être déplacée et redimensionnée.)</span><span class="sxs-lookup"><span data-stu-id="da490-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="da490-p113">Vous ne pouvez ouvrir qu’une seule boîte de dialogue à partir d’une fenêtre hôte. Toute tentative d’ouverture d’une autre boîte de dialogue génère une erreur. Par exemple, si un utilisateur ouvre une boîte de dialogue à partir d’un volet Office, il ne peut pas ouvrir une seconde boîte de dialogue à partir d’une autre page dans le volet Office. Toutefois, quand une boîte de dialogue est ouverte à partir d’une [commande de complément](../design/add-in-commands.md), la commande ouvre un nouveau fichier HTML (mais invisible) chaque fois qu’elle est sélectionnée. Cela crée une nouvelle fenêtre hôte (invisible), afin que chaque fenêtre de ce type puisse lancer sa propre boîte de dialogue. Pour plus d’informations, reportez-vous à [Erreurs provenant de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="da490-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="da490-149">Tirer parti d’une option de performances dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="da490-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="da490-150">La propriété `displayInIframe` est une propriété supplémentaire dans l’objet de configuration que vous pouvez transmettre à `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="da490-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="da490-151">Lorsque cette propriété est définie sur `true` et que le complément est en cours d’exécution dans un document ouvert dans Office sur le web, la boîte de dialogue s’ouvre sous la forme d’un iframe flottant et non d’une fenêtre indépendante ; elle s’ouvre ainsi plus rapidement.</span><span class="sxs-lookup"><span data-stu-id="da490-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="da490-152">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="da490-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="da490-153">La valeur par défaut est `false`, ce qui revient à omettre entièrement la propriété.</span><span class="sxs-lookup"><span data-stu-id="da490-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="da490-154">Si le complément n’est pas exécuté dans Office sur le Web, le `displayInIframe` est ignoré.</span><span class="sxs-lookup"><span data-stu-id="da490-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="da490-155">Vous ne devez **pas** utiliser `displayInIframe: true` si la boîte de dialogue redirige à un moment donné l’utilisateur vers une page qui ne peut pas être ouverte dans un IFrame.</span><span class="sxs-lookup"><span data-stu-id="da490-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="da490-156">Par exemple, les pages de connexion de nombreux services web connus, comme un compte Microsoft et Google, ne peuvent pas être ouvertes dans un IFrame.</span><span class="sxs-lookup"><span data-stu-id="da490-156">For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="da490-157">Envoi d’informations à la page hôte à partir de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="da490-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="da490-158">La boîte de dialogue ne peut pas communiquer avec la page hôte dans le volet Office, sauf si :</span><span class="sxs-lookup"><span data-stu-id="da490-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="da490-159">la page active dans la boîte de dialogue se trouve dans le même domaine que la page hôte ;</span><span class="sxs-lookup"><span data-stu-id="da490-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="da490-p117">la bibliothèque JavaScript Office est chargée dans la page. (Comme n’importe quelle page qui utilise la bibliothèque JavaScript Office, le script de la page doit attribuer une méthode à la propriété `Office.initialize`, bien qu’il puisse s’agir d’une méthode vide. Pour plus d’informations, voir [Initialisation de votre complément](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span><span class="sxs-lookup"><span data-stu-id="da490-p117">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="da490-163">Le code de la boîte de dialogue utilise la fonction [messageParent](/javascript/api/office/office.ui#messageparent-message-) pour envoyer une valeur booléenne ou un message de type chaîne à la page hôte.</span><span class="sxs-lookup"><span data-stu-id="da490-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="da490-164">La chaîne peut être un mot, une phrase, un blob XML, un JSON converti en chaîne ou un autre élément pouvant être sérialisé en chaîne.</span><span class="sxs-lookup"><span data-stu-id="da490-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="da490-165">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="da490-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="da490-p119">La fonction `messageParent` est l’une des deux *seules* API Office pouvant être appelées dans la boîte de dialogue. L’autre est `Office.context.requirements.isSetSupported`. Pour plus d’informations, consultez la rubrique relative à la [spécification d’hôtes Office et de conditions requises d’API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="da490-p119">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="da490-169">La fonction `messageParent` peut uniquement être appelée sur une page ayant le même domaine (y compris les mêmes protocole et port) que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="da490-169">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="da490-170">Dans l’exemple suivant, `googleProfile` est une version convertie en chaîne du profil Google de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="da490-170">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="da490-p120">La page hôte doit être configurée de façon à recevoir le message. Pour ce faire, ajoutez un paramètre de rappel à l’appel d’origine de `displayDialogAsync`. Le rappel attribue un gestionnaire à l’événement `DialogMessageReceived`. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="da490-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - <span data-ttu-id="da490-175">Office transmet un objet [AsyncResult](/javascript/api/office/office.asyncresult) au rappel.</span><span class="sxs-lookup"><span data-stu-id="da490-175">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="da490-176">Il représente le résultat de la tentative d’ouverture de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-176">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="da490-177">Il ne représente pas le résultat de tous les événements dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-177">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="da490-178">Pour plus d’informations sur cette distinction, consultez la [Gestion des erreurs et des événements](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="da490-178">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="da490-179">La propriété `value` de `asyncResult` est définie sur un objet [Dialog](/javascript/api/office/office.dialog), qui existe dans la page hôte, pas dans le contexte d’exécution de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-179">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="da490-p122">`processMessage` est la fonction qui gère l’événement. Vous pouvez lui donner le nom que vous souhaitez.</span><span class="sxs-lookup"><span data-stu-id="da490-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="da490-182">La variable `dialog` est déclarée avec une portée plus large que le rappel, car elle est également référencée dans `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="da490-182">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="da490-183">Voici un exemple simple de gestionnaire pour l’événement `DialogMessageReceived` :</span><span class="sxs-lookup"><span data-stu-id="da490-183">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="da490-184">Office transmet l’objet `arg` au gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="da490-184">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="da490-185">Sa propriété `message` est la valeur booléenne ou la chaîne envoyée par l’appel de `messageParent` dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-185">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="da490-186">Dans cet exemple, il s’agit d’une représentation convertie en chaîne du profil de l’utilisateur à partir d’un service tel qu’un compte Microsoft ou Google, qui est donc désérialisé en objet avec `JSON.parse`.</span><span class="sxs-lookup"><span data-stu-id="da490-186">In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="da490-p124">L’implémentation `showUserName` n’est pas visible. Elle peut afficher un message de bienvenue personnalisé dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="da490-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="da490-189">Lorsque l’intervention de l’utilisateur sur la boîte de dialogue est terminée, votre gestionnaire de messages doit fermer la boîte de dialogue, comme indiqué dans cet exemple.</span><span class="sxs-lookup"><span data-stu-id="da490-189">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="da490-190">L’objet `dialog` doit être le même que celui renvoyé par l’appel de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="da490-190">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="da490-191">L’appel de `dialog.close` indique à Office de fermer immédiatement la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="da490-191">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="da490-192">Pour voir un exemple de complément qui utilise ces techniques, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="da490-192">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="da490-p125">Si le complément a besoin d’ouvrir une autre page du volet Office après avoir reçu le message, vous pouvez utiliser la méthode `window.location.replace` (ou `window.location.href`) en tant que dernière ligne du gestionnaire. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="da490-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="da490-195">Pour voir un exemple de complément qui effectue ce type d’action, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="da490-195">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="da490-196">Messagerie conditionnelle</span><span class="sxs-lookup"><span data-stu-id="da490-196">Conditional messaging</span></span>

<span data-ttu-id="da490-p126">Étant donné que vous pouvez envoyer plusieurs appels `messageParent` à partir de la boîte de dialogue, mais que vous n’avez qu’un seul gestionnaire dans la page hôte pour l’événement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages. Par exemple, si la boîte de dialogue invite un utilisateur à se connecter à un fournisseur d’identité tel qu’un compte Microsoft ou Google, elle envoie le profil de l’utilisateur sous la forme d’un message. Si l’authentification échoue, la boîte de dialogue envoie des informations sur l’erreur à la page hôte, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="da490-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - <span data-ttu-id="da490-200">La variable `loginSuccess` serait initialisée en lisant la réponse HTTP à partir du fournisseur d’identité.</span><span class="sxs-lookup"><span data-stu-id="da490-200">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="da490-p127">L’implémentation des fonctions `getProfile` et `getError` n’est pas affichée. Chacune obtient des données à partir d’un paramètre de requête ou du corps de la réponse HTTP.</span><span class="sxs-lookup"><span data-stu-id="da490-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="da490-p128">Des objets anonymes de différents types sont envoyés selon que la connexion a réussi ou non. Tous deux ont une propriété `messageType`, mais un a une propriété `profile` et l’autre une propriété `error`.</span><span class="sxs-lookup"><span data-stu-id="da490-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="da490-p129">Le code du gestionnaire dans la page hôte utilise la valeur de la propriété `messageType` pour créer une branche comme le montre l’exemple suivant. Notez que la fonction `showUserName` est identique à celle de l’exemple précédent et que la fonction `showNotification` affiche l’erreur dans l’interface utilisateur de la page hôte.</span><span class="sxs-lookup"><span data-stu-id="da490-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> <span data-ttu-id="da490-207">L'implémentation `showNotification` n'est pas montrée dans l'exemple de code fourni par cet article.</span><span class="sxs-lookup"><span data-stu-id="da490-207">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="da490-208">Pour un exemple d'implémentation de cette fonction dans votre complément, voir [Exemple d'API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="da490-208">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="da490-209">Transmission d’informations à la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="da490-209">Pass information to the dialog box</span></span>

<span data-ttu-id="da490-p131">Parfois, la page hôte doit transmettre des informations à la boîte de dialogue. Pour ce faire, il existe deux moyens :</span><span class="sxs-lookup"><span data-stu-id="da490-p131">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="da490-212">ajouter des paramètres de requête à l’URL qui est transmise à `displayDialogAsync` ;</span><span class="sxs-lookup"><span data-stu-id="da490-212">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="da490-213">stocker les informations à un emplacement auquel à la fois la fenêtre hôte et la boîte de dialogue ont accès.</span><span class="sxs-lookup"><span data-stu-id="da490-213">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="da490-214">Les deux fenêtres ne partagent pas un stockage de session commun, mais *si elles ont le même domaine* (y compris le même numéro de port, le cas échéant), elles utilisent un [Stockage local](https://www.w3schools.com/html/html5_webstorage.asp) commun.\*</span><span class="sxs-lookup"><span data-stu-id="da490-214">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>

> [!NOTE]
> <span data-ttu-id="da490-215">\* Un bogue peut affecter votre stratégie de gestion des jetons.</span><span class="sxs-lookup"><span data-stu-id="da490-215">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="da490-216">Si le complément s’exécute dans **Office sur le web** dans le navigateur Safari ou Edge, la boîte de dialogue et le volet des tâches Office ne partagent pas le même stockage local. Il ne peut donc pas être utilisé pour communiquer entre eux.</span><span class="sxs-lookup"><span data-stu-id="da490-216">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="da490-217">Utilisation du stockage local</span><span class="sxs-lookup"><span data-stu-id="da490-217">Use local storage</span></span>

<span data-ttu-id="da490-218">Pour utiliser le stockage local, votre code appelle la méthode `setItem` de l’objet `window.localStorage` dans la page hôte avant l’appel de `displayDialogAsync`, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="da490-218">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="da490-219">Le code dans la boîte de dialogue qui lit l’élément lorsqu’il est nécessaire, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="da490-219">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="da490-220">Utiliser les paramètres de requête</span><span class="sxs-lookup"><span data-stu-id="da490-220">Use query parameters</span></span>

<span data-ttu-id="da490-221">L’exemple suivant montre comment transmettre des données à l’aide d’un paramètre de requête :</span><span class="sxs-lookup"><span data-stu-id="da490-221">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="da490-222">Pour obtenir un exemple qui utilise cette technique, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="da490-222">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="da490-223">Le code dans votre boîte de dialogue peut analyser l’URL et lire la valeur du paramètre.</span><span class="sxs-lookup"><span data-stu-id="da490-223">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="da490-p134">Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. (Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est pas ajouté à toutes les autres URL auxquelles la boîte de dialogue accède.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer entièrement, à l’avenir, donc votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue. Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.</span><span class="sxs-lookup"><span data-stu-id="da490-p134">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="da490-229">Fermeture de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="da490-229">Closing the dialog box</span></span>

<span data-ttu-id="da490-p135">Vous pouvez implémenter un bouton de fermeture dans la boîte de dialogue. Pour ce faire, le gestionnaire d’événements Click du bouton doit utiliser `messageParent` pour indiquer à la page hôte que vous avez cliqué sur le bouton. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="da490-p135">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="da490-233">Le gestionnaire de la page hôte pour `DialogMessageReceived` appelle `dialog.close`, comme dans cet exemple.</span><span class="sxs-lookup"><span data-stu-id="da490-233">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="da490-234">(consultez les exemples précédents qui montrent comment l’objet `dialog` est initialisé).</span><span class="sxs-lookup"><span data-stu-id="da490-234">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="da490-235">Même lorsque vous ne disposez pas de votre propre interface utilisateur de fermeture de boîte de dialogue, un utilisateur final peut fermer la boîte de dialogue en choisissant le **X** dans le coin supérieur droit.</span><span class="sxs-lookup"><span data-stu-id="da490-235">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="da490-236">Cette action déclenche l’événement `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="da490-236">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="da490-237">Si votre volet hôte a besoin de savoir quand cela se produit, il doit déclarer un gestionnaire pour cet événement.</span><span class="sxs-lookup"><span data-stu-id="da490-237">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="da490-238">Pour plus d’informations, consultez la section [Erreurs et événements dans la boîte de dialogue](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="da490-238">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="da490-239">Rubriques plus complexes et scénarios spéciaux</span><span class="sxs-lookup"><span data-stu-id="da490-239">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="da490-240">Utilisation d'un API de boîte de dialogue pour afficher une vidéo</span><span class="sxs-lookup"><span data-stu-id="da490-240">Use the Dialog API to show a video</span></span>

<span data-ttu-id="da490-241">Voir [Utiliser la boîte de dialogue Office pour afficher une vidéo](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="da490-241">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="da490-242">Utilisation des API de boîte de dialogue dans un flux d’authentification</span><span class="sxs-lookup"><span data-stu-id="da490-242">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="da490-243">Voir [Authentifier avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="da490-243">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="da490-244">Utilisation de l’API de boîte de dialogue Office avec des applications à page unique et routage côté client</span><span class="sxs-lookup"><span data-stu-id="da490-244">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="da490-245">Les authentifications par mot de passe (SPA) et le routage client doivent être gérés avec précaution lorsque vous utilisez l’API de boîte de dialogue Office.</span><span class="sxs-lookup"><span data-stu-id="da490-245">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="da490-246">Consultez les [Pratiques recommandées pour l’utilisation de l’API de boîte de dialogue Office dans une SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="da490-246">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="da490-247">Gestion d'erreurs et d'événements</span><span class="sxs-lookup"><span data-stu-id="da490-247">Error and event handling</span></span>

<span data-ttu-id="da490-248">Voir [Gestion des erreurs et des événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="da490-248">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="da490-249">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="da490-249">Next steps</span></span>

<span data-ttu-id="da490-250">Découvrez les pièges et pratiques recommandées pour l’API de boîte de dialogue Office dans les [Meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="da490-250">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>