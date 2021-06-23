---
title: Utiliser l’API de boîte de dialogue Office dans vos compléments Office
description: Découvrez les principes de base de la création d’une boîte de dialogue dans un Office de recherche.
ms.date: 01/28/2021
localization_priority: Normal
ms.openlocfilehash: 210b12f826e0d0d360163ee7663d6afca740a24d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076103"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="1e4fb-103">Utiliser l’API de boîte de dialogue Office dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="1e4fb-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="1e4fb-104">Vous pouvez utiliser l’[API de dialogue Office](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="1e4fb-105">Cet article fournit des conseils concernant l’utilisation de l’API de dialogue dans votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1e4fb-106">Pour plus d’informations sur les compléments où l’API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l’API de dialogue](../reference/requirement-sets/dialog-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-106">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md).</span></span> <span data-ttu-id="1e4fb-107">L’API de dialogue est actuellement prise en charge pour Excel, PowerPoint et Word.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-107">The Dialog API is currently supported for Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="1e4fb-108">Outlook prise en charge est incluse dans différents ensembles de conditions requises de boîte aux lettres, consultez la référence &mdash; d’API pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-108">Outlook support is included across various Mailbox requirement sets&mdash;see the API reference for more details.</span></span>

<span data-ttu-id="1e4fb-109">Un scénario principal pour l’API de dialogue consiste à activer l’authentification à l'aide d'une ressource telle que Google, Facebook, ou Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-109">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="1e4fb-110">Pour plus d’informations, voir [S’authentifier auprès de l'API de boîte de dialogue Office](auth-with-office-dialog-api.md) *une fois* que vous êtes familiarisé(e) avec cet article.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-110">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="1e4fb-111">Envisagez d’ouvrir une boîte de dialogue à partir d’un volet Office, d’un complément de contenu ou d’un [complément de commande](../design/add-in-commands.md) pour effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="1e4fb-112">afficher les pages de connexion qui ne peuvent pas être ouvertes directement dans un volet Office ;</span><span class="sxs-lookup"><span data-stu-id="1e4fb-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="1e4fb-113">fournir davantage d’espace à l’écran, ou même un plein écran, pour certaines tâches exécutées dans votre complément ;</span><span class="sxs-lookup"><span data-stu-id="1e4fb-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="1e4fb-114">héberger une vidéo qui serait trop petite si elle était limitée à un volet Office.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="1e4fb-115">Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-115">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="1e4fb-116">Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-116">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="1e4fb-117">Pour obtenir un exemple de volet De tâches à onglets, voir l’exemple Excel de [salestracker JavaScript pour le](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) module de développement.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-117">For an example of a tabbed task pane, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="1e4fb-118">L’image suivante montre un exemple de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-118">The following image shows an example of a dialog box.</span></span>

![Capture d’écran montrant la boîte de dialogue avec 3 options de personnalisation affichées devant Word.](../images/auth-o-dialog-open.png)

<span data-ttu-id="1e4fb-120">Notez que la boîte de dialogue s’ouvre toujours au centre de l’écran.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-120">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="1e4fb-121">L’utilisateur peut la déplacer et la redimensionner.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-121">The user can move and resize it.</span></span> <span data-ttu-id="1e4fb-122">La fenêtre *n’est* pasmodale : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office et avec la page dans le volet Des tâches, s’il en existe un.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-122">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="1e4fb-123">Ouvrir une boîte de dialogue à partir d’une page hôte</span><span class="sxs-lookup"><span data-stu-id="1e4fb-123">Open a dialog box from a host page</span></span>

<span data-ttu-id="1e4fb-124">Les API JavaScript Office incluent un objet [Dialog](/javascript/api/office/office.dialog) et deux fonctions dans l’[espace de noms Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-124">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="1e4fb-125">Pour ouvrir une boîte de dialogue, généralement une page dans un volet des tâches, votre code appelle la méthode [displayDialogAsync](/javascript/api/office/office.ui) et lui transmet l’URL de la ressource que vous voulez ouvrir.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-125">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="1e4fb-126">La page sur laquelle cette méthode est appelée est connue sous le nom de « page hôte ».</span><span class="sxs-lookup"><span data-stu-id="1e4fb-126">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="1e4fb-127">Par exemple, si vous appelez cette méthode dans le script sur index.html d'un volet de tâches, la page index.html correspond à la page hôte de la boîte de dialogue ouverte par la méthode.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-127">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="1e4fb-128">La ressource ouverte dans la boîte de dialogue correspond généralement à une page, mais ce peut être une méthode du contrôleur dans une application MVC, un itinéraire, une méthode de service web ou toute autre ressource.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-128">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="1e4fb-129">Dans cet article, les termes « page » ou « site web » font référence à la ressource dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-129">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="1e4fb-130">Le code suivant est un exemple simple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-130">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="1e4fb-p108">L’URL utilise le protocole HTTP **S**. Ceci est obligatoire pour toutes les pages chargées dans une boîte de dialogue, pas seulement la première page chargée.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p108">The URL uses the HTTP **S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="1e4fb-133">Le domaine de la boîte de dialogue est le même que celui de la page hôte, qui peut être la page d’un volet Office ou le [fichier de fonctions](../reference/manifest/functionfile.md) d’une commande de complément.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-133">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="1e4fb-134">Obligatoire : la page, la méthode du contrôleur ou toute autre ressource qui est transmise à la méthode `displayDialogAsync` doit se trouver dans le même domaine que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-134">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1e4fb-135">La page hôte et les ressources s'ouvrant dans la boîte de dialogue doivent avoir le même domaine complet.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-135">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="1e4fb-136">Si vous tentez de transmettre `displayDialogAsync` à un sous-domaine du domaine du complément, cela ne fonctionnera pas.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-136">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="1e4fb-137">Le domaine complet et tous les sous-domaines doivent être exactement les mêmes.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-137">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="1e4fb-138">Une fois que la première page (ou toute autre ressource) est chargée, un utilisateur peut utiliser des liens ou une autre interface utilisateur pour accéder à n’importe quel site web (ou n’importe quelle autre ressource) qui utilise le protocole HTTPS.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-138">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="1e4fb-139">Vous pouvez également concevoir la première page de façon à ce que l’utilisateur soit immédiatement redirigé vers un autre site.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-139">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="1e4fb-140">Par défaut, la boîte de dialogue occupera 80 % de la hauteur et de la largeur de l’écran de l’appareil, mais vous pouvez définir des pourcentages différents en transmettant un objet de configuration à la méthode, comme indiqué dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-140">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="1e4fb-141">Pour voir un exemple de complément qui effectue ce type d’action, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-141">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span> <span data-ttu-id="1e4fb-142">Pour plus d’exemples qui `displayDialogAsync` utilisent, voir [Exemples.](#samples)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-142">For more samples that use `displayDialogAsync`, see [Samples](#samples).</span></span>

<span data-ttu-id="1e4fb-p113">Définissez les deux valeurs sur 100 % pour bénéficier d’une réelle d’expérience de plein écran. (Le maximum réel est de 99,5 %, et la fenêtre peut toujours être déplacée et redimensionnée.)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p113">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="1e4fb-p114">Vous ne pouvez ouvrir qu’une seule boîte de dialogue à partir d’une fenêtre hôte. Toute tentative d’ouverture d’une autre boîte de dialogue génère une erreur. Par exemple, si un utilisateur ouvre une boîte de dialogue à partir d’un volet Office, il ne peut pas ouvrir une seconde boîte de dialogue à partir d’une autre page dans le volet Office. Toutefois, quand une boîte de dialogue est ouverte à partir d’une [commande de complément](../design/add-in-commands.md), la commande ouvre un nouveau fichier HTML (mais invisible) chaque fois qu’elle est sélectionnée. Cela crée une nouvelle fenêtre hôte (invisible), afin que chaque fenêtre de ce type puisse lancer sa propre boîte de dialogue. Pour plus d’informations, reportez-vous à [Erreurs provenant de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p114">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="1e4fb-151">Tirer parti d’une option de performances dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="1e4fb-151">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="1e4fb-152">La propriété `displayInIframe` est une propriété supplémentaire dans l’objet de configuration que vous pouvez transmettre à `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-152">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="1e4fb-153">Lorsque cette propriété est définie sur `true` et que le complément est en cours d’exécution dans un document ouvert dans Office sur le web, la boîte de dialogue s’ouvre sous la forme d’un iframe flottant et non d’une fenêtre indépendante ; elle s’ouvre ainsi plus rapidement.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-153">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="1e4fb-154">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-154">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="1e4fb-155">La valeur par défaut est `false`, ce qui revient à omettre entièrement la propriété.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-155">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="1e4fb-156">Si le complément n’est pas exécuté dans Office sur le Web, le `displayInIframe` est ignoré.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-156">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="1e4fb-157">Vous ne devez **pas** utiliser `displayInIframe: true` si la boîte de dialogue redirige à un moment donné l’utilisateur vers une page qui ne peut pas être ouverte dans un IFrame.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-157">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="1e4fb-158">Par exemple, les pages de signature de nombreux services web populaires, tels que les comptes Google et Microsoft, ne peuvent pas être ouvertes dans un iFrame.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-158">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="1e4fb-159">Envoi d’informations à la page hôte à partir de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1e4fb-159">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="1e4fb-160">La boîte de dialogue ne peut pas communiquer avec la page hôte dans le volet Office, sauf si :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-160">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="1e4fb-161">la page active dans la boîte de dialogue se trouve dans le même domaine que la page hôte ;</span><span class="sxs-lookup"><span data-stu-id="1e4fb-161">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="1e4fb-162">La Office’API JavaScript est chargée dans la page.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-162">The Office JavaScript API library is loaded in the page.</span></span> <span data-ttu-id="1e4fb-163">(Comme pour toute page qui utilise la bibliothèque d’API JavaScript Office, le script de la page doit affecter une méthode à la propriété, bien qu’il puisse s’agit `Office.initialize` d’une méthode vide.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-163">(Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method.</span></span> <span data-ttu-id="1e4fb-164">Pour plus d’informations, [voir Initialize your Office Add-in](initialize-add-in.md).)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-164">For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="1e4fb-165">Le code de la boîte de dialogue utilise [la fonction messageParent](/javascript/api/office/office.ui#messageparent-message-) pour envoyer un message de chaîne à la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-165">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send a string message to the host page.</span></span> <span data-ttu-id="1e4fb-166">La chaîne peut être un mot, une phrase, un blob XML, un JSON stringified ou toute autre chaîne qui peut être sérialisée en chaîne ou castée en chaîne.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-166">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string or cast to a string.</span></span> <span data-ttu-id="1e4fb-167">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-167">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - <span data-ttu-id="1e4fb-168">La fonction `messageParent` peut uniquement être appelée sur une page ayant le même domaine (y compris les mêmes protocole et port) que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-168">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="1e4fb-169">La `messageParent` fonction est l’une *des* deux seules API JS Office qui peuvent être appelées dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-169">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span>
> - <span data-ttu-id="1e4fb-170">L’autre API JS qui peut être appelée dans la boîte de dialogue est `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="1e4fb-170">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="1e4fb-171">Pour plus d’informations à ce sujet, voir Spécifier les Office [applications et les api requises.](specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-171">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="1e4fb-172">Toutefois, dans la boîte de dialogue, cette API n’est pas prise en charge dans Outlook 2016'achat à prix seul (c’est-à-dire, la version MSI).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-172">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>

<span data-ttu-id="1e4fb-173">Dans l’exemple suivant, `googleProfile` est une version convertie en chaîne du profil Google de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-173">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="1e4fb-p121">La page hôte doit être configurée de façon à recevoir le message. Pour ce faire, ajoutez un paramètre de rappel à l’appel d’origine de `displayDialogAsync`. Le rappel attribue un gestionnaire à l’événement `DialogMessageReceived`. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p121">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="1e4fb-178">Office transmet un objet [AsyncResult](/javascript/api/office/office.asyncresult) au rappel.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-178">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="1e4fb-179">Il représente le résultat de la tentative d’ouverture de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-179">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="1e4fb-180">Il ne représente pas le résultat de tous les événements dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-180">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="1e4fb-181">Pour plus d’informations sur cette distinction, consultez la [Gestion des erreurs et des événements](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-181">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="1e4fb-182">La propriété `value` de `asyncResult` est définie sur un objet [Dialog](/javascript/api/office/office.dialog), qui existe dans la page hôte, pas dans le contexte d’exécution de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-182">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="1e4fb-p123">`processMessage` est la fonction qui gère l’événement. Vous pouvez lui donner le nom que vous souhaitez.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p123">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="1e4fb-185">La variable `dialog` est déclarée avec une portée plus large que le rappel, car elle est également référencée dans `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-185">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="1e4fb-186">Voici un exemple simple de gestionnaire pour l’événement `DialogMessageReceived` :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-186">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="1e4fb-187">Office transmet l’objet `arg` au gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-187">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="1e4fb-188">Sa `message` propriété est la chaîne envoyée par l’appel de la boîte de `messageParent` dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-188">Its `message` property is the string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="1e4fb-189">Dans cet exemple, il s’agit d’une représentation sous la chaîne du profil d’un utilisateur à partir d’un service tel qu’un compte Microsoft ou Google, de sorte qu’il est désérialisé à un objet avec `JSON.parse` .</span><span class="sxs-lookup"><span data-stu-id="1e4fb-189">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="1e4fb-p125">L’implémentation `showUserName` n’est pas visible. Elle peut afficher un message de bienvenue personnalisé dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p125">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="1e4fb-192">Lorsque l’intervention de l’utilisateur sur la boîte de dialogue est terminée, votre gestionnaire de messages doit fermer la boîte de dialogue, comme indiqué dans cet exemple.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-192">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="1e4fb-193">L’objet `dialog` doit être le même que celui renvoyé par l’appel de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-193">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="1e4fb-194">L’appel de `dialog.close` indique à Office de fermer immédiatement la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-194">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="1e4fb-195">Pour voir un exemple de complément qui utilise ces techniques, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-195">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="1e4fb-p126">Si le complément a besoin d’ouvrir une autre page du volet Office après avoir reçu le message, vous pouvez utiliser la méthode `window.location.replace` (ou `window.location.href`) en tant que dernière ligne du gestionnaire. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p126">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="1e4fb-198">Pour voir un exemple de complément qui effectue ce type d’action, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-198">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="1e4fb-199">Messagerie conditionnelle</span><span class="sxs-lookup"><span data-stu-id="1e4fb-199">Conditional messaging</span></span>

<span data-ttu-id="1e4fb-200">Étant donné que vous pouvez envoyer plusieurs appels `messageParent` à partir de la boîte de dialogue, mais que vous n’avez qu’un seul gestionnaire dans la page hôte pour l’événement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-200">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="1e4fb-201">Par exemple, si la boîte de dialogue invite un utilisateur à se connecter à un fournisseur d’identité tel qu’un compte Microsoft ou Google, elle envoie le profil de l’utilisateur en tant que message.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-201">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="1e4fb-202">Si l’authentification échoue, la boîte de dialogue envoie des informations sur l’erreur à la page hôte, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-202">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="1e4fb-203">La variable `loginSuccess` serait initialisée en lisant la réponse HTTP à partir du fournisseur d’identité.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-203">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="1e4fb-p128">L’implémentation des fonctions `getProfile` et `getError` n’est pas affichée. Chacune obtient des données à partir d’un paramètre de requête ou du corps de la réponse HTTP.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p128">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="1e4fb-p129">Des objets anonymes de différents types sont envoyés selon que la connexion a réussi ou non. Tous deux ont une propriété `messageType`, mais un a une propriété `profile` et l’autre une propriété `error`.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p129">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="1e4fb-p130">Le code du gestionnaire dans la page hôte utilise la valeur de la propriété `messageType` pour créer une branche comme le montre l’exemple suivant. Notez que la fonction `showUserName` est identique à celle de l’exemple précédent et que la fonction `showNotification` affiche l’erreur dans l’interface utilisateur de la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p130">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="1e4fb-210">L'implémentation `showNotification` n'est pas montrée dans l'exemple de code fourni par cet article.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-210">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="1e4fb-211">Pour un exemple d'implémentation de cette fonction dans votre complément, voir [Exemple d'API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-211">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="1e4fb-212">Transmission d’informations à la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1e4fb-212">Pass information to the dialog box</span></span>

<span data-ttu-id="1e4fb-213">Votre add-in peut envoyer des messages à partir de la [page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) vers une boîte de dialogue à l’aide [de Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-213">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="1e4fb-214">Utilisation `messageChild()` à partir de la page hôte</span><span class="sxs-lookup"><span data-stu-id="1e4fb-214">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="1e4fb-215">Lorsque vous appelez l’API Office boîte de dialogue pour ouvrir une boîte de dialogue, un objet [Dialog](/javascript/api/office/office.dialog) est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-215">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="1e4fb-216">Elle doit être affectée à une variable dont l’étendue est supérieure à celle de la méthode [displayDialogAsync,](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) car l’objet sera référencé par d’autres méthodes.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-216">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="1e4fb-217">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-217">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="1e4fb-218">Cet objet possède une méthode messageChild qui envoie toute chaîne, y compris les données `Dialog` stringified, [](/javascript/api/office/office.dialog#messagechild-message-) à la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-218">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="1e4fb-219">Cela lève un événement `DialogParentMessageReceived` dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-219">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="1e4fb-220">Votre code doit gérer cet événement, comme illustré dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-220">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="1e4fb-221">Envisagez un scénario dans lequel l’interface utilisateur de la boîte de dialogue est liée à la feuille de calcul active et à sa position par rapport aux autres feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-221">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="1e4fb-222">Dans l’exemple suivant, `sheetPropertiesChanged` Excel propriétés de feuille de calcul à la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-222">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="1e4fb-223">Dans ce cas, la feuille de calcul actuelle est nommée « My Sheet » et il s’agit de la deuxième feuille du workbook.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-223">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="1e4fb-224">Les données sont encapsulées dans un objet et stringified afin de pouvoir être transmises à `messageChild` .</span><span class="sxs-lookup"><span data-stu-id="1e4fb-224">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="1e4fb-225">Gérer DialogParentMessageReceived dans la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1e4fb-225">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="1e4fb-226">Dans le javaScript de la boîte de dialogue, inscrivez un handler pour l’événement à l’auprès de la méthode `DialogParentMessageReceived` [UI.addHandlerAsync.](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-226">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="1e4fb-227">Cette méthode est généralement effectuée dans [les méthodes Office.onReady ou Office.initialize,](initialize-add-in.md)comme illustré ci-après.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-227">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="1e4fb-228">(Vous trouverez ci-dessous un exemple plus robuste.)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-228">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="1e4fb-229">Ensuite, définissez `onMessageFromParent` le handler.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-229">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="1e4fb-230">Le code suivant poursuit l’exemple de la section précédente.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-230">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="1e4fb-231">Notez Office passe un argument au handler et que la propriété de l’objet argument contient la chaîne de `message` la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-231">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="1e4fb-232">Dans cet exemple, le message est reconverti en objet et jQuery est utilisé pour définir le titre supérieur de la boîte de dialogue afin qu’il corresponde au nouveau nom de feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-232">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="1e4fb-233">Il est préférable de vérifier que votre handler est correctement enregistré.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-233">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="1e4fb-234">Vous pouvez le faire en passant un rappel à la `addHandlerAsync` méthode.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-234">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="1e4fb-235">Cette tentative s’exécute lorsque la tentative d’inscription du handler est terminée.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-235">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="1e4fb-236">Utilisez le handler pour enregistrer ou afficher une erreur si le handler n’a pas été enregistré avec succès.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-236">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="1e4fb-237">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-237">The following is an example.</span></span> <span data-ttu-id="1e4fb-238">Notez `reportError` qu’il s’agit d’une fonction, qui n’est pas définie ici, qui enregistre ou affiche l’erreur.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-238">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="1e4fb-239">Messagerie conditionnelle d’une page parent à une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1e4fb-239">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="1e4fb-240">Étant donné que vous pouvez effectuer plusieurs appels à partir de la page hôte, mais que vous n’avez qu’un seul responsable dans la boîte de dialogue pour l’événement, le responsable doit utiliser une logique conditionnelle pour distinguer les `messageChild` `DialogParentMessageReceived` différents messages.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-240">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="1e4fb-241">Vous pouvez le faire d’une manière qui est précisément parallèle à la façon dont vous structureriez la messagerie conditionnelle lorsque la boîte de dialogue envoie un message à la page hôte, comme décrit dans la messagerie [conditionnelle.](#conditional-messaging)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-241">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="1e4fb-242">Dans certains cas, l’API, qui fait partie de l’ensemble de conditions `messageChild` [requises DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)peut ne pas être prise en charge.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-242">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="1e4fb-243">D’autres méthodes de messagerie de parent à boîte de dialogue sont décrites de manière alternative pour transmettre des messages à une boîte de dialogue à partir de [sa page hôte.](parent-to-dialog.md)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-243">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1e4fb-244">L’ensemble de conditions [requises DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) ne peut pas être spécifié dans la `<Requirements>` section d’un manifeste de add-in.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-244">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="1e4fb-245">Vous devez vérifier la prise en charge de DialogApi 1.2 lors de l’runtime à l’aide de la méthode [isSetSupported.](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-245">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="1e4fb-246">La prise en charge des exigences de manifeste est en cours de développement.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-246">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="1e4fb-247">Fermeture de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1e4fb-247">Closing the dialog box</span></span>

<span data-ttu-id="1e4fb-p141">Vous pouvez implémenter un bouton de fermeture dans la boîte de dialogue. Pour ce faire, le gestionnaire d’événements Click du bouton doit utiliser `messageParent` pour indiquer à la page hôte que vous avez cliqué sur le bouton. Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="1e4fb-p141">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="1e4fb-251">Le gestionnaire de la page hôte pour `DialogMessageReceived` appelle `dialog.close`, comme dans cet exemple.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-251">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="1e4fb-252">(consultez les exemples précédents qui montrent comment l’objet `dialog` est initialisé).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-252">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="1e4fb-253">Même lorsque vous ne disposez pas de votre propre interface utilisateur de fermeture de boîte de dialogue, un utilisateur final peut fermer la boîte de dialogue en choisissant le **X** dans le coin supérieur droit.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-253">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="1e4fb-254">Cette action déclenche l’événement `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-254">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="1e4fb-255">Si votre volet hôte a besoin de savoir quand cela se produit, il doit déclarer un gestionnaire pour cet événement.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-255">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="1e4fb-256">Pour plus d’informations, consultez la section [Erreurs et événements dans la boîte de dialogue](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-256">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="1e4fb-257">Rubriques plus complexes et scénarios spéciaux</span><span class="sxs-lookup"><span data-stu-id="1e4fb-257">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="1e4fb-258">Utilisation d'un API de boîte de dialogue pour afficher une vidéo</span><span class="sxs-lookup"><span data-stu-id="1e4fb-258">Use the Dialog API to show a video</span></span>

<span data-ttu-id="1e4fb-259">Voir [Utiliser la boîte de dialogue Office pour afficher une vidéo](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-259">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="1e4fb-260">Utilisation des API de boîte de dialogue dans un flux d’authentification</span><span class="sxs-lookup"><span data-stu-id="1e4fb-260">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="1e4fb-261">Voir [Authentifier avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-261">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="1e4fb-262">Utilisation de l’API de boîte de dialogue Office avec des applications à page unique et routage côté client</span><span class="sxs-lookup"><span data-stu-id="1e4fb-262">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="1e4fb-263">Les authentifications par mot de passe (SPA) et le routage client doivent être gérés avec précaution lorsque vous utilisez l’API de boîte de dialogue Office.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-263">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="1e4fb-264">Consultez les [Pratiques recommandées pour l’utilisation de l’API de boîte de dialogue Office dans une SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-264">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="1e4fb-265">Gestion d'erreurs et d'événements</span><span class="sxs-lookup"><span data-stu-id="1e4fb-265">Error and event handling</span></span>

<span data-ttu-id="1e4fb-266">Voir [Gestion des erreurs et des événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-266">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="1e4fb-267">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="1e4fb-267">Next steps</span></span>

<span data-ttu-id="1e4fb-268">Découvrez les pièges et pratiques recommandées pour l’API de boîte de dialogue Office dans les [Meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="1e4fb-268">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

## <a name="samples"></a><span data-ttu-id="1e4fb-269">Exemples</span><span class="sxs-lookup"><span data-stu-id="1e4fb-269">Samples</span></span>

<span data-ttu-id="1e4fb-270">Tous les exemples suivants utilisent `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="1e4fb-270">All of the following samples use `displayDialogAsync`.</span></span> <span data-ttu-id="1e4fb-271">Certains ont des serveurs NodeJS et d’autres des serveurs basés sur ASP.NET/IIS, mais la logique d’utilisation de la méthode est la même quelle que soit la façon dont le côté serveur du module est implémenté.</span><span class="sxs-lookup"><span data-stu-id="1e4fb-271">Some have NodeJS-based servers and others have ASP.NET/IIS-based servers, but the logic of using the method is the same regardless of how the server-side of the add-in is implemented.</span></span>

<span data-ttu-id="1e4fb-272">**Informations de base :**</span><span class="sxs-lookup"><span data-stu-id="1e4fb-272">**Basics:**</span></span>

- [<span data-ttu-id="1e4fb-273">Exemple d’API de dialogue pour compléments Office</span><span class="sxs-lookup"><span data-stu-id="1e4fb-273">Office Add-in Dialog API Example</span></span>](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [<span data-ttu-id="1e4fb-274">Contenu de formation/ Création de add-ins (plusieurs exemples)</span><span class="sxs-lookup"><span data-stu-id="1e4fb-274">Training Content / Building Add-ins (several samples)</span></span>](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

<span data-ttu-id="1e4fb-275">**Exemples plus complexes :**</span><span class="sxs-lookup"><span data-stu-id="1e4fb-275">**More complex samples:**</span></span>

- [<span data-ttu-id="1e4fb-276">Office Add-in Microsoft Graph ASPNET</span><span class="sxs-lookup"><span data-stu-id="1e4fb-276">Office Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="1e4fb-277">Complément Office Microsoft Graph React</span><span class="sxs-lookup"><span data-stu-id="1e4fb-277">Office Add-in Microsoft Graph React</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [<span data-ttu-id="1e4fb-278">SSO NodeJS pour complément Office</span><span class="sxs-lookup"><span data-stu-id="1e4fb-278">Office Add-in NodeJS SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [<span data-ttu-id="1e4fb-279">Office Add-in ASPNET SSO</span><span class="sxs-lookup"><span data-stu-id="1e4fb-279">Office Add-in ASPNET SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [<span data-ttu-id="1e4fb-280">Office Exemple de monétisation SAAS pour le add-in</span><span class="sxs-lookup"><span data-stu-id="1e4fb-280">Office Add-in SAAS Monetization Sample</span></span>](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [<span data-ttu-id="1e4fb-281">Outlook Add-in Microsoft Graph ASPNET</span><span class="sxs-lookup"><span data-stu-id="1e4fb-281">Outlook Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="1e4fb-282">Outlook Add-in SSO</span><span class="sxs-lookup"><span data-stu-id="1e4fb-282">Outlook Add-in SSO</span></span>](https://github.com/OfficeDev/Outlook-Add-in-SSO)
- [<span data-ttu-id="1e4fb-283">Outlook Visionneuse de jetons de add-in</span><span class="sxs-lookup"><span data-stu-id="1e4fb-283">Outlook Add-in Token Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="1e4fb-284">Outlook Message actionnable du add-in</span><span class="sxs-lookup"><span data-stu-id="1e4fb-284">Outlook Add-in Actionable Message</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [<span data-ttu-id="1e4fb-285">Outlook Partage de OneDrive</span><span class="sxs-lookup"><span data-stu-id="1e4fb-285">Outlook Add-in Sharing to OneDrive</span></span>](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [<span data-ttu-id="1e4fb-286">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="1e4fb-286">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="1e4fb-287">Excel Scénario d’runtime partagé</span><span class="sxs-lookup"><span data-stu-id="1e4fb-287">Excel Shared Runtime Scenario</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/900b5769bca9bbcff79d6cd6106d9fcc55c70d5a/Samples/excel-shared-runtime-scenario)
- [<span data-ttu-id="1e4fb-288">Excel Add-in ASPNET QuickBooks</span><span class="sxs-lookup"><span data-stu-id="1e4fb-288">Excel Add-in ASPNET QuickBooks</span></span>](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [<span data-ttu-id="1e4fb-289">Word Add-in JS Redact</span><span class="sxs-lookup"><span data-stu-id="1e4fb-289">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="1e4fb-290">Word Add-in JS SpecKit</span><span class="sxs-lookup"><span data-stu-id="1e4fb-290">Word Add-in JS SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [<span data-ttu-id="1e4fb-291">Word Add-in AngularJS Client OAuth</span><span class="sxs-lookup"><span data-stu-id="1e4fb-291">Word Add-in AngularJS Client OAuth</span></span>](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [<span data-ttu-id="1e4fb-292">Complément Office dans Auth0</span><span class="sxs-lookup"><span data-stu-id="1e4fb-292">Office Add-in Auth0</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="1e4fb-293">Office Ajout de OAuth.io</span><span class="sxs-lookup"><span data-stu-id="1e4fb-293">Office Add-in OAuth.io</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [<span data-ttu-id="1e4fb-294">Office Code de modèles de conception de l’UX de l’UX de l’ajout</span><span class="sxs-lookup"><span data-stu-id="1e4fb-294">Office Add-in UX Design Patterns Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
