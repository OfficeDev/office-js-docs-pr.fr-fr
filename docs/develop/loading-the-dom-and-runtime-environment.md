---
title: Chargement du DOM et de l’environnement d’exécution
description: Charger le DOM et l’environnement d’exécution des compléments Office
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 02f950ca23d52b333f704c7d8aed431cb426a6f0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293274"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="51cd9-103">Chargement du DOM et de l’environnement d’exécution</span><span class="sxs-lookup"><span data-stu-id="51cd9-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="51cd9-104">Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.</span><span class="sxs-lookup"><span data-stu-id="51cd9-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="51cd9-105">Démarrage d’un complément de contenu ou du volet Office</span><span class="sxs-lookup"><span data-stu-id="51cd9-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="51cd9-106">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project ou Word.</span><span class="sxs-lookup"><span data-stu-id="51cd9-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Flux des événements au démarrage d’un complément de contenu ou du volet Office](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="51cd9-108">Les événements suivants se produisent lors du démarrage d’un complément de contenu ou du volet Office :</span><span class="sxs-lookup"><span data-stu-id="51cd9-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="51cd9-109">L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.</span><span class="sxs-lookup"><span data-stu-id="51cd9-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="51cd9-110">L’application cliente Office lit le manifeste XML du complément à partir de AppSource, d’un catalogue d’applications sur SharePoint ou du catalogue de dossiers partagés duquel il provient.</span><span class="sxs-lookup"><span data-stu-id="51cd9-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="51cd9-111">L’application cliente Office ouvre la page HTML du complément dans un contrôle de navigateur.</span><span class="sxs-lookup"><span data-stu-id="51cd9-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="51cd9-p101">Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="51cd9-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="51cd9-114">Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d’événements pour l' `window.onload` événement.</span><span class="sxs-lookup"><span data-stu-id="51cd9-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="51cd9-115">L’application cliente Office charge l’environnement d’exécution, qui télécharge et met en cache les fichiers de la bibliothèque de l’API JavaScript pour Office à partir du serveur de réseau de distribution de contenu (CDN), puis appelle le gestionnaire d’événements du complément pour l’événement [Initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) , si un gestionnaire lui a été attribué.</span><span class="sxs-lookup"><span data-stu-id="51cd9-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="51cd9-116">Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="51cd9-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="51cd9-117">Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [initialiser votre complément](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="51cd9-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="51cd9-118">Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.</span><span class="sxs-lookup"><span data-stu-id="51cd9-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="51cd9-119">Démarrage d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="51cd9-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="51cd9-120">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.</span><span class="sxs-lookup"><span data-stu-id="51cd9-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Flux des événements au démarrage du complément Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="51cd9-122">Les événements suivants se produisent lors du démarrage d’un complément Outlook :</span><span class="sxs-lookup"><span data-stu-id="51cd9-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="51cd9-123">Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="51cd9-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="51cd9-124">L’utilisateur sélectionne un élément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="51cd9-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="51cd9-125">Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="51cd9-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="51cd9-p103">Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.</span><span class="sxs-lookup"><span data-stu-id="51cd9-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="51cd9-128">Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d’événements pour l' `onload` événement.</span><span class="sxs-lookup"><span data-stu-id="51cd9-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="51cd9-129">Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté.</span><span class="sxs-lookup"><span data-stu-id="51cd9-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="51cd9-130">Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="51cd9-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="51cd9-131">Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [initialiser votre complément](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="51cd9-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="51cd9-132">Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.</span><span class="sxs-lookup"><span data-stu-id="51cd9-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="51cd9-133">Vérification du statut de chargement</span><span class="sxs-lookup"><span data-stu-id="51cd9-133">Checking the load status</span></span>

<span data-ttu-id="51cd9-134">Vous pouvez vérifier que le chargement du DOM et de l’environnement d’exécution est bien terminé en utilisant la fonction jQuery [.ready()](https://api.jquery.com/ready/) : `$(document).ready()`.</span><span class="sxs-lookup"><span data-stu-id="51cd9-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="51cd9-135">Par exemple, le `onReady` Gestionnaire d’événements suivant vérifie que le DOM est chargé pour la première fois avant l’exécution du code spécifique à l’initialisation du complément.</span><span class="sxs-lookup"><span data-stu-id="51cd9-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="51cd9-136">Par la suite, le `onReady` Gestionnaire continue d’utiliser la propriété [Mailbox. Item](/javascript/api/outlook/office.mailbox#item) pour obtenir l’élément actuellement sélectionné dans Outlook et appelle la fonction principale du complément, `initDialer` .</span><span class="sxs-lookup"><span data-stu-id="51cd9-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox#item) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

<span data-ttu-id="51cd9-137">Vous pouvez également utiliser le même code dans un gestionnaire d' `initialize` événements comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="51cd9-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

<span data-ttu-id="51cd9-138">Cette même technique peut être utilisée dans les `onReady` `initialize` gestionnaires ou des compléments Office.</span><span class="sxs-lookup"><span data-stu-id="51cd9-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="51cd9-139">Le numéroteur téléphonique fourni comme exemple de complément Outlook présente une approche légèrement différente, puisqu’il utilise uniquement JavaScript pour vérifier ces mêmes conditions.</span><span class="sxs-lookup"><span data-stu-id="51cd9-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="51cd9-140">Même si aucune tâche d’initialisation n’est à effectuer dans votre complément, vous devez inclure au moins un appel de `Office.onReady` `Office.initialize` la fonction de gestionnaire d’événements minimal, comme illustré dans les exemples suivants.</span><span class="sxs-lookup"><span data-stu-id="51cd9-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="51cd9-141">Si vous n’appelez pas `Office.onReady` ou n’assignez pas de `Office.initialize` Gestionnaire d’événements, votre complément peut déclencher une erreur lors de son démarrage.</span><span class="sxs-lookup"><span data-stu-id="51cd9-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="51cd9-142">En outre, si un utilisateur essaie d’utiliser votre complément avec un client web Office, notamment Excel, PowerPoint ou Outlook, l’exécution du complément échouera.</span><span class="sxs-lookup"><span data-stu-id="51cd9-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="51cd9-143">Si votre complément comprend plusieurs pages, chaque fois qu’il charge une nouvelle page, celle-ci doit appeler `Office.onReady` ou assigner un `Office.initialize` Gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="51cd9-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="51cd9-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="51cd9-144">See also</span></span>

- [<span data-ttu-id="51cd9-145">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="51cd9-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="51cd9-146">Initialiser votre complément Office</span><span class="sxs-lookup"><span data-stu-id="51cd9-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)
