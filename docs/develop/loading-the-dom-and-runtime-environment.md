---
title: Chargement du DOM et de l’environnement d’exécution
description: Chargez le DOM et Office’environnement d’runtime des add-ins.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 0cfdcf3750d9c0a3dd21667729da59dbfedf61c8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349839"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="ee33e-103">Chargement du DOM et de l’environnement d’exécution</span><span class="sxs-lookup"><span data-stu-id="ee33e-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="ee33e-104">Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.</span><span class="sxs-lookup"><span data-stu-id="ee33e-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="ee33e-105">Démarrage d’un complément de contenu ou du volet Office</span><span class="sxs-lookup"><span data-stu-id="ee33e-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="ee33e-106">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project ou Word.</span><span class="sxs-lookup"><span data-stu-id="ee33e-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Flow événements lors du démarrage d’un module de contenu ou du volet Des tâches.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="ee33e-108">Les événements suivants se produisent lors du démarrage d’un module de contenu ou du volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="ee33e-108">The following events occur when a content or task pane add-in starts.</span></span>

1. <span data-ttu-id="ee33e-109">L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.</span><span class="sxs-lookup"><span data-stu-id="ee33e-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="ee33e-110">L’application cliente Office lit le manifeste XML du add-in à partir d’AppSource, d’un catalogue d’applications sur SharePoint ou du catalogue de dossiers partagés dont il est issu.</span><span class="sxs-lookup"><span data-stu-id="ee33e-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="ee33e-111">L Office application cliente ouvre la page HTML du module dans un contrôle de navigateur.</span><span class="sxs-lookup"><span data-stu-id="ee33e-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="ee33e-p101">Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="ee33e-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="ee33e-114">Le contrôle de navigateur charge le DOM et le corps HTML, puis appelle le responsable de l’événement `window.onload` pour l’événement.</span><span class="sxs-lookup"><span data-stu-id="ee33e-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="ee33e-115">L’application cliente Office charge l’environnement d’utilisation, qui télécharge et met en cache les fichiers de bibliothèque d’API JavaScript Office à partir du serveur de réseau de distribution de contenu (CDN), puis appelle le responsable des événements du module pour [l’événement d’initialisation](/javascript/api/office#office-initialize-reason-) de l’objet [Office,](/javascript/api/office) si un handler lui a été affecté.</span><span class="sxs-lookup"><span data-stu-id="ee33e-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="ee33e-116">Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="ee33e-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="ee33e-117">Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [Initialiser votre add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ee33e-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="ee33e-118">Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.</span><span class="sxs-lookup"><span data-stu-id="ee33e-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="ee33e-119">Démarrage d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="ee33e-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="ee33e-120">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.</span><span class="sxs-lookup"><span data-stu-id="ee33e-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Flow d’événements au démarrage Outlook de votre module.](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="ee33e-122">Les événements suivants se produisent lorsqu’un Outlook de démarrage.</span><span class="sxs-lookup"><span data-stu-id="ee33e-122">The following events occur when an Outlook add-in starts.</span></span>

1. <span data-ttu-id="ee33e-123">Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ee33e-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="ee33e-124">L’utilisateur sélectionne un élément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee33e-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="ee33e-125">Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ee33e-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="ee33e-p103">Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.</span><span class="sxs-lookup"><span data-stu-id="ee33e-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="ee33e-128">Le contrôle de navigateur charge le DOM et le corps HTML, puis appelle le responsable de l’événement `onload` pour l’événement.</span><span class="sxs-lookup"><span data-stu-id="ee33e-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="ee33e-129">Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté.</span><span class="sxs-lookup"><span data-stu-id="ee33e-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="ee33e-130">Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="ee33e-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="ee33e-131">Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [Initialiser votre add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ee33e-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="ee33e-132">Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.</span><span class="sxs-lookup"><span data-stu-id="ee33e-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee33e-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ee33e-133">See also</span></span>

- [<span data-ttu-id="ee33e-134">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="ee33e-134">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="ee33e-135">Initialiser votre complément Office</span><span class="sxs-lookup"><span data-stu-id="ee33e-135">Initialize your Office Add-in</span></span>](initialize-add-in.md)
