---
title: Chargement du DOM et de l’environnement d’exécution
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: ac4d26d964f844f08e1d2975c1be8bbccf40349f
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27271061"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="ba24e-102">Chargement du DOM et de l’environnement d’exécution</span><span class="sxs-lookup"><span data-stu-id="ba24e-102">Loading the DOM and runtime environment</span></span>



<span data-ttu-id="ba24e-103">Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.</span><span class="sxs-lookup"><span data-stu-id="ba24e-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="ba24e-104">Démarrage d’un complément de contenu ou du volet Office</span><span class="sxs-lookup"><span data-stu-id="ba24e-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="ba24e-105">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project, Word ou Access.</span><span class="sxs-lookup"><span data-stu-id="ba24e-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![Flux des événements au démarrage d’un complément de contenu ou du volet Office](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="ba24e-107">Les événements suivants se produisent lors du démarrage d’un complément de contenu ou du volet Office :</span><span class="sxs-lookup"><span data-stu-id="ba24e-107">The following events occur when a content or task pane add-in starts:</span></span> 



1. <span data-ttu-id="ba24e-108">L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.</span><span class="sxs-lookup"><span data-stu-id="ba24e-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>
    
2. <span data-ttu-id="ba24e-109">L’application hôte Office lit le manifeste XML du complément à partir d’AppSource, d’un catalogue de compléments sur SharePoint ou du catalogue de dossiers partagés duquel il provient.</span><span class="sxs-lookup"><span data-stu-id="ba24e-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>
    
3. <span data-ttu-id="ba24e-110">L’application hôte Office ouvre la page HTML du complément dans un contrôle de navigateur.</span><span class="sxs-lookup"><span data-stu-id="ba24e-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>
    
    <span data-ttu-id="ba24e-p101">Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="ba24e-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>
    
4. <span data-ttu-id="ba24e-113">Le contrôle de navigateur charge le DOM et le corps HTML, puis demande au gestionnaire d’événements l’événement  **window.onload**.</span><span class="sxs-lookup"><span data-stu-id="ba24e-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>
    
5. <span data-ttu-id="ba24e-114">L’application hôte Office charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) de l’objet [Office](https://docs.microsoft.com/javascript/api/office?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="ba24e-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event of the [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) object.</span></span>
    
6. <span data-ttu-id="ba24e-115">Lorsque le chargement du modèle objet de document (DOM) et du corps HTML est terminé et que le complément s’est initialisé, la fonction principale de l’application peut s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="ba24e-115">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="ba24e-116">Démarrage d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="ba24e-116">Startup of an Outlook add-in</span></span>



<span data-ttu-id="ba24e-117">La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.</span><span class="sxs-lookup"><span data-stu-id="ba24e-117">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Flux des événements au démarrage du complément Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="ba24e-119">Les événements suivants se produisent lors du démarrage d’un complément Outlook :</span><span class="sxs-lookup"><span data-stu-id="ba24e-119">The following events occur when an Outlook add-in starts:</span></span> 



1. <span data-ttu-id="ba24e-120">Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ba24e-120">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>
    
2. <span data-ttu-id="ba24e-121">L’utilisateur sélectionne un élément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="ba24e-121">The user selects an item in Outlook.</span></span>
    
3. <span data-ttu-id="ba24e-122">Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ba24e-122">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>
    
4. <span data-ttu-id="ba24e-p102">Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.</span><span class="sxs-lookup"><span data-stu-id="ba24e-p102">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>
    
5. <span data-ttu-id="ba24e-125">Le contrôle de navigateur charge le modèle objet de document (DOM) et le corps HTML, puis appelle le gestionnaire d’événements pour l’événement  **onload**.</span><span class="sxs-lookup"><span data-stu-id="ba24e-125">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>
    
6. <span data-ttu-id="ba24e-126">Outlook appelle le gestionnaire d’événements pour l’événement [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) de l’objet [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) du complément.</span><span class="sxs-lookup"><span data-stu-id="ba24e-126">Outlook calls the event handler for the [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event of the [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) object of the add-in.</span></span>
    
7. <span data-ttu-id="ba24e-127">Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.</span><span class="sxs-lookup"><span data-stu-id="ba24e-127">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="checking-the-load-status"></a><span data-ttu-id="ba24e-128">Vérification du statut de chargement</span><span class="sxs-lookup"><span data-stu-id="ba24e-128">Checking the load status</span></span>


<span data-ttu-id="ba24e-p103">Pour vérifier que le chargement du modèle objet de document (DOM) et de l’environnement d’exécution des est terminé, il est notamment possible d’utiliser la fonction jQuery [.ready()](https://api.jquery.com/ready/) :  `$(document).ready()`. Par exemple, la fonction de gestionnaire d’événements  **initialize** ci-dessous s’assure d’abord que le DOM est bien chargé avant l’exécution du code d’initialisation du complément. Par conséquent, le gestionnaire d’événements **initialize** utilise la propriété [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) pour obtenir l’élément actuellement sélectionné dans Outlook, puis appelle la fonction principale du complément, `initDialer`.</span><span class="sxs-lookup"><span data-stu-id="ba24e-p103">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`. For example, the following  **initialize** event handler function makes sure the DOM is first loaded before the code specific to initializing the add-in runs. Subsequently, the **initialize** event handler proceeds to use the [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>


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

<span data-ttu-id="ba24e-132">Il est possible d’utiliser cette même technique dans le gestionnaire  **initialize** de toute Complément Office.</span><span class="sxs-lookup"><span data-stu-id="ba24e-132">This same technique can be used in the  **initialize** handler of any Office Add-in.</span></span>

<span data-ttu-id="ba24e-133">Le numéroteur téléphonique fourni comme exemple de complément Outlook présente une approche légèrement différente, puisqu’il utilise uniquement JavaScript pour vérifier ces mêmes conditions.</span><span class="sxs-lookup"><span data-stu-id="ba24e-133">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="ba24e-134">Même si aucune tâche d’initialisation n’est à effectuer dans votre complément, vous devez inclure au moins une fonction de gestionnaire d’événements **Office.initialize** minimale comme l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="ba24e-134">Even if your add-in has no initialization tasks to perform, you must include at least a minimal **Office.initialize** event handler function like the following example.</span></span>

```js
Office.initialize = function () {
};
```

<span data-ttu-id="ba24e-p104">Si vous n’incluez pas de gestionnaire d’événements  **Office.initialize**, votre complément peut générer une erreur au démarrage. En outre, si un utilisateur tente d’utiliser votre complément avec un client web Office Online, comme Excel Online, PowerPoint Online ou Outlook Web App, il n’est pas exécuté.</span><span class="sxs-lookup"><span data-stu-id="ba24e-p104">If you fail to include an  **Office.initialize** event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>

<span data-ttu-id="ba24e-137">Si votre complément comprend plusieurs pages, chaque fois qu’il charge une nouvelle page, celle-ci doit inclure ou appeler un gestionnaire d’événements  **Office.initialize**.</span><span class="sxs-lookup"><span data-stu-id="ba24e-137">If your add-in includes more than one page, whenever it loads a new page that page must include or call an  **Office.initialize** event handler.</span></span>


## <a name="see-also"></a><span data-ttu-id="ba24e-138">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ba24e-138">See also</span></span>

- [<span data-ttu-id="ba24e-139">Présentation de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="ba24e-139">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
    
