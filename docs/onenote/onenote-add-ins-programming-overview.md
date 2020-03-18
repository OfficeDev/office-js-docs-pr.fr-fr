---
title: Vue d’ensemble de la programmation de l’API JavaScript de OneNote
description: En savoir plus sur l’API JavaScript de OneNote pour les compléments OneNote sur le web.
ms.date: 02/19/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 0e551b75d55da77d383e1335c27724834bfb2df0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720896"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="54172-103">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="54172-103">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="54172-104">OneNote présente une API JavaScript pour les compléments OneNote sur le web.</span><span class="sxs-lookup"><span data-stu-id="54172-104">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="54172-105">Vous pouvez créer des compléments de volet Office et de contenu, ainsi que des commandes de complément qui interagissent avec les objets OneNote et se connectent à des services web ou à d’autres ressources basées sur le web.</span><span class="sxs-lookup"><span data-stu-id="54172-105">You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="54172-p102">Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="54172-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="54172-108">Composants d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="54172-108">Components of an Office Add-in</span></span>

<span data-ttu-id="54172-109">Les compléments sont constitués de deux composants de base :</span><span class="sxs-lookup"><span data-stu-id="54172-109">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="54172-110">Une **application web** comportant une page web et les fichiers CSS, JavaScript ou autres requis.</span><span class="sxs-lookup"><span data-stu-id="54172-110">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="54172-111">Ces fichiers sont hébergés sur un serveur web ou un service d’hébergement web, tel que Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="54172-111">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="54172-112">Dans OneNote sur le web, l’application web s’affiche dans un contrôle de navigateur ou un iFrame.</span><span class="sxs-lookup"><span data-stu-id="54172-112">In OneNote on the web, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="54172-p104">Un **manifeste XML** spécifiant l’URL de la page web du complément, ainsi que les conditions d’accès, les paramètres et fonctionnalités du complément. Ce fichier est stocké sur le client. Les compléments OneNote utilisent le même format de [manifeste](../develop/add-in-manifests.md) que les autres compléments Office.</span><span class="sxs-lookup"><span data-stu-id="54172-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="54172-116">**Complément pour Office = manifeste + page web**</span><span class="sxs-lookup"><span data-stu-id="54172-116">**Office Add-in = Manifest + Webpage**</span></span>

![Un complément Office se compose d’un manifeste et d’une page web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="54172-118">Utilisation de l’API JavaScript</span><span class="sxs-lookup"><span data-stu-id="54172-118">Using the JavaScript API</span></span>

<span data-ttu-id="54172-p105">Les compléments utilisent le contexte d’exécution de l’application hôte pour accéder à l’API JavaScript. L’API comporte deux couches:</span><span class="sxs-lookup"><span data-stu-id="54172-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span>

- <span data-ttu-id="54172-121">Une **API enrichie** pour les opérations spécifiques de OneNote, accessible via l’objet`Application`Application.</span><span class="sxs-lookup"><span data-stu-id="54172-121">A **host-specific API** for OneNote-specific operations, accessed through the `Application` object.</span></span>
- <span data-ttu-id="54172-122">Une**API commune** qui est partagée entre les applications Office, accessible via l’objet `Document`.</span><span class="sxs-lookup"><span data-stu-id="54172-122">A **Common API** that's shared across Office applications, accessed through the `Document` object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="54172-123">Accès à l’API enrichie via l’objet*Application*</span><span class="sxs-lookup"><span data-stu-id="54172-123">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="54172-124">Utilisez l’objet`Application` pour accéder aux objets OneNote tels que **Notebook**, **Section** et **Page**.</span><span class="sxs-lookup"><span data-stu-id="54172-124">Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="54172-125">Grâce à l’API enrichie, vous pouvez exécuter des opérations par lot sur les objets proxy.</span><span class="sxs-lookup"><span data-stu-id="54172-125">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="54172-126">Le flux de base ressemble à ceci:</span><span class="sxs-lookup"><span data-stu-id="54172-126">The basic flow goes something like this:</span></span>

1. <span data-ttu-id="54172-127">Obtenir l’instance de l’application à partir du contexte.</span><span class="sxs-lookup"><span data-stu-id="54172-127">Get the application instance from the context.</span></span>

2. <span data-ttu-id="54172-p107">Créer un proxy qui représente l’objet OneNote que vous souhaitez utiliser. Vous interagissez simultanément avec les objets proxy en lisant et en écrivant leurs propriétés et en appelant leurs méthodes.</span><span class="sxs-lookup"><span data-stu-id="54172-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="54172-p108">Appelez la méthode `load` sur le serveur proxy pour la remplir avec les valeurs de propriété spécifiées dans le paramètre. Cet appel est ajouté à la file d’attente des commandes.</span><span class="sxs-lookup"><span data-stu-id="54172-p108">Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="54172-132">Les appels de méthode à l’API (tels que `context.application.getActiveSection().pages;`) sont également ajoutés à la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="54172-132">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="54172-p109">Appelez la méthode `context.sync` pour exécuter toutes les commandes en attente dans l’ordre dans lequel elles ont été mises en file d’attente. Cela permet de synchroniser l’état entre votre script d’exécution et les objets réels, en récupérant les propriétés des objets OneNote chargés à utiliser dans vos scripts. Vous pouvez utiliser l’objet Promise renvoyé pour créer une chaîne avec les actions supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="54172-p109">Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="54172-136">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="54172-136">For example:</span></span>

```js
function getPagesInSection() {
    OneNote.run(function (context) {

        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Read the id and title of each page.
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId);
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

<span data-ttu-id="54172-137">Vous pouvez déterminer les objets et les opérations OneNote pris en charge dans la [référence de l’API](../reference/overview/onenote-add-ins-javascript-reference.md).</span><span class="sxs-lookup"><span data-stu-id="54172-137">You can find supported OneNote objects and operations in the [API reference](../reference/overview/onenote-add-ins-javascript-reference.md).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="54172-138">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="54172-138">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="54172-139">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="54172-139">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="54172-140">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="54172-140">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="54172-141">Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote, consultez [Ensembles de conditions requises de l’API JavaScript pour OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="54172-141">For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="54172-142">Accès à l’API commune via l’objet*Document*</span><span class="sxs-lookup"><span data-stu-id="54172-142">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="54172-143">Utilisez l’objet `Document` pour accéder à l’API commune, par exemple les méthodes[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) et [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="54172-143">Use the `Document` object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span>


<span data-ttu-id="54172-144">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="54172-144">For example:</span></span>  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

<span data-ttu-id="54172-145">Les compléments OneNote prennent en charge uniquement les API communes suivantes:</span><span class="sxs-lookup"><span data-stu-id="54172-145">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="54172-146">API</span><span class="sxs-lookup"><span data-stu-id="54172-146">API</span></span> | <span data-ttu-id="54172-147">Commentaires</span><span class="sxs-lookup"><span data-stu-id="54172-147">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="54172-148">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="54172-148">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="54172-149">Office.CoercionType.Text`Office.CoercionType.Text` et Office.CoercionType.Matrix`Office.CoercionType.Matrix` uniquement</span><span class="sxs-lookup"><span data-stu-id="54172-149">`Office.CoercionType.Text` and `Office.CoercionType.Matrix` only</span></span> |
| [<span data-ttu-id="54172-150">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="54172-150">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="54172-151">`Office.CoercionType.Text`, `Office.CoercionType.Image`et `Office.CoercionType.Html` uniquement</span><span class="sxs-lookup"><span data-stu-id="54172-151">`Office.CoercionType.Text`, `Office.CoercionType.Image`, and `Office.CoercionType.Html` only</span></span> | 
| [<span data-ttu-id="54172-152">var mySetting = Office.context.document.settings.get(name);</span><span class="sxs-lookup"><span data-stu-id="54172-152">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="54172-153">Les paramètres sont pris en charge par les compléments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="54172-153">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="54172-154">Office.context.document.settings.set(name, value);</span><span class="sxs-lookup"><span data-stu-id="54172-154">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="54172-155">Les paramètres sont pris en charge par les compléments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="54172-155">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="54172-156">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="54172-156">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="54172-157">En règle générale, vous utilisez l’API commune pour effectuer une action qui n’est pas prise en charge dans l’API enrichie.</span><span class="sxs-lookup"><span data-stu-id="54172-157">In general, you use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="54172-158">Pour plus d’informations sur les API communes, voir le [Modèle d’objet API JavaScript communes](../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="54172-158">To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="54172-159">Diagramme du modèle objet OneNote</span><span class="sxs-lookup"><span data-stu-id="54172-159">OneNote object model diagram</span></span> 
<span data-ttu-id="54172-160">Le diagramme suivant représente ce qui est actuellement disponible dans l’API JavaScript de OneNote.</span><span class="sxs-lookup"><span data-stu-id="54172-160">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagramme du modèle objet OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="54172-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="54172-162">See also</span></span>

- [<span data-ttu-id="54172-163">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="54172-163">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="54172-164">Créer votre premier complément OneNote</span><span class="sxs-lookup"><span data-stu-id="54172-164">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="54172-165">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="54172-165">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="54172-166">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="54172-166">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="54172-167">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="54172-167">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
