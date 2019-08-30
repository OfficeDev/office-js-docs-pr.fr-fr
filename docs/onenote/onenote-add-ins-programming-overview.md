---
title: Vue d’ensemble de la programmation de l’API JavaScript de OneNote
description: ''
ms.date: 07/05/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Priority
ms.openlocfilehash: 214a74c29d748b0354c4dc75f85c3ada442214d5
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672802"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="5c372-102">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="5c372-103">OneNote présente une API JavaScript pour les compléments OneNote sur le web.</span><span class="sxs-lookup"><span data-stu-id="5c372-103">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="5c372-104">Vous pouvez créer des compléments de volet Office et de contenu, ainsi que des commandes de complément qui interagissent avec les objets OneNote et se connectent à des services web ou à d’autres ressources basées sur le web.</span><span class="sxs-lookup"><span data-stu-id="5c372-104">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="5c372-p102">Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="5c372-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="5c372-107">Composants d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="5c372-107">Components of an Office Add-in</span></span>

<span data-ttu-id="5c372-108">Les compléments sont constitués de deux composants de base :</span><span class="sxs-lookup"><span data-stu-id="5c372-108">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="5c372-109">Une **application web** comportant une page web et les fichiers CSS, JavaScript ou autres requis.</span><span class="sxs-lookup"><span data-stu-id="5c372-109">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="5c372-110">Ces fichiers sont hébergés sur un serveur web ou un service d’hébergement web, tel que Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="5c372-110">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="5c372-111">Dans OneNote sur le web, l’application web s’affiche dans un contrôle de navigateur ou un iFrame.</span><span class="sxs-lookup"><span data-stu-id="5c372-111">In OneNote Online, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="5c372-p104">Un **manifeste XML** spécifiant l’URL de la page web du complément, ainsi que les conditions d’accès, les paramètres et fonctionnalités du complément. Ce fichier est stocké sur le client. Les compléments OneNote utilisent le même format de [manifeste](../develop/add-in-manifests.md) que les autres compléments Office.</span><span class="sxs-lookup"><span data-stu-id="5c372-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="5c372-115">**Complément pour Office = manifeste + page web**</span><span class="sxs-lookup"><span data-stu-id="5c372-115">**Office Add-in = Manifest + Webpage**</span></span>

![Un complément Office se compose d’un manifeste et d’une page web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="5c372-117">Utilisation de l’API JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c372-117">Using the JavaScript API</span></span>

<span data-ttu-id="5c372-p105">Les compléments utilisent le contexte d’exécution de l’application hôte pour accéder à l’API JavaScript. L’API comporte deux couches:</span><span class="sxs-lookup"><span data-stu-id="5c372-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="5c372-120">Une **API enrichie** pour les opérations spécifiques de OneNote, accessible via l’objet**Application**.</span><span class="sxs-lookup"><span data-stu-id="5c372-120">A **host-specific API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="5c372-121">Une**API commune** qui est partagée entre les applications Office, accessible via l’objet **Document**.</span><span class="sxs-lookup"><span data-stu-id="5c372-121">A **Common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="5c372-122">Accès à l’API enrichie via l’objet*Application*</span><span class="sxs-lookup"><span data-stu-id="5c372-122">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="5c372-123">Utilisez l’objet**Application** pour accéder aux objets OneNote tels que **Notebook**, **Section** et **Page**.</span><span class="sxs-lookup"><span data-stu-id="5c372-123">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="5c372-124">Grâce à l’API enrichie, vous pouvez exécuter des opérations par lot sur les objets proxy.</span><span class="sxs-lookup"><span data-stu-id="5c372-124">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="5c372-125">Le flux de base ressemble à ceci:</span><span class="sxs-lookup"><span data-stu-id="5c372-125">The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="5c372-126">Obtenir l’instance de l’application à partir du contexte.</span><span class="sxs-lookup"><span data-stu-id="5c372-126">Get the application instance from the context.</span></span>

2. <span data-ttu-id="5c372-p107">Créer un proxy qui représente l’objet OneNote que vous souhaitez utiliser. Vous interagissez simultanément avec les objets proxy en lisant et en écrivant leurs propriétés et en appelant leurs méthodes.</span><span class="sxs-lookup"><span data-stu-id="5c372-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="5c372-p108">Appelez la méthode **load** sur le serveur proxy pour la remplir avec les valeurs de propriété spécifiées dans le paramètre. Cet appel est ajouté à la file d’attente des commandes.</span><span class="sxs-lookup"><span data-stu-id="5c372-p108">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="5c372-131">Les appels de méthode à l’API (tels que `context.application.getActiveSection().pages;`) sont également ajoutés à la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="5c372-131">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="5c372-p109">Appelez la méthode **context.sync** pour exécuter toutes les commandes en attente dans l’ordre dans lequel elles ont été mises en file d’attente. Cela permet de synchroniser l’état entre votre script d’exécution et les objets réels, en récupérant les propriétés des objets OneNote chargés à utiliser dans vos scripts. Vous pouvez utiliser l’objet Promise renvoyé pour créer une chaîne avec les actions supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="5c372-p109">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="5c372-135">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="5c372-135">For example:</span></span>

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

<span data-ttu-id="5c372-136">Vous pouvez déterminer les objets et les opérations OneNote pris en charge dans la [référence de l’API](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="5c372-136">You can find supported OneNote objects and operations in the [API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="5c372-137">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-137">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="5c372-138">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="5c372-138">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="5c372-139">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="5c372-139">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="5c372-140">Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote, consultez [Ensembles de conditions requises de l’API JavaScript pour OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5c372-140">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="5c372-141">Accès à l’API commune via l’objet*Document*</span><span class="sxs-lookup"><span data-stu-id="5c372-141">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="5c372-142">Utilisez l’objet **Document** pour accéder à l’API commune, par exemple les méthodes[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) et [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="5c372-142">Use the **Document** object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span> 


<span data-ttu-id="5c372-143">Par exemple:</span><span class="sxs-lookup"><span data-stu-id="5c372-143">For example:</span></span>  

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

<span data-ttu-id="5c372-144">Les compléments OneNote prennent en charge uniquement les API communes suivantes:</span><span class="sxs-lookup"><span data-stu-id="5c372-144">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="5c372-145">API</span><span class="sxs-lookup"><span data-stu-id="5c372-145">API</span></span> | <span data-ttu-id="5c372-146">Commentaires</span><span class="sxs-lookup"><span data-stu-id="5c372-146">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="5c372-147">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5c372-147">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="5c372-148">**Office.CoercionType.Text** et **Office.CoercionType.Matrix** uniquement</span><span class="sxs-lookup"><span data-stu-id="5c372-148">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="5c372-149">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5c372-149">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="5c372-150">**Office.CoercionType.Text**, **Office.CoercionType.Image** et **Office.CoercionType.Html** uniquement</span><span class="sxs-lookup"><span data-stu-id="5c372-150">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="5c372-151">var mySetting = Office.context.document.settings.get(name);</span><span class="sxs-lookup"><span data-stu-id="5c372-151">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="5c372-152">Les paramètres sont pris en charge par les compléments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="5c372-152">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="5c372-153">Office.context.document.settings.set(name, value);</span><span class="sxs-lookup"><span data-stu-id="5c372-153">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="5c372-154">Les paramètres sont pris en charge par les compléments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="5c372-154">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="5c372-155">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="5c372-155">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="5c372-156">En règle générale, vous utilisez uniquement l’API commune pour effectuer une action qui n’est pas prise en charge dans l’API enrichie.</span><span class="sxs-lookup"><span data-stu-id="5c372-156">In general, you only use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="5c372-157">Pour en savoir plus sur l’utilisation de l’API commune, reportez-vous à la[documentation](../overview/office-add-ins.md) et les[références](../reference/javascript-api-for-office.md)concernant les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="5c372-157">To learn more about using the Common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](../reference/javascript-api-for-office.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="5c372-158">Diagramme du modèle objet OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-158">OneNote object model diagram</span></span> 
<span data-ttu-id="5c372-159">Le diagramme suivant représente ce qui est actuellement disponible dans l’API JavaScript de OneNote.</span><span class="sxs-lookup"><span data-stu-id="5c372-159">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagramme du modèle objet OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="5c372-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5c372-161">See also</span></span>

- [<span data-ttu-id="5c372-162">Documentation sur les compléments OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-162">OneNote add-ins documentation</span></span>](index.md)
- [<span data-ttu-id="5c372-163">Créer votre premier complément OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-163">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="5c372-164">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="5c372-164">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="5c372-165">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="5c372-165">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="5c372-166">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="5c372-166">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
