---
title: Vue d?ensemble de la programmation de l?API JavaScript de OneNote
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: aded1210abc11a80c6200a207d3896df8ef4218b
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="87042-102">Vue d?ensemble de la programmation de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="87042-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="87042-103">OneNote pr?sente une API JavaScript pour les compl?ments OneNote Online. Vous pouvez cr?er des compl?ments de volet de t?ches et de contenu, ainsi que des commandes de compl?ment qui interagissent avec les objets OneNote et se connectent ? des services web ou ? d?autres ressources bas?es sur le web.</span><span class="sxs-lookup"><span data-stu-id="87042-103">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="87042-p101">Si vous pr?voyez de [publier](../publish/publish.md) votre compl?ment sur AppSource et de le rendre disponible dans l?exp?rience Office, assurez-vous que vous respectez les [strat?gies de validation AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Par exemple, pour r?ussir la validation, votre compl?ment doit fonctionner sur toutes les plateformes prenant en charge les m?thodes d?finies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative ? la disponibilit? des compl?ments Office sur les plateformes et les h?tes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="87042-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="87042-106">Composants d?un compl?ment Office</span><span class="sxs-lookup"><span data-stu-id="87042-106">Components of an Office Add-in</span></span>

<span data-ttu-id="87042-107">Les compl?ments sont constitu?s de deux composants de base :</span><span class="sxs-lookup"><span data-stu-id="87042-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="87042-p102">Une **application web** comportant une page web et les fichiers CSS, JavaScript ou autres requis. Ces fichiers sont h?berg?s sur un serveur web ou un service d?h?bergement web, tel que Microsoft Azure. Dans OneNote Online, l?application web s?affiche dans un contr?le de navigateur ou un iFrame.</span><span class="sxs-lookup"><span data-stu-id="87042-p102">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote Online, the web application displays in a browser control or iframe.</span></span>
    
- <span data-ttu-id="87042-p103">Un **manifeste XML** sp?cifiant l?URL de la page web du compl?ment, ainsi que les conditions d?acc?s, les param?tres et fonctionnalit?s du compl?ment. Ce fichier est stock? sur le client. Les compl?ments OneNote utilisent le m?me format de [manifeste](../develop/add-in-manifests.md) que les autres compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="87042-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="87042-114">**Compl?ment pour Office = manifeste + page web**</span><span class="sxs-lookup"><span data-stu-id="87042-114">**Office Add-in = Manifest + Webpage**</span></span>

![Un compl?ment Office se compose d?un manifeste et d?une page web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="87042-116">Utilisation de l?API JavaScript</span><span class="sxs-lookup"><span data-stu-id="87042-116">Using the JavaScript API</span></span>

<span data-ttu-id="87042-p104">Les compl?ments utilisent le contexte d?ex?cution de l?application h?te pour acc?der ? l?API JavaScript. L?API comporte deux couches :</span><span class="sxs-lookup"><span data-stu-id="87042-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="87042-119">Une **API enrichie** pour les op?rations sp?cifiques de OneNote, accessible via l?objet **Application**.</span><span class="sxs-lookup"><span data-stu-id="87042-119">A **rich API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="87042-120">Une **API commune** qui est partag?e entre les applications Office, accessible via l?objet **Document**.</span><span class="sxs-lookup"><span data-stu-id="87042-120">A **common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-rich-api-through-the-application-object"></a><span data-ttu-id="87042-121">Acc?s ? l?API enrichie via l?objet *Application*</span><span class="sxs-lookup"><span data-stu-id="87042-121">Accessing the rich API through the *Application* object</span></span>

<span data-ttu-id="87042-p105">Utilisez l?objet **Application** pour acc?der aux objets OneNote tels que **Notebook**, **Section** et **Page**. Gr?ce ? l?API enrichie, vous pouvez ex?cuter des op?rations par lot sur les objets proxy. Le flux de base ressemble ? ceci :</span><span class="sxs-lookup"><span data-stu-id="87042-p105">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With rich APIs, you run batch operations on proxy objects. The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="87042-125">Obtenir l?instance de l?application ? partir du contexte.</span><span class="sxs-lookup"><span data-stu-id="87042-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="87042-p106">Cr?er un proxy qui repr?sente l?objet OneNote que vous souhaitez utiliser. Vous interagissez simultan?ment avec les objets proxy en lisant et en ?crivant leurs propri?t?s et en appelant leurs m?thodes.</span><span class="sxs-lookup"><span data-stu-id="87042-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span> 

3. <span data-ttu-id="87042-p107">Appelez la m?thode **load** sur le serveur proxy pour la remplir avec les valeurs de propri?t? sp?cifi?es dans le param?tre. Cet appel est ajout? ? la file d?attente des commandes.</span><span class="sxs-lookup"><span data-stu-id="87042-p107">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="87042-130">Les appels de m?thode ? l?API (tels que `context.application.getActiveSection().pages;`) sont ?galement ajout?s ? la file d?attente.</span><span class="sxs-lookup"><span data-stu-id="87042-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="87042-p108">Appelez la m?thode **context.sync** pour ex?cuter toutes les commandes en attente dans l?ordre dans lequel elles ont ?t? mises en file d?attente. Cela permet de synchroniser l??tat entre votre script d?ex?cution et les objets r?els, en r?cup?rant les propri?t?s des objets OneNote charg?s ? utiliser dans vos scripts. Vous pouvez utiliser l?objet Promise renvoy? pour cr?er une cha?ne avec les actions suppl?mentaires.</span><span class="sxs-lookup"><span data-stu-id="87042-p108">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="87042-134">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="87042-134">For example:</span></span> 

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

<span data-ttu-id="87042-135">Vous pouvez d?terminer les objets et les op?rations OneNote pris en charge dans la [r?f?rence de l?API](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="87042-135">You can find supported OneNote objects and operations in the [API reference](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="87042-136">Acc?s ? l?API commune via l?objet *Document*</span><span class="sxs-lookup"><span data-stu-id="87042-136">Accessing the common API through the *Document* object</span></span>

<span data-ttu-id="87042-137">Utilisez l?objet **Document** pour acc?der ? l?API commune, par exemple les m?thodes [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) et [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync).</span><span class="sxs-lookup"><span data-stu-id="87042-137">Use the **Document** object to access the common API, such as the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) and [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) methods.</span></span> 

<span data-ttu-id="87042-138">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="87042-138">For example:</span></span>  

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
<span data-ttu-id="87042-139">Les compl?ments OneNote prennent en charge uniquement les API communes suivantes :</span><span class="sxs-lookup"><span data-stu-id="87042-139">OneNote add-ins support only the following common APIs:</span></span>

| <span data-ttu-id="87042-140">API</span><span class="sxs-lookup"><span data-stu-id="87042-140">API</span></span> | <span data-ttu-id="87042-141">Commentaires</span><span class="sxs-lookup"><span data-stu-id="87042-141">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="87042-142">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="87042-142">Office.context.document.getSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) | <span data-ttu-id="87042-143">**Office.CoercionType.Text** et **Office.CoercionType.Matrix** uniquement</span><span class="sxs-lookup"><span data-stu-id="87042-143">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="87042-144">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="87042-144">Office.context.document.setSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) | <span data-ttu-id="87042-145">**Office.CoercionType.Text**, **Office.CoercionType.Image** et **Office.CoercionType.Html** uniquement</span><span class="sxs-lookup"><span data-stu-id="87042-145">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="87042-146">var mySetting = Office.context.document.settings.get(name);</span><span class="sxs-lookup"><span data-stu-id="87042-146">var mySetting = Office.context.document.settings.get(name);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.get) | <span data-ttu-id="87042-147">Les param?tres sont pris en charge par les compl?ments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="87042-147">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="87042-148">Office.context.document.settings.set(name, value);</span><span class="sxs-lookup"><span data-stu-id="87042-148">Office.context.document.settings.set(name, value);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.set) | <span data-ttu-id="87042-149">Les param?tres sont pris en charge par les compl?ments de contenu uniquement</span><span class="sxs-lookup"><span data-stu-id="87042-149">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="87042-150">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="87042-150">Office.EventType.DocumentSelectionChanged</span></span>](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

<span data-ttu-id="87042-p109">En r?gle g?n?rale, vous utilisez uniquement l?API commune pour effectuer une action qui n?est pas prise en charge dans l?API enrichie. Pour en savoir plus sur l?utilisation de l?API commune, voir la [documentation](../overview/office-add-ins.md) et les [r?f?rences](https://dev.office.com/reference/add-ins/javascript-api-for-office) concernant les compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="87042-p109">In general, you only use the common API to do something that isn't supported in the rich API. To learn more about using the common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](https://dev.office.com/reference/add-ins/javascript-api-for-office).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="87042-153">Diagramme du mod?le objet OneNote</span><span class="sxs-lookup"><span data-stu-id="87042-153">OneNote object model diagram</span></span> 
<span data-ttu-id="87042-154">Le diagramme suivant repr?sente ce qui est actuellement disponible dans l?API JavaScript de OneNote.</span><span class="sxs-lookup"><span data-stu-id="87042-154">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagramme du mod?le objet OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="87042-156">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="87042-156">See also</span></span>

- [<span data-ttu-id="87042-157">Cr?er votre premier compl?ment OneNote</span><span class="sxs-lookup"><span data-stu-id="87042-157">Build your first OneNote add-in</span></span>](onenote-add-ins-getting-started.md)
- [<span data-ttu-id="87042-158">R?f?rence de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="87042-158">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="87042-159">Exemple de grille d??valuation</span><span class="sxs-lookup"><span data-stu-id="87042-159">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="87042-160">Vue d?ensemble de la plateforme des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="87042-160">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
