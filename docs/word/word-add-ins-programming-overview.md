---
title: Présentation des compléments Word
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 5516e4dc847d4872a12f769530d0a5cb7d779c7c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126805"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="bad95-102">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="bad95-102">Word add-ins overview</span></span>

<span data-ttu-id="bad95-p101">Vous souhaitez créer une solution qui étend les fonctionnalités de Word ? Par exemple, une solution qui assemble automatiquement les documents ? Ou une solution qui relie les données et y accède dans un document Word à partir d’autres sources de données ? Vous pouvez utiliser la plateforme de compléments Office. Elle comprend l’API JavaScript pour Word et l’API JavaScript pour Office, pour développer les clients Word qui s’exécutent sur un ordinateur de bureau Windows, un Mac ou dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="bad95-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="bad95-p102">Les compléments Word font partie des nombreuses options de développement disponibles sur la [plateforme de compléments Office](../overview/office-add-ins.md). Vous pouvez utiliser les commandes de complément pour développer l’interface utilisateur Word et créer des volets Office qui exécutent un code JavaScript pour interagir avec le contenu d’un document Word. Tout code que vous pouvez exécuter dans un navigateur peut s’exécuter dans un complément Word. Les compléments qui interagissent avec le contenu d’un document Word créent des requêtes qui agissent sur des objets Word et synchronisent l’état des objets.</span><span class="sxs-lookup"><span data-stu-id="bad95-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="bad95-p103">Lorsque vous créez votre complément, si vous envisagez de le [publier](../publish/publish.md) dans AppSource, assurez-vous que vous respectez les [stratégies de validation AppSource](/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="bad95-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="bad95-113">La figure suivante montre un exemple d’un complément Word qui s’exécute dans un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="bad95-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="bad95-114">*Figure 1. Complément exécuté dans un volet Office de Word*</span><span class="sxs-lookup"><span data-stu-id="bad95-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Complément exécuté dans un volet Office de Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="bad95-p104">Le complément Word (1) peut envoyer des demandes dans le document Word (2) et utiliser JavaScript pour accéder à l’objet de paragraphe et mettre à jour, supprimer ou déplacer le paragraphe. Par exemple, le code suivant montre comment ajouter une nouvelle phrase à ce paragraphe.</span><span class="sxs-lookup"><span data-stu-id="bad95-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="bad95-p105">Vous pouvez utiliser n’importe quelle technologie de serveur web pour héberger votre complément Word, comme ASP.NET, NodeJS ou Python. Utilisez votre infrastructure côté client préférée (Ember, Backbone, Angular, React), ou utilisez VanillaJS pour développer votre solution et utilisez des services comme Azure pour [authentifier](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) et héberger votre application.</span><span class="sxs-lookup"><span data-stu-id="bad95-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="bad95-p106">Les interfaces API JavaScript pour Word permettent à votre application d’accéder aux objets et aux métadonnées situés dans le document Word. Vous pouvez utiliser ces API pour créer des compléments destinés à :</span><span class="sxs-lookup"><span data-stu-id="bad95-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="bad95-122">Word 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="bad95-122">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="bad95-123">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="bad95-123">Outlook on the web</span></span>
* <span data-ttu-id="bad95-124">Word 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="bad95-124">Word 2016 or later for Mac</span></span>
* <span data-ttu-id="bad95-125">Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="bad95-125">Word on iPad</span></span>

<span data-ttu-id="bad95-p107">Écrivez votre complément une seule fois. Celui-ci s’exécutera dans toutes les versions de Word sur plusieurs plateformes. Pour plus d’informations, voir la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="bad95-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="bad95-128">APIs JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="bad95-128">JavaScript APIs for Word</span></span>

<span data-ttu-id="bad95-129">Vous pouvez utiliser les deux ensembles d’APIs JavaScript pour interagir avec les objets et les métadonnées d’un document Word.</span><span class="sxs-lookup"><span data-stu-id="bad95-129">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="bad95-130">Le premier groupe est l’[API commune](../reference/javascript-api-for-office.md), qui a été introduit dans Office 2013.</span><span class="sxs-lookup"><span data-stu-id="bad95-130">The first is the [Common API](../reference/javascript-api-for-office.md), which was introduced in Office 2013.</span></span> <span data-ttu-id="bad95-131">La plupart des objets dans l’API commune peuvent être utilisés dans des compléments hébergés par deux clients Office ou plus.</span><span class="sxs-lookup"><span data-stu-id="bad95-131">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="bad95-132">Cette API utilise largement les rappels.</span><span class="sxs-lookup"><span data-stu-id="bad95-132">This API uses callbacks extensively.</span></span>

<span data-ttu-id="bad95-p109">Le deuxième est l’[API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md). Il s’agit d’un modèle d’objet fortement typé qui vous permet de créer des compléments Word destinés à Word 2016 sur Mac et Windows. Ce modèle d’objet utilise les promesses et fournit un accès aux objets Word, tels que le [corps](/javascript/api/word/word.body), les [contrôles de contenu](/javascript/api/word/word.contentcontrol), les [images incluses](/javascript/api/word/word.inlinepicture) et les [paragraphes](/javascript/api/word/word.paragraph). L’API JavaScript pour Word inclut les définitions TypeScript et les fichiers vsdoc pour vous permettre d’obtenir des conseils concernant votre code dans votre IDE.</span><span class="sxs-lookup"><span data-stu-id="bad95-p109">The second is the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="bad95-p110">Actuellement, tous les clients Word prennent en charge l’API JavaScript partagé pour Office, et la plupart des clients prennent en charge l’API JavaScript pour Word. Pour plus d’informations sur les clients pris en charge, voir la [Documentation de référence de l’API](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span><span class="sxs-lookup"><span data-stu-id="bad95-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="bad95-p111">Nous vous recommandons de démarrer avec l’API JavaScript pour Word car le modèle d’objet est plus facile à utiliser. Utilisez l’API JavaScript pour Word pour :</span><span class="sxs-lookup"><span data-stu-id="bad95-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="bad95-141">Accéder aux objets d’un document Word.</span><span class="sxs-lookup"><span data-stu-id="bad95-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="bad95-142">Utilisez l’API JavaScript partagé pour Office pour :</span><span class="sxs-lookup"><span data-stu-id="bad95-142">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="bad95-143">Cibler Word 2013.</span><span class="sxs-lookup"><span data-stu-id="bad95-143">Target Word 2013.</span></span>
* <span data-ttu-id="bad95-144">Effectuer des actions initiales pour l’application.</span><span class="sxs-lookup"><span data-stu-id="bad95-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="bad95-145">Vérifier l’ensemble de conditions requises pris en charge.</span><span class="sxs-lookup"><span data-stu-id="bad95-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="bad95-146">Accéder aux métadonnées, aux paramètres et aux informations de l’environnement du document.</span><span class="sxs-lookup"><span data-stu-id="bad95-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="bad95-147">Établir des liaisons avec des sections d’un document et capturer les événements.</span><span class="sxs-lookup"><span data-stu-id="bad95-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="bad95-148">Utiliser des parties XML personnalisées.</span><span class="sxs-lookup"><span data-stu-id="bad95-148">Use custom XML parts.</span></span>
* <span data-ttu-id="bad95-149">Ouvrir une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="bad95-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="bad95-150">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="bad95-150">Next steps</span></span>

<span data-ttu-id="bad95-p112">Prêt à créer votre premier complément Word ? Consultez la page [Création de votre premier complément Word](word-add-ins.md). Utilisez le [manifeste de complément](../develop/add-in-manifests.md) pour décrire l’emplacement d’hébergement de votre complément et son affichage, et définir des autorisations et d’autres informations.</span><span class="sxs-lookup"><span data-stu-id="bad95-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](../develop/add-in-manifests.md). Use the add-in manifest to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="bad95-154">Pour savoir comment concevoir un complément Word de qualité qui offre une expérience intéressante aux utilisateurs, consultez les [recommandations de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="bad95-154">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="bad95-155">Une fois le développement de votre complément terminé, vous pouvez le [publier](../publish/publish.md) sur un partage réseau, dans un catalogue d’applications ou dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="bad95-155">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="bad95-156">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bad95-156">See also</span></span>

* [<span data-ttu-id="bad95-157">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="bad95-157">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="bad95-158">Référence d’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="bad95-158">Word JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
