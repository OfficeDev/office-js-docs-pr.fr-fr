---
title: Présentation des compléments Word
description: Découvrez les concepts de base des Compléments Word.
ms.date: 07/28/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: b531ec5c2a5fa1e3e9366f703a57e815a5711b5a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293071"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="c5c06-103">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="c5c06-103">Word add-ins overview</span></span>

<span data-ttu-id="c5c06-p101">Vous souhaitez créer une solution qui étend les fonctionnalités de Word ? Par exemple, une solution qui assemble automatiquement les documents ? Ou une solution qui relie les données et y accède dans un document Word à partir d’autres sources de données ? Vous pouvez utiliser la plateforme de compléments Office. Elle comprend l’API JavaScript pour Word et l’API Office JavaScript, pour développer les clients Word qui s’exécutent sur un ordinateur de bureau Windows, un Mac ou dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="c5c06-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="c5c06-p102">Les compléments Word font partie des nombreuses options de développement disponibles sur la [plateforme de compléments Office](../overview/office-add-ins.md). Vous pouvez utiliser les commandes de complément pour développer l’interface utilisateur Word et créer des volets Office qui exécutent un code JavaScript pour interagir avec le contenu d’un document Word. Tout code que vous pouvez exécuter dans un navigateur peut s’exécuter dans un complément Word. Les compléments qui interagissent avec le contenu d’un document Word créent des requêtes qui agissent sur des objets Word et synchronisent l’état des objets.</span><span class="sxs-lookup"><span data-stu-id="c5c06-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="c5c06-112">La figure suivante montre un exemple d’un complément Word qui s’exécute dans un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="c5c06-112">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="c5c06-113">*Figure 1. Complément exécuté dans un volet Office de Word*</span><span class="sxs-lookup"><span data-stu-id="c5c06-113">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Complément exécuté dans un volet Office de Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="c5c06-p103">Le complément Word (1) peut envoyer des demandes dans le document Word (2) et utiliser JavaScript pour accéder à l’objet de paragraphe et mettre à jour, supprimer ou déplacer le paragraphe. Par exemple, le code suivant montre comment ajouter une nouvelle phrase à ce paragraphe.</span><span class="sxs-lookup"><span data-stu-id="c5c06-p103">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

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

<span data-ttu-id="c5c06-p104">Vous pouvez utiliser n’importe quelle technologie de serveur web pour héberger votre complément Word, comme ASP.NET, NodeJS ou Python. Utilisez votre infrastructure côté client préférée (Ember, Backbone, Angular, React), ou utilisez VanillaJS pour développer votre solution et utilisez des services comme Azure pour [authentifier](../develop/overview-authn-authz.md) et héberger votre application.</span><span class="sxs-lookup"><span data-stu-id="c5c06-p104">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.</span></span>

<span data-ttu-id="c5c06-p105">Les interfaces API JavaScript pour Word permettent à votre application d’accéder aux objets et aux métadonnées situés dans le document Word. Vous pouvez utiliser ces API pour créer des compléments destinés à :</span><span class="sxs-lookup"><span data-stu-id="c5c06-p105">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="c5c06-121">Word 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="c5c06-121">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="c5c06-122">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="c5c06-122">Word on the web</span></span>
* <span data-ttu-id="c5c06-123">Word 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="c5c06-123">Word 2016 or later on Mac</span></span>
* <span data-ttu-id="c5c06-124">Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="c5c06-124">Word on iPad</span></span>

<span data-ttu-id="c5c06-p106">Écrivez votre complément une seule fois. Celui-ci s’exécutera dans toutes les versions de Word sur plusieurs plateformes. Pour plus d’informations, voir [Disponibilité des compléments Office sur les plateformes et les applications clientes](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="c5c06-p106">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="c5c06-127">APIs JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="c5c06-127">JavaScript APIs for Word</span></span>

<span data-ttu-id="c5c06-128">Vous pouvez utiliser les deux ensembles d’APIs JavaScript pour interagir avec les objets et les métadonnées d’un document Word.</span><span class="sxs-lookup"><span data-stu-id="c5c06-128">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="c5c06-129">Le premier groupe est l’[API commune](/javascript/api/office), qui a été introduit dans Office 2013.</span><span class="sxs-lookup"><span data-stu-id="c5c06-129">The first is the [Common API](/javascript/api/office), which was introduced in Office 2013.</span></span> <span data-ttu-id="c5c06-130">La plupart des objets dans l’API commune peuvent être utilisés dans des compléments hébergés par deux clients Office ou plus.</span><span class="sxs-lookup"><span data-stu-id="c5c06-130">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="c5c06-131">Cette API utilise largement les rappels.</span><span class="sxs-lookup"><span data-stu-id="c5c06-131">This API uses callbacks extensively.</span></span>

<span data-ttu-id="c5c06-p108">Le deuxième est l’[API JavaScript pour Word](/javascript/api/word) qui est un [modèle d’API spécifique à l’application](../develop/application-specific-api-model.md) introduit avec Word 2016. Il s’agit d’un modèle objet fortement typé qui vous permet de créer des compléments Word destinés à Word 2016 sur Mac et Windows. Ce modèle objet utilise les promesses et fournit un accès aux objets Word, tels que le [corps](/javascript/api/word/word.body), les [contrôles de contenu](/javascript/api/word/word.contentcontrol), les [images incorporées](/javascript/api/word/word.inlinepicture) et les [paragraphes](/javascript/api/word/word.paragraph). L’API JavaScript pour Word inclut des définitions TypeScript et des fichiers vsdoc pour vous permettre d’obtenir des conseils concernant votre code dans votre environnement de développement intégré (IDE).</span><span class="sxs-lookup"><span data-stu-id="c5c06-p108">The second is the [Word JavaScript API](/javascript/api/word). This is a [application-specific API model](../develop/application-specific-api-model.md) that was introduced with Word 2016. It's a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="c5c06-137">Actuellement, tous les clients Word prennent en charge l’API JavaScript Office partagée, et la plupart des clients prennent en charge l’API JavaScript pour Word.</span><span class="sxs-lookup"><span data-stu-id="c5c06-137">Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API.</span></span> <span data-ttu-id="c5c06-138">Pour plus d’informations sur les clients pris en charge, voir [Disponibilité des applications clientes Office et des plateformes pour les compléments Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="c5c06-138">For details about supported clients, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

<span data-ttu-id="c5c06-p110">Nous vous recommandons de démarrer avec l’API JavaScript pour Word car le modèle d’objet est plus facile à utiliser. Utilisez l’API JavaScript pour Word pour :</span><span class="sxs-lookup"><span data-stu-id="c5c06-p110">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="c5c06-141">Accéder aux objets d’un document Word.</span><span class="sxs-lookup"><span data-stu-id="c5c06-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="c5c06-142">Utilisez l’API Office JavaScript partagée pour :</span><span class="sxs-lookup"><span data-stu-id="c5c06-142">Use the shared Office JavaScript API when you need to:</span></span>

* <span data-ttu-id="c5c06-143">Cibler Word 2013.</span><span class="sxs-lookup"><span data-stu-id="c5c06-143">Target Word 2013.</span></span>
* <span data-ttu-id="c5c06-144">Effectuer des actions initiales pour l’application.</span><span class="sxs-lookup"><span data-stu-id="c5c06-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="c5c06-145">Vérifier l’ensemble de conditions requises pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c5c06-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="c5c06-146">Accéder aux métadonnées, aux paramètres et aux informations de l’environnement du document.</span><span class="sxs-lookup"><span data-stu-id="c5c06-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="c5c06-147">Établir des liaisons avec des sections d’un document et capturer les événements.</span><span class="sxs-lookup"><span data-stu-id="c5c06-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="c5c06-148">Utiliser des parties XML personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c5c06-148">Use custom XML parts.</span></span>
* <span data-ttu-id="c5c06-149">Ouvrir une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="c5c06-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c5c06-150">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="c5c06-150">Next steps</span></span>

<span data-ttu-id="c5c06-p111">Prêt à créer votre premier complément Word ? Consultez la page [Création de votre premier complément Word](word-add-ins.md). Utilisez le [manifeste de complément](../develop/add-in-manifests.md) pour décrire l’emplacement d’hébergement de votre complément et son affichage, et définir des autorisations et d’autres informations.</span><span class="sxs-lookup"><span data-stu-id="c5c06-p111">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="c5c06-154">Pour savoir comment concevoir un complément Word de qualité qui offre une expérience intéressante aux utilisateurs, consultez les [recommandations de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="c5c06-154">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="c5c06-155">Une fois le développement de votre complément terminé, vous pouvez le [publier](../publish/publish.md) sur un partage réseau, dans un catalogue d’applications ou dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="c5c06-155">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="c5c06-156">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c5c06-156">See also</span></span>

* [<span data-ttu-id="c5c06-157">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="c5c06-157">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="c5c06-158">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c5c06-158">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="c5c06-159">Référence sur l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="c5c06-159">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)
