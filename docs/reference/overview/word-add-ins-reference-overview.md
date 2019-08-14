---
title: Présentation des API JavaScript pour Word
description: ''
ms.date: 07/05/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: fbc9e8293642d1ab8edf32d568a5dab7ef77a8f0
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575624"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="6b6af-102">Présentation des APIs JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="6b6af-102">Word JavaScript API overview</span></span>

<span data-ttu-id="6b6af-103">Un complément Word interagit avec des objets dans Word via l’API JavaScript pour Office, qui inclut deux modèles objet JavaScript :</span><span class="sxs-lookup"><span data-stu-id="6b6af-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="6b6af-104">**API JavaScript pour Word** : inclut dans Office 2016, l’[API JavaScript pour Word](/javascript/api/word) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des objets et à des métadonnées dans un document Word.</span><span class="sxs-lookup"><span data-stu-id="6b6af-104">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="6b6af-105">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) permettent d’accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="6b6af-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="6b6af-106">Cette section de la documentation traite de l’API JavaScript pour Word, que vous allez utiliser pour développer la majorité des fonctionnalités des compléments utilisés dans Word sur le web ou dans Word 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="6b6af-106">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="6b6af-107">Pour plus d’informations sur les API communes, voir [API JavaScript pour Office](../javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="6b6af-107">For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](../javascript-api-for-office.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="6b6af-108">Découvrir les concepts de programmation</span><span class="sxs-lookup"><span data-stu-id="6b6af-108">Learn programming concepts</span></span>

<span data-ttu-id="6b6af-109">Pour plus d’informations sur les concepts de programmation essentiels, voir [Concepts fondamentaux de programmation avec l’API JavaScript pour Word](../../word/word-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="6b6af-109">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="6b6af-110">En savoir plus sur les fonctionnalités des API</span><span class="sxs-lookup"><span data-stu-id="6b6af-110">Learn about API capabilities</span></span>

<span data-ttu-id="6b6af-111">Reportez-vous aux autres articles présents dans cette section de la documentation pour découvrir comment [obtenir l’ensemble d’un document à partir d’un complément](../../word/get-the-whole-document-from-an-add-in-for-word.md), [utiliser les options de recherche pour trouver du texte dans votre complément Word](../../word/search-option-guidance.md), etc.</span><span class="sxs-lookup"><span data-stu-id="6b6af-111">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="6b6af-112">Reportez-vous à la table des matières pour obtenir la liste complète des articles disponibles.</span><span class="sxs-lookup"><span data-stu-id="6b6af-112">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="6b6af-113">Pour apprendre à utiliser l’API JavaScript pour Word afin d’accéder à des objets dans Word, suivez le [didacticiel sur les compléments Word](../../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="6b6af-113">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="6b6af-114">Pour en savoir plus sur le modèle objet de l’API JavaScript pour Word, consultez la [documentation de référence sur l’API JavaScript pour Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="6b6af-114">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="6b6af-115">Tester les exemples de code dans Script Lab</span><span class="sxs-lookup"><span data-stu-id="6b6af-115">Try out code samples in Script Lab</span></span>

<span data-ttu-id="6b6af-116">Utilisez [Script Lab](../../overview/explore-with-script-lab.md) pour commencer rapidement avec une collection d’exemples intégrés qui montrent comment accomplir des tâches avec l’API.</span><span class="sxs-lookup"><span data-stu-id="6b6af-116">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="6b6af-117">Vous pouvez exécuter les exemples dans Script Lab pour afficher instantanément le résultat dans le volet Office ou le document, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="6b6af-117">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="6b6af-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6b6af-118">See also</span></span>

- [<span data-ttu-id="6b6af-119">Documentation sur les compléments Word</span><span class="sxs-lookup"><span data-stu-id="6b6af-119">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="6b6af-120">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="6b6af-120">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="6b6af-121">Référence sur l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="6b6af-121">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="6b6af-122">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="6b6af-122">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="6b6af-123">Spécifications ouvertes des API</span><span class="sxs-lookup"><span data-stu-id="6b6af-123">API open specifications</span></span>](../openspec/openspec.md)
