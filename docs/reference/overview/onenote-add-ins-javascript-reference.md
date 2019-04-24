---
title: Référence de l’API JavaScript pour OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 53b120fbe2bba3967c1b89699daef6bd452b5c24
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450253"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="90d62-102">Référence de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="90d62-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="90d62-103">S’applique à : OneNote Online</span><span class="sxs-lookup"><span data-stu-id="90d62-103">Applies to: OneNote Online</span></span>

<span data-ttu-id="90d62-104">Les liens suivants affichent les objets OneNote de niveau supérieur disponibles dans l’API.</span><span class="sxs-lookup"><span data-stu-id="90d62-104">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="90d62-105">Chaque lien vers la page d’un objet contient une description des propriétés, des événements et des méthodes disponibles sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="90d62-105">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="90d62-106">Cliquez sur ces liens pour en savoir plus.</span><span class="sxs-lookup"><span data-stu-id="90d62-106">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="90d62-107">[Application](/javascript/api/onenote/onenote.application) : Objet de niveau supérieur utilisé pour accéder à tous les objets OneNote globalement adressables, tels que le bloc-notes actif et la section active.</span><span class="sxs-lookup"><span data-stu-id="90d62-107">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="90d62-p102">[Bloc-notes](/javascript/api/onenote/onenote.notebook) : Bloc-notes. Les blocs-notes contiennent des groupes de sections et des sections.</span><span class="sxs-lookup"><span data-stu-id="90d62-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="90d62-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection) : Collection de blocs-notes.</span><span class="sxs-lookup"><span data-stu-id="90d62-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="90d62-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup) : Groupe de sections. Les groupes de sections contiennent des sections et des groupes de sections.</span><span class="sxs-lookup"><span data-stu-id="90d62-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="90d62-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection) : Collection de groupes de sections.</span><span class="sxs-lookup"><span data-stu-id="90d62-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="90d62-p104">[Section](/javascript/api/onenote/onenote.section) : Section. Les sections contiennent des pages.</span><span class="sxs-lookup"><span data-stu-id="90d62-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="90d62-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection) : Collection de sections.</span><span class="sxs-lookup"><span data-stu-id="90d62-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="90d62-p105">[Page](/javascript/api/onenote/onenote.page) : Page. Les pages contiennent des objets PageContent.</span><span class="sxs-lookup"><span data-stu-id="90d62-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="90d62-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection) : Collection de pages.</span><span class="sxs-lookup"><span data-stu-id="90d62-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="90d62-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent) : Zone de niveau supérieur sur une page qui contient des types de contenu tels que des plans ou des images. Un objet PageContent peut être affecté à une position sur la page.</span><span class="sxs-lookup"><span data-stu-id="90d62-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="90d62-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection) : Collection d’objets PageContent qui représente le contenu d’une page.</span><span class="sxs-lookup"><span data-stu-id="90d62-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="90d62-p107">[Outline](/javascript/api/onenote/onenote.outline) : Conteneur pour les objets Paragraph. Un plan est un enfant direct d’un objet PageContent.</span><span class="sxs-lookup"><span data-stu-id="90d62-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="90d62-p108">[Image](/javascript/api/onenote/onenote.image) : Objet Image. Une image peut être un enfant direct d’un objet Paragraph ou PageContent.</span><span class="sxs-lookup"><span data-stu-id="90d62-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="90d62-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph) : Conteneur pour le contenu visible d’une page. Un paragraphe est un enfant direct d’un plan.</span><span class="sxs-lookup"><span data-stu-id="90d62-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="90d62-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection) : Collection d’objets Paragraph dans un plan.</span><span class="sxs-lookup"><span data-stu-id="90d62-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="90d62-130">[Richtext](/javascript/api/onenote/onenote.richtext) : Objet RichText.</span><span class="sxs-lookup"><span data-stu-id="90d62-130">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="90d62-131">[Table](/javascript/api/onenote/onenote.table) : Conteneur pour les objets TableRow.</span><span class="sxs-lookup"><span data-stu-id="90d62-131">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="90d62-132">[TableRow](/javascript/api/onenote/onenote.tablerow) : Conteneur pour les objets TableCell.</span><span class="sxs-lookup"><span data-stu-id="90d62-132">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="90d62-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection) : Collection d’objets TableRow dans un tableau.</span><span class="sxs-lookup"><span data-stu-id="90d62-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="90d62-134">[TableCell](/javascript/api/onenote/onenote.tablecell) : Conteneur pour les objets Paragraph.</span><span class="sxs-lookup"><span data-stu-id="90d62-134">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="90d62-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection) : Collection d’objets TableCell dans un élément TableRow.</span><span class="sxs-lookup"><span data-stu-id="90d62-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="90d62-136">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="90d62-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="90d62-137">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="90d62-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="90d62-138">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="90d62-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="90d62-139">Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote, consultez l’article [Ensembles de conditions requises de l’API JavaScript pour OneNote](../requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="90d62-139">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="90d62-140">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="90d62-140">OneNote JavaScript API reference</span></span>

<span data-ttu-id="90d62-141">Pour en savoir plus sur l’API JavaScript pour OneNote, consultez la [documentation de référence de l’API JavaScript pour OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="90d62-141">For detailed information about the OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="90d62-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="90d62-142">See also</span></span>

- [<span data-ttu-id="90d62-143">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="90d62-143">OneNote JavaScript API programming overview</span></span>](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="90d62-144">Créer votre premier complément OneNote</span><span class="sxs-lookup"><span data-stu-id="90d62-144">Build your first OneNote add-in</span></span>](../../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="90d62-145">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="90d62-145">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="90d62-146">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="90d62-146">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
