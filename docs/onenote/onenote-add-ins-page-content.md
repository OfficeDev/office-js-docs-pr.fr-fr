---
title: Utiliser du contenu de page OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d05f251a798a7670983187bfa4c80140b30f6147
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="eccf0-102">Utiliser du contenu de page OneNote</span><span class="sxs-lookup"><span data-stu-id="eccf0-102">Work with OneNote page content</span></span> 

<span data-ttu-id="eccf0-103">Dans l?API JavaScript des compl?ments OneNote, le contenu de page est repr?sent? par le mod?le objet suivant.</span><span class="sxs-lookup"><span data-stu-id="eccf0-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagramme du mod?le objet de page OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="eccf0-105">Un objet Page contient une collection d?objets PageContent.</span><span class="sxs-lookup"><span data-stu-id="eccf0-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="eccf0-106">Un objet PageContent contient un type de contenu de Outline, Image ou Other.</span><span class="sxs-lookup"><span data-stu-id="eccf0-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="eccf0-107">Un objet Outline contient une collection d?objets Paragraph.</span><span class="sxs-lookup"><span data-stu-id="eccf0-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="eccf0-108">Un objet Paragraph contient un type de contenu RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="eccf0-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="eccf0-109">Pour cr?er une page OneNote vide, utilisez l?une des m?thodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="eccf0-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="eccf0-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="eccf0-110">Section.addPage</span></span>](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [<span data-ttu-id="eccf0-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="eccf0-111">Page.insertPageAsSibling</span></span>](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

<span data-ttu-id="eccf0-112">Utilisez ensuite les m?thodes dans les objets suivants pour travailler avec le contenu de la page, comme Page.addOutline et Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="eccf0-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="eccf0-113">Page</span><span class="sxs-lookup"><span data-stu-id="eccf0-113">Page</span></span>](https://dev.office.com/reference/add-ins/onenote/page)
- [<span data-ttu-id="eccf0-114">Outline</span><span class="sxs-lookup"><span data-stu-id="eccf0-114">Outline</span></span>](https://dev.office.com/reference/add-ins/onenote/outline)
- [<span data-ttu-id="eccf0-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="eccf0-115">Paragraph</span></span>](https://dev.office.com/reference/add-ins/onenote/paragraph)

<span data-ttu-id="eccf0-p101">Le contenu et la structure d?une page OneNote sont repr?sent?s par du code HTML. Seul un sous-ensemble de code HTML est pris en charge pour cr?er ou mettre ? jour du contenu de page, comme d?crit ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="eccf0-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="eccf0-118">HTML pris en charge</span><span class="sxs-lookup"><span data-stu-id="eccf0-118">Supported HTML</span></span>

<span data-ttu-id="eccf0-119">L?API JavaScript des compl?ments OneNote prend en charge le code HTML suivant pour cr?er et mettre ? jour du contenu de page :</span><span class="sxs-lookup"><span data-stu-id="eccf0-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="eccf0-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="eccf0-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="eccf0-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="eccf0-121"></span></span> 
- <span data-ttu-id="eccf0-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="eccf0-122"></span></span>
- <span data-ttu-id="eccf0-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="eccf0-123"></span></span>
- <span data-ttu-id="eccf0-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="eccf0-124"></span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="eccf0-125">Acc?s au contenu de la page</span><span class="sxs-lookup"><span data-stu-id="eccf0-125">Accessing page contents</span></span>

<span data-ttu-id="eccf0-p102">Vous pouvez uniquement acc?der au *contenu de la page* via `Page#load` pour la page actuellement active. Pour modifier la page active, appelez `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="eccf0-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="eccf0-128">Des m?tadonn?es, telles que le titre, peuvent toujours ?tre interrog?es pour n?importe quelle page.</span><span class="sxs-lookup"><span data-stu-id="eccf0-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="eccf0-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eccf0-129">See also</span></span>

- [<span data-ttu-id="eccf0-130">Vue d?ensemble de la programmation de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="eccf0-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="eccf0-131">R?f?rence de l?API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="eccf0-131">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="eccf0-132">Exemple de grille d??valuation</span><span class="sxs-lookup"><span data-stu-id="eccf0-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="eccf0-133">Vue d?ensemble de la plateforme des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="eccf0-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
