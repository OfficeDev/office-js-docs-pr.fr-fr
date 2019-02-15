---
title: Utiliser du contenu de page OneNote
description: ''
ms.date: 01/10/2019
localization_priority: Normal
ms.openlocfilehash: b92f83ee85943ca4e819b04e2b6fa90e6d4fe4b3
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052720"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="e5a62-102">Utiliser du contenu de page OneNote</span><span class="sxs-lookup"><span data-stu-id="e5a62-102">Work with OneNote page content</span></span>

<span data-ttu-id="e5a62-103">Dans l’API JavaScript des compléments OneNote, le contenu de page est représenté par le modèle objet suivant.</span><span class="sxs-lookup"><span data-stu-id="e5a62-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagramme du modèle objet de page OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="e5a62-105">Un objet Page contient une collection d’objets PageContent.</span><span class="sxs-lookup"><span data-stu-id="e5a62-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="e5a62-106">Un objet PageContent contient un type de contenu de Outline, Image ou Other.</span><span class="sxs-lookup"><span data-stu-id="e5a62-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="e5a62-107">Un objet Outline contient une collection d’objets Paragraph.</span><span class="sxs-lookup"><span data-stu-id="e5a62-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="e5a62-108">Un objet Paragraph contient un type de contenu RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="e5a62-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="e5a62-109">Pour créer une page OneNote vide, utilisez l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="e5a62-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="e5a62-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="e5a62-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="e5a62-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="e5a62-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="e5a62-112">Utilisez ensuite les méthodes dans les objets suivants pour travailler avec le contenu de la page, comme `Page.addOutline` et `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="e5a62-112">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="e5a62-113">Page</span><span class="sxs-lookup"><span data-stu-id="e5a62-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="e5a62-114">Outline</span><span class="sxs-lookup"><span data-stu-id="e5a62-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="e5a62-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="e5a62-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="e5a62-p101">Le contenu et la structure d’une page OneNote sont représentés par du code HTML. Seul un sous-ensemble de code HTML est pris en charge pour créer ou mettre à jour du contenu de page, comme décrit ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="e5a62-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="e5a62-118">HTML pris en charge</span><span class="sxs-lookup"><span data-stu-id="e5a62-118">Supported HTML</span></span>

<span data-ttu-id="e5a62-119">L’API JavaScript des compléments OneNote prend en charge le code HTML suivant pour créer et mettre à jour du contenu de page :</span><span class="sxs-lookup"><span data-stu-id="e5a62-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="e5a62-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="e5a62-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="e5a62-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="e5a62-121">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="e5a62-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="e5a62-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="e5a62-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="e5a62-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="e5a62-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="e5a62-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="e5a62-125">L’importation du code HTML dans OneNote consolide les espaces blancs.</span><span class="sxs-lookup"><span data-stu-id="e5a62-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="e5a62-126">Le contenu obtenu est collé dans un plan.</span><span class="sxs-lookup"><span data-stu-id="e5a62-126">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="e5a62-127">OneNote fait de son mieux pour traduire le code HTML en contenu de page tout en assurant la sécurité des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="e5a62-127">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="e5a62-128">Les normes HTML et CSS ne correspondent pas exactement au modèle de contenu de OneNote, il y aura donc des différences d'apparence, en particulier avec les styles CSS.</span><span class="sxs-lookup"><span data-stu-id="e5a62-128">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="e5a62-129">Nous vous recommandons d’utiliser les objets JavaScript si une mise en forme spécifique est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e5a62-129">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="e5a62-130">Accès au contenu de la page</span><span class="sxs-lookup"><span data-stu-id="e5a62-130">Accessing page contents</span></span>

<span data-ttu-id="e5a62-p104">Vous pouvez uniquement accéder au *contenu de la page* via `Page#load` pour la page actuellement active. Pour modifier la page active, appelez `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="e5a62-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="e5a62-133">Des métadonnées, telles que le titre, peuvent toujours être interrogées pour n’importe quelle page.</span><span class="sxs-lookup"><span data-stu-id="e5a62-133">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="e5a62-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e5a62-134">See also</span></span>

- [<span data-ttu-id="e5a62-135">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="e5a62-135">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="e5a62-136">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="e5a62-136">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="e5a62-137">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="e5a62-137">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="e5a62-138">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e5a62-138">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
