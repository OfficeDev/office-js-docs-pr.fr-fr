---
title: Utiliser du contenu de page OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 3ceb693b85490e5b7046880a79ae46753a1d3238
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944126"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="b63dd-102">Utiliser du contenu de page OneNote</span><span class="sxs-lookup"><span data-stu-id="b63dd-102">Work with OneNote page content</span></span> 

<span data-ttu-id="b63dd-103">Dans l’API JavaScript des compléments OneNote, le contenu de page est représenté par le modèle objet suivant.</span><span class="sxs-lookup"><span data-stu-id="b63dd-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagramme du modèle objet de page OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="b63dd-105">Un objet Page contient une collection d’objets PageContent.</span><span class="sxs-lookup"><span data-stu-id="b63dd-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="b63dd-106">Un objet PageContent contient un type de contenu de Outline, Image ou Other.</span><span class="sxs-lookup"><span data-stu-id="b63dd-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="b63dd-107">Un objet Outline contient une collection d’objets Paragraph.</span><span class="sxs-lookup"><span data-stu-id="b63dd-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="b63dd-108">Un objet Paragraph contient un type de contenu RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="b63dd-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="b63dd-109">Pour créer une page OneNote vide, utilisez l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="b63dd-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="b63dd-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="b63dd-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [<span data-ttu-id="b63dd-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="b63dd-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

<span data-ttu-id="b63dd-112">Utilisez ensuite les méthodes dans les objets suivants pour travailler avec le contenu de la page, comme Page.addOutline et Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="b63dd-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="b63dd-113">Page</span><span class="sxs-lookup"><span data-stu-id="b63dd-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [<span data-ttu-id="b63dd-114">Structure</span><span class="sxs-lookup"><span data-stu-id="b63dd-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [<span data-ttu-id="b63dd-115">Paragraphe</span><span class="sxs-lookup"><span data-stu-id="b63dd-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

<span data-ttu-id="b63dd-p101">Le contenu et la structure d’une page OneNote sont représentés par du code HTML. Seul un sous-ensemble de code HTML est pris en charge pour créer ou mettre à jour du contenu de page, comme décrit ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="b63dd-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="b63dd-118">HTML pris en charge</span><span class="sxs-lookup"><span data-stu-id="b63dd-118">Supported HTML</span></span>

<span data-ttu-id="b63dd-119">L’API JavaScript des compléments OneNote prend en charge le code HTML suivant pour créer et mettre à jour du contenu de page :</span><span class="sxs-lookup"><span data-stu-id="b63dd-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="b63dd-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="b63dd-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="b63dd-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="b63dd-121">`<ul>`, `<ol>`, `<li>`</span></span> 
- <span data-ttu-id="b63dd-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="b63dd-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="b63dd-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="b63dd-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="b63dd-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="b63dd-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="b63dd-125">Accès au contenu de la page</span><span class="sxs-lookup"><span data-stu-id="b63dd-125">Accessing page contents</span></span>

<span data-ttu-id="b63dd-p102">Vous pouvez uniquement accéder au *contenu de la page* via `Page#load` pour la page actuellement active. Pour modifier la page active, appelez `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="b63dd-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="b63dd-128">Des métadonnées, telles que le titre, peuvent toujours être interrogées pour n’importe quelle page.</span><span class="sxs-lookup"><span data-stu-id="b63dd-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="b63dd-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b63dd-129">See also</span></span>

- [<span data-ttu-id="b63dd-130">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="b63dd-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="b63dd-131">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="b63dd-131">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="b63dd-132">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="b63dd-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="b63dd-133">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b63dd-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
