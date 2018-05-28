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
# <a name="work-with-onenote-page-content"></a>Utiliser du contenu de page OneNote 

Dans l?API JavaScript des compl?ments OneNote, le contenu de page est repr?sent? par le mod?le objet suivant.

  ![Diagramme du mod?le objet de page OneNote](../images/one-note-om-page.png)

- Un objet Page contient une collection d?objets PageContent.
- Un objet PageContent contient un type de contenu de Outline, Image ou Other.
- Un objet Outline contient une collection d?objets Paragraph.
- Un objet Paragraph contient un type de contenu RichText, Image, Table ou Other.

Pour cr?er une page OneNote vide, utilisez l?une des m?thodes suivantes :

- [Section.addPage](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [Page.insertPageAsSibling](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

Utilisez ensuite les m?thodes dans les objets suivants pour travailler avec le contenu de la page, comme Page.addOutline et Outline.appendHtml. 

- [Page](https://dev.office.com/reference/add-ins/onenote/page)
- [Outline](https://dev.office.com/reference/add-ins/onenote/outline)
- [Paragraph](https://dev.office.com/reference/add-ins/onenote/paragraph)

Le contenu et la structure d?une page OneNote sont repr?sent?s par du code HTML. Seul un sous-ensemble de code HTML est pris en charge pour cr?er ou mettre ? jour du contenu de page, comme d?crit ci-dessous.

## <a name="supported-html"></a>HTML pris en charge

L?API JavaScript des compl?ments OneNote prend en charge le code HTML suivant pour cr?er et mettre ? jour du contenu de page :

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>Acc?s au contenu de la page

Vous pouvez uniquement acc?der au *contenu de la page* via `Page#load` pour la page actuellement active. Pour modifier la page active, appelez `navigateToPage($page)`.

Des m?tadonn?es, telles que le titre, peuvent toujours ?tre interrog?es pour n?importe quelle page.

## <a name="see-also"></a>Voir aussi

- [Vue d?ensemble de la programmation de l?API JavaScript de OneNote](onenote-add-ins-programming-overview.md)
- [R?f?rence de l?API JavaScript de OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Exemple de grille d??valuation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d?ensemble de la plateforme des compl?ments Office](../overview/office-add-ins.md)
