---
title: Utiliser du contenu de page OneNote
description: ''
ms.date: 12/04/2017
---

# <a name="work-with-onenote-page-content"></a>Utiliser du contenu de page OneNote 

Dans l’API JavaScript des compléments OneNote, le contenu de page est représenté par le modèle objet suivant.

  ![Diagramme du modèle objet de page OneNote](../images/one-note-om-page.png)

- Un objet Page contient une collection d’objets PageContent.
- Un objet PageContent contient un type de contenu de Outline, Image ou Other.
- Un objet Outline contient une collection d’objets Paragraph.
- Un objet Paragraph contient un type de contenu RichText, Image, Table ou Other.

Pour créer une page OneNote vide, utilisez l’une des méthodes suivantes :

- [Section.addPage](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [Page.insertPageAsSibling](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

Utilisez ensuite les méthodes dans les objets suivants pour travailler avec le contenu de la page, comme Page.addOutline et Outline.appendHtml. 

- [Page](https://dev.office.com/reference/add-ins/onenote/page)
- [Outline](https://dev.office.com/reference/add-ins/onenote/outline)
- [Paragraph](https://dev.office.com/reference/add-ins/onenote/paragraph)

Le contenu et la structure d’une page OneNote sont représentés par du code HTML. Seul un sous-ensemble de code HTML est pris en charge pour créer ou mettre à jour du contenu de page, comme décrit ci-dessous.

## <a name="supported-html"></a>HTML pris en charge

L’API JavaScript des compléments OneNote prend en charge le code HTML suivant pour créer et mettre à jour du contenu de page :

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>Accès au contenu de la page

Vous pouvez uniquement accéder au *contenu de la page* via `Page#load` pour la page actuellement active. Pour modifier la page active, appelez `navigateToPage($page)`.

Des métadonnées, telles que le titre, peuvent toujours être interrogées pour n’importe quelle page.

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](onenote-add-ins-programming-overview.md)
- [Référence de l’API JavaScript de OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
