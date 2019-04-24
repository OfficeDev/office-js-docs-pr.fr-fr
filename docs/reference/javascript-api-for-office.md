---
title: Interface API JavaScript pour Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: c8b33bbf9d0107786c0272410c59b1a3fe998cba
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450715"
---
# <a name="javascript-api-for-office"></a>Interface API JavaScript pour Office

L’interface API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans les applications hôtes Office. Votre application fera référence à la bibliothèque office.js, qui est un chargeur de script. La bibliothèque office.js charge les modèles objet applicables à l’application Office qui exécute le complément. Vous pouvez utiliser les modèles objet JavaScript suivants :

- **API courantes** – API qui ont été introduites avec **Office 2013**. Il est chargé pour **toutes les applications hôtes Office** et connecte votre application de complément à l’application cliente Office. Le modèle objet contient les API propres aux clients Office et les API applicables à plusieurs applications hôtes clientes Office. Tout ce contenu se trouve sous **API partagé**. Ce modèle objet utilise des rappels. 

  **Outlook** utilise également la syntaxe des API courantes. Tous les éléments sous l’alias Office dans le code contiennent des objets que vous pouvez utiliser pour écrire un script qui interagit avec le contenu dans les documents, feuilles de calcul, présentations, éléments de courrier et projets de vos compléments Office. Vous devez utiliser ces API communes si le complément est destiné pour Office 2013 ou une version récente. Ce modèle objet utilise des rappels.

- **API propres à l’hôte** – API qui ont été introduites avec **Office 2016**. Ce modèle objet fournit des objets propres à l’hôte fortement typés qui correspondent aux objets habituels que vous voyez lorsque vous utilisez des clients Office. Il représente l’avenir des API JavaScript Office. Les API propres à l’hôte incluent actuellement l’API JavaScript Word et l’API JavaScript Excel.

## <a name="supported-host-applications"></a>Applications hôtes prises en charge

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API communes](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint et Project](requirement-sets/powerpoint-and-project-note.md) prennent en charge des compléments conçus à l’aide de l’API JavaScript. Toutefois, ils ne disposent pas actuellement des APIs propres à l’hôte spécifiques. Vous pouvez interagir avec ces hôtes via l’API commune.

En savoir plus sur les [hôtes pris en charge et les autres exigences](../concepts/requirements-for-running-office-add-ins.md).

## <a name="open-api-specifications"></a>Spécifications d’ouverture de l’API

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.

## <a name="see-also"></a>Voir aussi

- [Référence de l’API JavaScript d’Office](/javascript/api/overview/office)
