---
title: Interface API JavaScript pour Office
description: ''
ms.date: 05/13/2019
localization_priority: Priority
ms.openlocfilehash: 8d834aee4c21448210d9619fedd42d5ebb79e09d
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575323"
---
# <a name="javascript-api-for-office"></a>Interface API JavaScript pour Office

L’interface API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans les applications hôtes Office. Votre application fera référence à la bibliothèque office.js, qui est un chargeur de script. La bibliothèque office.js charge les modèles objet applicables à l’application Office qui exécute le complément. Vous pouvez utiliser les modèles objet JavaScript suivants :

- **API courantes** – API qui ont été introduites avec **Office 2013**. Il est chargé pour **toutes les applications hôtes Office** et connecte votre application de complément à l’application cliente Office. Le modèle objet contient les API propres aux clients Office et les API applicables à plusieurs applications hôtes clientes Office. Tout ce contenu se trouve sous **API partagé**. Ce modèle objet utilise des rappels. 

  **Outlook** utilise également la syntaxe des API courantes. Tous les éléments sous l’alias Office dans le code contiennent des objets que vous pouvez utiliser pour écrire un script qui interagit avec le contenu dans les documents, feuilles de calcul, présentations, éléments de courrier et projets de vos compléments Office. Vous devez utiliser ces API communes si le complément est destiné pour Office 2013 ou une version récente. Ce modèle objet utilise des rappels.

- **API propres à l’hôte** – API qui ont été introduites avec **Office 2016**. Ce modèle objet fournit des objets propres à l’hôte fortement typés qui correspondent aux objets habituels que vous voyez lorsque vous utilisez des clients Office. Il représente l’avenir des API JavaScript Office. Des API JavaScript spécifiques de l’hôte sont actuellement disponibles pour Excel, OneNote, PowerPoint et Word.

## <a name="supported-host-applications"></a>Applications hôtes prises en charge

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint](overview/powerpoint-add-ins-reference-overview.md)
- [Project](overview/project-add-ins-reference-overview.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API communes](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [Project](overview/project-add-ins-reference-overview.md) prend en charge les compléments créés avec l’API JavaScript, mais il n’existe actuellement aucune API JavaScript spécifiquement conçue pour interagir avec Project. Vous pouvez utiliser l’API commune pour créer des compléments Project.

En savoir plus sur les [hôtes pris en charge et les autres exigences](../concepts/requirements-for-running-office-add-ins.md).

## <a name="open-api-specifications"></a>Spécifications d’ouverture de l’API

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](openspec/openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.

## <a name="see-also"></a>Voir aussi

- [Référence de l’API JavaScript d’Office](/javascript/api/overview/office)
