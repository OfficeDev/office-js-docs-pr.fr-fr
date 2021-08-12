---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la callout qui s’affiche lorsque le add-in est installé dans Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86a96ba0b448e518f48deefc7aceb5133904c9147ba39a28bf646d922c449af7
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086662"
---
# <a name="getstarted-element"></a>Élément GetStarted

Fournit des informations utilisées par la callout qui s’affiche lorsque le add-in est installé dans Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                       | Obligatoire | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Titre](#title)               | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément     |
| [Description](#description)   | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Oui       | URL vers une page qui décrit le complément de façon plus détaillée.   |

### <a name="title"></a>Titre 

Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. **L’attribut resid** fait référence à un ID valide dans l’élément **ShortStrings** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.

### <a name="description"></a>Description

Obligatoire. Description/Contenu du corps de la légende. **L’attribut resid** fait référence à un ID valide dans l’élément **LongStrings** de la section [Resources](resources.md) et ne peut pas être plus de 32 caractères.

### <a name="learnmoreurl"></a>LearnMoreUrl

Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. **L’attribut resid** fait référence à un ID valide dans l’élément **Urls** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.

> [!NOTE]
> **LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint. Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible. 

## <a name="see-also"></a>Voir aussi

Les exemples de code suivants utilisent **l’élément GetStarted.**

* [Complément web Excel pour manipuler la mise en forme de tableau et de graphique](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Complément Word JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
