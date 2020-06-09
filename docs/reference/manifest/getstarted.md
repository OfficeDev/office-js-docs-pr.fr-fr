---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1fbdd5d4f4365f9f8190805519fc7a70c8c87ca
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611833"
---
# <a name="getstarted-element"></a>GetStarted, élément

Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                       | Requis | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Titre](#title)               | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément     |
| [Description](#description)   | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Oui       | URL vers une page qui décrit le complément de façon plus détaillée.   |

### <a name="title"></a>Title 

Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **ShortStrings** dans la section [Resources](resources.md).

### <a name="description"></a>Description

Obligatoire. Description/Contenu du corps de la légende. L’attribut **resid** fait référence à un ID valide de l’élément **LongStrings** dans la section [Resources](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. L’attribut **resid** fait référence à un ID valide de l’élément **Urls** dans la section [Resources](resources.md).

> [!NOTE]
> **LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint. Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible. 

## <a name="see-also"></a>Voir aussi

Les exemples de code suivants utilisent l’élément **GetStarted** :

* [Complément web Excel pour manipuler la mise en forme de tableau et de graphique](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Complément Word JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
