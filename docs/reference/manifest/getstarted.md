---
title: Élément GetStarted dans le fichier manifeste
description: Fournit des informations utilisées par la callout qui s’affiche lorsque le module complémentaire est installé dans Word, Excel, PowerPoint et OneNote.
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 493526c3ad4a8486b76a18ccf23c64720a359784
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340994"
---
# <a name="getstarted-element"></a>Élément GetStarted

Fournit des informations utilisées par la callout qui s’affiche lorsque le module complémentaire est installé dans Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Si **l’élément GetStarted** est omis, la légende utilise à la place les valeurs des éléments [DisplayName](displayname.md) et [Description](description.md) .

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## <a name="child-elements"></a>Éléments enfants

| Élément                       | Requis | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Titre](#title)               | Oui      | Le titre est utilisé pour la partie supérieure de la légende.     |
| [Description](#description)   | Oui      | Description/Contenu du corps de la légende.|
| [LearnMoreUrl](#learnmoreurl) | Oui       | URL vers une page qui décrit le complément de façon plus détaillée.   |

### <a name="title"></a>Titre 

Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. **L’attribut resid** fait référence à un ID valide dans **l’élément ShortStrings** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.

### <a name="description"></a>Description

Obligatoire. Description/Contenu du corps de la légende. **L’attribut resid** fait référence à un ID valide dans l’élément **LongStrings** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.

### <a name="learnmoreurl"></a>LearnMoreUrl

Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. **L’attribut resid** fait référence à un ID valide dans l’élément **Urls** de la section [Resources](resources.md) et ne peut pas avoir plus de 32 caractères.

> [!NOTE]
> **LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint. Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible. 

## <a name="see-also"></a>Voir aussi

Les exemples de code suivants utilisent **l’élément GetStarted** .

* [Complément web Excel pour manipuler la mise en forme de tableau et de graphique](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Complément Word JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
