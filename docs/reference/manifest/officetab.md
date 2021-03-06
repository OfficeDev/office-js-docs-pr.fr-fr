---
title: Élément OfficeTab dans le fichier manifest
description: L’élément OfficeTab définit l’onglet du ruban où votre commande de complément s’affiche.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 25e8044d8b3264bf9ee64c54487566bf11f0065e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292299"
---
# <a name="officetab-element"></a>Élément OfficeTab

Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut ( **domicile**, **message**ou **réunion**) ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  Groupe      | Oui |  Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.  |

Les valeurs d’onglet valides sont les suivantes `id` par application. Les valeurs en **gras** sont prises en charge dans l’ordinateur de bureau et en ligne (par exemple, Word 2016 ou version ultérieure sur Windows et Word sur le Web).

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut avoir jusqu’à six contrôles. L’attribut **ID** est obligatoire et chaque **ID** doit être unique dans le manifeste. L' **ID** est une chaîne avec un maximum de 125 caractères. Voir [Élément group](group.md).

## <a name="officetab-example"></a>Exemple OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
