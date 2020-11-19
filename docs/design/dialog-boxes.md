---
title: Boîtes de dialogue dans les compléments Office
description: Découvrez les meilleures pratiques pour la conception visuelle des boîtes de dialogue dans les compléments Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ab8ca2e768c63a53b05ed2d9ef459050455231fb
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132052"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Boîtes de dialogue dans les compléments Office

Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.

*Figure 1. Mise en page type pour une boîte de dialogue*

![Disposition classique d’une boîte de dialogue affichée dans une application Office](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:--------|
|<ul><li>Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
||<ul><li>Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</li></ul>|

## <a name="implementation"></a>Implémentation

Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.

## <a name="see-also"></a>Voir aussi

- [Dialog object](/javascript/api/office/office.dialog)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)
