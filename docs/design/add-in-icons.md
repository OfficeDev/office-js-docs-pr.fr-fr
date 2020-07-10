---
title: Instructions concernant les îcones pour les compléments Office
description: Obtenez une vue d’ensemble de la conception des icônes et des styles de conception frais et monolignes pour les commandes de complément.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 35d8e0337b412a9ddebcde5be4db4db802e88269
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093839"
---
# <a name="icons"></a>Icônes

Les icônes sont la représentation visuelle d’un comportement ou d’un concept. Elles sont souvent utilisées pour ajouter une signification aux contrôles et commandes. Les visuels, qu’ils soient réalistes ou symboliques, permettent à l’utilisateur de naviguer dans l’interface utilisateur de la même façon que les signes l’aident à naviguer dans son environnement. Ils doivent être simples, clairs et contenir uniquement les informations nécessaires pour permettre aux clients d’analyser rapidement l’action qui se produit lorsqu’ils choisissent un contrôle.

Les interfaces du ruban de l’application Office ont un style visuel standard. Cela garantit la cohérence dans les applications Office. Les instructions vous aident à créer un ensemble de composants PNG pour votre solution qui s’intègrent naturellement dans Office.

Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

## <a name="design-icons-for-add-in-commands"></a>Concevoir des icônes pour les commandes de complément

[Commandes de complément](add-in-commands.md) Ajoutez des boutons, du texte et des icônes à l’interface utilisateur Office. Vos boutons de commande de complément doivent fournir des icônes significatives et des étiquettes qui identifient clairement l’action que l’utilisateur effectue lorsqu’il utilise une commande. Les articles suivants fournissent des directives stylistiques et de production pour vous aider à concevoir des icônes qui s’intègrent en toute transparence avec Office.

- Pour le style monoligne de Microsoft 365, reportez-vous à la rubrique [règles d’icône de style monoligne pour les compléments Office](add-in-icons-monoline.md).
- Pour le nouveau style de non-abonnement Office 2013 +, voir [règles d’icône de style frais pour les compléments Office](add-in-icons-fresh.md).

> [!NOTE]
> Vous devez choisir un style ou l’autre, et votre complément utilisera les mêmes icônes, qu’il soit en cours d’exécution dans Microsoft 365 ou sans abonnement Office.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
- [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)
