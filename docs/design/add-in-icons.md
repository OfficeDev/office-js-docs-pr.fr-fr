---
title: Instructions concernant les îcones pour les compléments Office
description: Obtenez une vue d’ensemble de la conception des icônes et des styles de conception frais et monolignes pour les commandes de complément.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: ce474ef20493e738fca7072d5b6a3bcd28594fbb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718503"
---
# <a name="icons"></a>Icônes

Les icônes sont la représentation visuelle d’un comportement ou d’un concept. Elles sont souvent utilisées pour ajouter une signification aux contrôles et commandes. Les visuels, qu’ils soient réalistes ou symboliques, permettent à l’utilisateur de naviguer dans l’interface utilisateur de la même façon que les signes l’aident à naviguer dans son environnement. Ils doivent être simples, clairs et contenir uniquement les informations nécessaires pour permettre aux clients d’analyser rapidement l’action qui se produit lorsqu’ils choisissent un contrôle.

Les interfaces de ruban Office ont un style visuel standard. Cela garantit la cohérence dans les applications Office. Les instructions vous aident à créer un ensemble de composants PNG pour votre solution qui s’intègrent naturellement dans Office.

De nombreux conteneurs HTML contiennent des contrôles avec iconographie. Utilisez la police personnalisée d’Office UI Fabric pour le rendu des icônes de style Office dans votre complément. La police d’icône de Fabric contient de nombreux glyphes pour les métaphores Office courantes que vous pouvez redimensionner, colorier et personnaliser selon vos besoins. Si vous avez un langage visuel existant avec votre propre jeu d’icônes, n’hésitez pas à l’utiliser dans vos canevas HTML. Créer la continuité avec votre marque avec un jeu d’icônes standard est une partie importante de tout langage de création. Soyez prudent pour éviter de créer de la confusion pour les clients en conflit avec les métaphores Office.

## <a name="design-icons-for-add-in-commands"></a>Concevoir des icônes pour les commandes de complément

[Commandes de complément](add-in-commands.md) Ajoutez des boutons, du texte et des icônes à l’interface utilisateur Office. Vos boutons de commande de complément doivent fournir des icônes significatives et des étiquettes qui identifient clairement l’action que l’utilisateur effectue lorsqu’il utilise une commande. Les articles suivants fournissent des directives stylistiques et de production pour vous aider à concevoir des icônes qui s’intègrent en toute transparence avec Office.

- Pour le style monoligne d’Office 365, reportez-vous à la rubrique [règles d’icône de style monoligne pour les compléments Office](add-in-icons-monoline.md).
- Pour le nouveau style de non-abonnement Office 2013 +, voir [règles d’icône de style frais pour les compléments Office](add-in-icons-fresh.md).

> [!NOTE]
> Vous devez choisir un style ou l’autre, et votre complément utilisera les mêmes icônes, qu’il s’exécute dans Office 365 ou sans abonnement Office.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
- [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)
