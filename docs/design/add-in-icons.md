---
title: Instructions concernant les îcones pour les compléments Office
description: Obtenez une vue d’ensemble de la conception des icônes et des styles de conception Fresh et Monoline pour les commandes de add-in.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0eeed616651535dfd20b6ac443a76fb1138ff51d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150188"
---
# <a name="icons"></a>Icônes

Les icônes sont la représentation visuelle d’un comportement ou d’un concept. Elles sont souvent utilisées pour ajouter une signification aux contrôles et commandes. Les visuels, qu’ils soient réalistes ou symboliques, permettent à l’utilisateur de naviguer dans l’interface utilisateur de la même façon que les signes l’aident à naviguer dans son environnement. Ils doivent être simples, clairs et contenir uniquement les informations nécessaires pour permettre aux clients d’analyser rapidement l’action qui se produit lorsqu’ils choisissent un contrôle.

application Office interfaces du ruban ont un style visuel standard. Cela garantit la cohérence dans les applications Office. Les instructions vous aident à créer un ensemble de composants PNG pour votre solution qui s’intègrent naturellement dans Office.

De nombreux conteneurs HTML contiennent des contrôles avec iconographie. Utilisez la police personnalisée de Fabric Core pour restituer Office icônes avec style dans votre add-in. La police d’icône fournie par [Fabric Core](fabric-core.md) contient de nombreux glyphes pour les métaphores Office communes que vous pouvez mettre à l’échelle, couleur et style en fonction de vos besoins. Si vous avez un langage visuel existant avec votre propre jeu d’icônes, n’hésitez pas à l’utiliser dans vos canevas HTML. Créer la continuité avec votre marque avec un jeu d’icônes standard est une partie importante de tout langage de création. Soyez prudent pour éviter de créer de la confusion pour les clients en conflit avec les métaphores Office.

## <a name="design-icons-for-add-in-commands"></a>Concevoir des icônes pour les commandes de complément

[Commandes de complément](add-in-commands.md) Ajoutez des boutons, du texte et des icônes à l’interface utilisateur Office. Vos boutons de commande de complément doivent fournir des icônes significatives et des étiquettes qui identifient clairement l’action que l’utilisateur effectue lorsqu’il utilise une commande. Les articles suivants fournissent des instructions stylistiques et de production pour vous aider à concevoir des icônes qui s’intègrent parfaitement à Office.

- Pour le style Monoline de Microsoft 365, voir recommandations en matière d’icônes de [style Monoline pour Office des modules.](add-in-icons-monoline.md)
- Pour obtenir le style Fresh de la Office 2013+, consultez les instructions relatives aux icônes de style Fresh pour Office [des modules.](add-in-icons-fresh.md)

> [!NOTE]
> Vous devez choisir un style ou l’autre et votre add-in utilisera les mêmes icônes qu’il soit en cours d’exécution dans Microsoft 365 ou sans abonnement Office.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
- [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)
