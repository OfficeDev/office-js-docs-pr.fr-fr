---
title: Concevoir l'interface utilisateur des modules complémentaires d'Office
description: Apprenez les meilleures pratiques pour la conception visuelle des compléments d'Office.
ms.date: 05/12/2021
localization_priority: Priority
ms.openlocfilehash: 7b5314a07e15c5d57b4e5c27e781ebba5c1a3492
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330051"
---
# <a name="design-the-ui-of-office-add-ins"></a>Concevoir l'interface utilisateur des modules complémentaires d'Office

Les compléments Office prolongent les fonctionnalités d’Office en offrant des fonctions contextuelles auxquelles les utilisateurs peuvent accéder au sein de clients Office. Les compléments permettent aux utilisateurs d’être plus productifs en leur donnant accès à des fonctionnalités tierces au sein d’Office, sans avoir à gérer de coûteux changements de contexte.

La conception de l'interface utilisateur de votre module complémentaire doit s'intégrer parfaitement à Office pour offrir une interaction efficace et naturelle à vos utilisateurs. Profitez des commandes [de complément](add-in-commands.md) pour donner accès à votre complément et appliquez les meilleures pratiques que nous recommandons lorsque vous créez une interface utilisateur personnalisée basée sur HTML.

## <a name="office-design-principles"></a>Principes de conception Office

Les applications Office suivent un ensemble général de directives d'interaction. Les applications partagent du contenu et possèdent des éléments dont l'apparence et le comportement sont similaires. Ces points communs reposent sur un ensemble de principes de conception. Ces principes aident l'équipe Office à créer des interfaces qui prennent en charge les tâches des clients. Comprendre et respecter ces principes vous aidera à atteindreles objectifs de vos clients dans Office.

Suivez les principes de conception d’Office pour créer des expériences de compléments positives :

- **Privilégiez une conception explicitement orientée vers Office.** Les fonctionnalités, ainsi que l'aspect et la convivialité, d'un module complémentaire doivent compléter harmonieusement l'expérience d'Office. Les modules complémentaires doivent avoir l'air d'être natifs. Ils doivent s'intégrer parfaitement à Word sur un iPad ou à PowerPoint sur le web. Un module complémentaire bien conçu sera un mélange approprié de votre expérience, de la plateforme et de l'application Office. Appliquez les thèmes des documents et de l'interface utilisateur, le cas échéant. Considérez l'utilisation de [Fluent UI pour le web](https://developer.microsoft.com/fluentui#/get-started/web) comme votre langage de conception et votre ensemble d'outils. L'interface utilisateur Fluent pour le web a deux saveurs :

  - **Pour les IU non React :** Utilisez **Fabric Core** , une collection open-source de classes CSS et de modules mixtes SASS qui vous donnent accès aux couleurs, aux animations, aux polices, aux icônes et aux grilles. (Il est appelé «Fabric Core» au lieu de «Fluent Core» pour des raisons historiques). Pour commencer, voir [Fabric Core dans Office Compléments](fabric-core.md).
  - **Pour les interfaces utilisateur React :** utilisez **Fluent UI React**, un cadre frontal React conçu pour créer des expériences qui s'intègrent parfaitement dans une large gamme de produits Microsoft. Il fournit des composants robustes, à jour, accessibles et basés sur React, qui sont hautement personnalisables à l'aide de CSS-in-JS. Pour commencer, voir [Fluent UI React dans Office Compléments](using-office-ui-fabric-react.md).

- **Privilégiez le contenu au chrome.** Permettez à la page, à la diapositive ou à la feuille de calcul du client de rester au centre de l'expérience. Un complément est une interface auxiliaire. Aucun gadget accessoire ne doit interférer avec le contenu et les fonctionnalités du complément. Personnalisez votre expérience de manière judicieuse. Nous savons qu'il est important d'offrir aux utilisateurs une expérience unique et reconnaissable, tout en évitant les distractions. Efforcez-vous de toujours privilégier le contenu et la capacité à effectuer des tâches plutôt que de chercher à attirer l’attention sur votre marque.

- **Rendez-la agréable et laissez suffisamment de contrôle aux utilisateurs.** Nous aimons tous utiliser des produits qui sont à la fois attrayants visuellement et fonctionnels. Créez votre expérience avec soin. Obtenez les détails directement en tenant compte de chaque interaction et détail visuel. Permettez aux utilisateurs de contrôler leur expérience. Les étapes nécessaires pour exécuter une tâche doivent être claires et pertinentes. Les décisions importantes doivent être faciles à comprendre. Les actions doivent être facilement réversibles. Un complément n’est pas une destination : c’est une amélioration des fonctionnalités Office.

- **Prenez en compte toutes les plateformes et les méthodes d’entrée lors de la conception**. Les compléments sont conçus pour fonctionner sur toutes les plateformes prenant en charge Office ; aussi l’expérience utilisateur de votre complément doit-elle être optimisée pour fonctionner avec toutes les plateformes et tous les facteurs de forme. Veillez à ce que votre complément prenne aussi bien en charge les périphériques de type souris/clavier que les appareils et assurez-vous que votre interface utilisateur HTML personnalisée puisse s’adapter à différents facteurs de forme. Pour plus d’informations, consultez notre section relative aux [fonctions tactiles](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
