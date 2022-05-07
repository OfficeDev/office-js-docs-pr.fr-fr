---
title: Concevoir l'interface utilisateur des modules complémentaires d'Office
description: Apprenez les meilleures pratiques pour la conception visuelle des compléments d'Office.
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: efbb0ee5f0ba75170b8bd4343392c07d9eda8501
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244750"
---
# <a name="design-the-ui-of-office-add-ins"></a>Concevoir l'interface utilisateur des modules complémentaires d'Office

Les compléments Office prolongent les fonctionnalités d’Office en offrant des fonctions contextuelles auxquelles les utilisateurs peuvent accéder au sein de clients Office. Les compléments permettent aux utilisateurs d’être plus productifs en leur donnant accès à des fonctionnalités externes au sein d’Office, sans avoir à gérer de coûteux changements de contexte.

Votre conception d’interface utilisateur de complément doit s’intégrer en toute transparence à Office afin de fournir une interaction efficace et naturelle à vos utilisateurs. Tirez parti des [commandes de complément](add-in-commands.md) pour fournir l’accès à votre complément et appliquer les meilleures pratiques que nous vous recommandons lorsque vous créez une interface utilisateur HTML personnalisée.

## <a name="office-design-principles"></a>Principes de conception Office

Les applications Office suivent un ensemble général de directives d'interaction. Les applications partagent du contenu et possèdent des éléments dont l'apparence et le comportement sont similaires. Ces points communs reposent sur un ensemble de principes de conception. Ces principes aident l'équipe Office à créer des interfaces qui prennent en charge les tâches des clients. Comprendre et respecter ces principes vous aidera à atteindreles objectifs de vos clients dans Office.

Suivez les principes de conception d’Office pour créer des expériences de complément positives.

- **Concevoir explicitement pour Office.** La fonctionnalité, ainsi que l’apparence, d’un complément doivent compléter l’expérience Office. Les compléments doivent être natifs. Ils doivent s’ajuster parfaitement à Word sur un iPad ou à PowerPoint sur le web. Un complément bien conçu sera un mélange approprié de votre expérience, de la plateforme et de l’application Office. Appliquez les thèmes de document et d’interface utilisateur le cas échéant. Envisagez d’utiliser l’[Interface utilisateur Fluent pour le web](https://developer.microsoft.com/fluentui#/get-started/web) comme langage de conception et ensemble d’outils. L’interface utilisateur Fluent pour le web a deux versions.

  - **Pour les IU non React :** Utilisez **Fabric Core** , une collection open-source de classes CSS et de modules mixtes SASS qui vous donnent accès aux couleurs, aux animations, aux polices, aux icônes et aux grilles. (Il est appelé «Fabric Core» au lieu de «Fluent Core» pour des raisons historiques). Pour commencer, voir [Fabric Core dans Office Compléments](fabric-core.md).
  - **Pour les interfaces utilisateur React :** utilisez **Fluent UI React**, un cadre frontal React conçu pour créer des expériences qui s'intègrent parfaitement dans une large gamme de produits Microsoft. Il fournit des composants robustes, à jour, accessibles et basés sur React, qui sont hautement personnalisables à l'aide de CSS-in-JS. Pour commencer, voir [Fluent UI React dans Office Compléments](using-office-ui-fabric-react.md).

- **Contenu Favor sur chrome.** Autoriser les clients&rsquo; page, diapositive ou feuille de calcul pour rester le focus de l’expérience. Un complément est une interface auxiliaire. Aucun chrome d’accessoire ne doit interférer avec le contenu et les fonctionnalités du complément. Marquez votre expérience avec personnalisation. Nous savons qu’il est important de fournir aux utilisateurs une expérience unique et reconnaissable, tout en évitant toute distraction. Efforcez-vous de garder le focus sur le contenu et l’achèvement des tâches, et non sur l’attention de la marque.

- **Rendez l’expérience agréable et garder les utilisateurs au contrôle.** Les utilisateurs aiment utiliser des produits fonctionnels et visuellement attrayants. Concevez votre expérience avec soin. Obtenez les détails correctement en tenant compte de chaque interaction et détail visuel. Autoriser les utilisateurs à contrôler leur expérience. Les étapes nécessaires pour effectuer une tâche doivent être claires et pertinentes. Les décisions importantes doivent être faciles à comprendre. Les actions doivent être facilement réversibles. Un complément n’est pas une – destination, mais ’ une amélioration des fonctionnalités d’Office.

- **Prenez en compte toutes les plateformes et les méthodes d’entrée lors de la conception**. Les compléments sont conçus pour fonctionner sur toutes les plateformes prenant en charge Office ; aussi l’expérience utilisateur de votre complément doit-elle être optimisée pour fonctionner avec toutes les plateformes et tous les facteurs de forme. Veillez à ce que votre complément prenne aussi bien en charge les périphériques de type souris/clavier que les appareils et assurez-vous que votre interface utilisateur HTML personnalisée puisse s’adapter à différents facteurs de forme. Pour plus d’informations, consultez notre section relative aux [fonctions tactiles](../concepts/add-in-development-best-practices.md#optimize-for-touch).

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
