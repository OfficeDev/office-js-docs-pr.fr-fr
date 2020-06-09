---
title: Concevoir des compléments Office
description: Découvrez les meilleures pratiques en matière de conception visuelle des Compléments Office.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: a2965c2ee148c82708b9c61edd853f112adcf93c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607679"
---
# <a name="design-office-add-ins"></a>Concevoir des compléments Office

Les compléments Office prolongent les fonctionnalités d’Office en offrant des fonctions contextuelles auxquelles les utilisateurs peuvent accéder au sein de clients Office. Les compléments permettent aux utilisateurs d’être plus productifs en leur donnant accès à des fonctionnalités tierces au sein d’Office, sans avoir à gérer de coûteux changements de contexte. 

Votre complément doit s’intégrer de façon harmonieuse avec Office pour fournir une interaction efficace et naturelle à vos utilisateurs. Vous pouvez tirer parti de [commandes de complément](add-in-commands.md) pour permettre aux utilisateurs d’accéder à votre complément et appliquer les meilleures pratiques que nous vous recommandons lorsque vous créez des éléments d’interface utilisateur HTML personnalisés.

## <a name="office-design-principles"></a>Principes de conception Office

Les applications Office suivent un ensemble général de directives d’interaction. Les applications partagent du contenu et ont des éléments dont l’aspect et le comportement sont similaires. Cette compatibilité est basée sur un ensemble de principes de conception. Les principes aident l’équipe d’Office à créer des interfaces qui prennent en charge les tâches des clients. Découvrez-les et suivez-les pour prendre en charge les objectifs de vos clients au sein d’Office.

Suivez les principes de conception d’Office pour créer des expériences de compléments positives :

- **Privilégiez une conception explicitement orientée vers Office.** La fonctionnalité et l’apparence d’un complément doivent compléter harmonieusement l’expérience Office. Les compléments doivent sembler natifs. Ils doivent s’intégrer de façon transparente dans Word sur un iPad ou PowerPoint sur le web. Un complément bien conçu sera une combinaison appropriée de votre expérience, de la plateforme et de l’application Office. Envisagez d’utiliser Office UI Fabric comme langage de création. Appliquez des thèmes de document et d’interface utilisateur le cas échéant.

- **Concentrez-vous sur quelques tâches clés et exécutez-les correctement.** Aidez les clients à mener leurs tâches à bien sans empiéter sur le reste de leur travail. Apportez une réelle valeur ajoutée aux clients. Concentrez-vous sur des scénarios d’utilisation courants, choisissez avec soin ceux qui profitent le plus aux utilisateurs lors de l’interaction avec les documents Office.

- **Privilégiez le contenu par rapport aux éléments de détail.** La page, la diapositive ou le tableur des clients doit rester le cœur de l’expérience. Un complément est une interface auxiliaire. Aucun gadget accessoire ne doit interférer avec le contenu et les fonctionnalités du complément. Personnalisez votre expérience de manière judicieuse. Nous savons qu’il est important de fournir aux utilisateurs une expérience unique, reconnaissable mais sans distraction. Efforcez-vous de toujours privilégier le contenu et la capacité à effectuer des tâches plutôt que de chercher à attirer l’attention sur votre marque.

- **Rendez-la agréable et laissez suffisamment de contrôle aux utilisateurs.** Nous aimons tous utiliser des produits qui sont à la fois attrayants visuellement et fonctionnels. Créez votre expérience avec soin. Obtenez les détails directement en tenant compte de chaque interaction et détail visuel. Permettez aux utilisateurs de contrôler leur expérience. Les étapes nécessaires pour exécuter une tâche doivent être claires et pertinentes. Les décisions importantes doivent être faciles à comprendre. Les actions doivent être facilement réversibles. Un complément n’est pas une destination : c’est une amélioration des fonctionnalités Office.

- **Prenez en compte toutes les plateformes et les méthodes d’entrée lors de la conception**. Les compléments sont conçus pour fonctionner sur toutes les plateformes prenant en charge Office ; aussi l’expérience utilisateur de votre complément doit-elle être optimisée pour fonctionner avec toutes les plateformes et tous les facteurs de forme. Veillez à ce que votre complément prenne aussi bien en charge les périphériques de type souris/clavier que les appareils et assurez-vous que votre interface utilisateur HTML personnalisée puisse s’adapter à différents facteurs de forme. Pour plus d’informations, consultez notre section relative aux [fonctions tactiles](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Voir aussi
- [Office UI Fabric](https://developer.microsoft.com/fabric) 
- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)

