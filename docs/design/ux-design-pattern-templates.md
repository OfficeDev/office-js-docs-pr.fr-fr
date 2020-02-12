---
title: Modèles de conception de l’expérience utilisateur pour les compléments Office
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d903f6cb2c6cad90c07b05303eac6b25a05a4af2
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950417"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Modèles de conception de l’expérience utilisateur pour les compléments Office

La conception de l’expérience utilisateur pour les compléments Office doit offrir une expérience attrayante aux utilisateurs d’Office et étendre l’expérience générale Office en s'intégrant parfaitement dans l’interface utilisateur Office par défaut.  

Nos modèles d’expérience utilisateur sont composés de composants. Les composants sont des contrôles qui aident vos clients à interagir avec les éléments de votre logiciel ou service. Les boutons, la navigation et les menus sont des exemples de composants courants qui ont souvent des comportements et des styles cohérents.

Office UI Fabric rend les composants qui ressemblent à une partie d’Office et se comportent comme une partie d’Office. Utilisez Fabric pour une intégration facile avec Office. Si votre complément a son propre langage de composant préexistant, vous n’avez pas besoin de l’abandonner en faveur de Fabric. Recherchez les opportunités pour le conserver lors de l’intégration avec Office. Pensez à remplacer les éléments stylistiques, à supprimer les conflits ou à adopter des styles et des comportements qui éliminent la confusion de l’utilisateur.

Les modèles fournis sont les meilleures solutions pratiques basées sur des scénarios courants d’utilisation et sur les recherches en expérience utilisateur. Ils sont destinés à fournir une entrée rapide à la conception et au développement de compléments, et fournir des conseils pour obtenir un équilibre entre les éléments Microsoft et les éléments de la marque. Fournir une expérience utilisateur propre et moderne qui assure un équilibre entre les éléments de conception du langage de conception Microsoft Fabric et l’identité de marque unique du partenaire peut vous aider à augmenter la rétention utilisateur et l’adoption de votre complément.

Utiliser les modèles de motif expérience utilisateur pour :

* Appliquer des solutions à des scénarios client courants.
* Appliquer les meilleures pratiques en matière de conception.
* Incorporer les composants et styles d’[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started).
* Créer des compléments qui s’intègrent visuellement à l’interface utilisateur d’Office par défaut.
* Imaginer et visualiser l’expérience utilisateur.

## <a name="getting-started"></a>Prise en main

Les modèles sont organisés par les actions clés ou les expériences qui sont courantes dans un complément. Les groupes principaux sont :

* [Première exécution](../design/first-run-experience-patterns.md)
* [Authentification](../design/authentication-patterns.md)
* [Navigation](../design/navigation-patterns.md)
* [Conception de personnalisation](../design/branding-patterns.md)

Étudiez chaque groupe pour apprendre comment concevoir votre complément en utilisant les meilleures pratiques.

> [!NOTE]
> Les écrans exemple illustrés dans l’ensemble de cette documentation sont conçus et affichés à une résolution de **1366 x 768**.

## <a name="see-also"></a>Voir aussi

* [Kits de ressources de conception](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [Meilleures pratiques en matière de développement de compléments Office](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Prise en main de Fabric React](/office/dev/add-ins/design/using-office-ui-fabric-react)
