---
title: Modèles de conception de l’expérience utilisateur pour les compléments Office
description: Obtenez une vue d’ensemble des modèles de conception d’interface utilisateur pour les Office, y compris les modèles de navigation, d’authentification, de première utilisation et de authentification.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 2c7de48aef161a335fe67dcd4cc87cd32c6be10373f078baac77c9407ae1466b
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081472"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Modèles de conception de l’expérience utilisateur pour les compléments Office

La conception de l’expérience utilisateur pour les compléments Office doit offrir une expérience attrayante aux utilisateurs d’Office et étendre l’expérience générale Office en s'intégrant parfaitement dans l’interface utilisateur Office par défaut.  

Nos modèles d’expérience utilisateur sont composés de composants. Les composants sont des contrôles qui aident vos clients à interagir avec les éléments de votre logiciel ou service. Les boutons, la navigation et les menus sont des exemples de composants courants qui ont souvent des comportements et des styles cohérents.

[Fluent’interface utilisateur React composants](using-office-ui-fabric-react.md) se comportent comme une partie de Office, tout comme les composants neutres de l’infrastructure de [Office UI Fabric JS](fabric-core.md). Tirez parti de l’un ou l’autre des ensembles de composants à intégrer à Office. Sinon, si votre add-in possède son propre langage de composant existant, vous n’avez pas besoin de l’ignorer. Recherchez les opportunités pour le conserver lors de l’intégration avec Office. Pensez à remplacer les éléments stylistiques, à supprimer les conflits ou à adopter des styles et des comportements qui éliminent la confusion de l’utilisateur.

Les modèles fournis sont les meilleures solutions pratiques basées sur des scénarios courants d’utilisation et sur les recherches en expérience utilisateur. Ils sont destinés à fournir un point d’entrée rapide pour concevoir et développer des modules, ainsi que des conseils pour trouver un équilibre entre les éléments de marque Microsoft et les vôtres. Fournir une expérience utilisateur propre et moderne qui équilibre les éléments de conception du langage de conception de l’interface utilisateur Fluent de Microsoft et l’identité de marque unique du partenaire peut aider à augmenter la rétention et l’adoption par les utilisateurs de votre add-in.

Utiliser les modèles de motif expérience utilisateur pour :

* Appliquer des solutions à des scénarios client courants.
* Appliquer les meilleures pratiques en matière de conception.
* Incorporez [Fluent composants et](https://developer.microsoft.com/fluentui#/get-started) styles d’interface utilisateur.
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

* [Boîte de forme d’outils De conception](design-toolkits.md)
* [Interface utilisateur Fluent](https://developer.microsoft.com/fluentui#)
* [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
* [Interface utilisateur Fluent - Comment faire pour les modules add-in Office ?](using-office-ui-fabric-react.md)
