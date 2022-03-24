---
title: Instructions de disposition pour les compléments Office
description: Obtenez des instructions sur la disposition d’un volet de tâches ou d’une boîte de dialogue dans un Office d’organisation.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3e9f64a4568015b4149eee537da2d80b0218ac2e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742975"
---
# <a name="layout"></a>Disposition

Chaque conteneur HTML incorporé dans Office aura une disposition. Ces dispositions sont les écrans principaux de votre complément. Dans ces dispositions, vous créerez des expériences qui permettent aux clients de lancer des actions, de modifier des paramètres, d’afficher, de faire défiler ou de parcourir du contenu. Concevez votre complément avec une disposition cohérente à travers les écrans afin de garantir la continuité de l’expérience. Si vous avez un site web existant que vos clients utilisent souvent, envisagez de réutiliser les dispositions de vos pages web existantes. Adaptez-les pour qu’elles s’intègrent harmonieusement dans des conteneurs HTML Office.

Pour des recommandations sur la disposition, voir [Volet des tâches](task-pane-add-ins.md), [Contenu](content-add-ins.md) et [Boîte de dialogue](dialog-boxes.md). Pour plus d’informations sur l’assemblage de [Fluent React](using-office-ui-fabric-react.md) d’interface utilisateur ou [de Office UI Fabric JS](fabric-core.md), composants dans des dispositions courantes et des flux d’expérience utilisateur, voir [modèles de modèles](ux-design-pattern-templates.md) de conception d’expérience utilisateur.

Appliquez les instructions générales suivantes pour les dispositions.

- Évitez les marges étroites ou larges sur vos conteneurs HTML. 20 pixels est une grande valeur par défaut.
- Alignez les éléments intentionnellement. Les retraits supplémentaires et les nouveaux points d’alignement doivent aider la hiérarchie visuelle.
- Les interfaces Office se trouvent sur une grille 4px. Essayez de conserver votre marge intérieure entre les éléments à des multiples de 4.
- Une interface surchargée peut être source de confusion et ne pas être utilisée facilement avec les interactions tactiles.
- Vérifiez que les dispositions sont cohérentes entre les écrans. Les modifications de disposition inattendues ressemblent à des bogues visuels qui contribuent à un manque de confiance en votre solution.
- Suivez les modèles de disposition courants. Les conventions permettent aux utilisateurs de comprendre comment utiliser une interface.
- Évitez les éléments redondants comme la personnalisation ou les commandes.
- Consolidez les contrôles et les affichages pour éviter une utilisation excessive de la souris.
- Créez des expériences réactives qui s’adaptent aux hauteurs et largeurs du conteneur HTML.
