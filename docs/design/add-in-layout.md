---
title: Instructions de disposition pour les compléments Office
description: Obtenir des instructions sur la mise en page d’un volet Office ou d’une boîte de dialogue dans un complément Office.
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: c88d2e9d978a22688eb10b57607750286209e5d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607665"
---
# <a name="layout"></a>Disposition
Chaque conteneur HTML incorporé dans Office aura une disposition. Ces dispositions sont les écrans principaux de votre complément. Dans ces dispositions, vous créerez des expériences qui permettent aux clients de lancer des actions, de modifier des paramètres, d’afficher, de faire défiler ou de parcourir du contenu. Concevez votre complément avec une disposition cohérente à travers les écrans afin de garantir la continuité de l’expérience. Si vous avez un site web existant que vos clients utilisent souvent, envisagez de réutiliser les dispositions de vos pages web existantes. Adaptez-les pour qu’elles s’intègrent harmonieusement dans des conteneurs HTML Office.

Pour des recommandations sur la disposition, voir [Volet des tâches](task-pane-add-ins.md), [Contenu](content-add-ins.md) et [Boîte de dialogue](dialog-boxes.md). Pour plus d’informations sur la façon d’assembler des composants Office UI Fabric dans des flux d’expérience utilisateur et des dispositions courants , voir [Modèles de conception UX](ux-design-pattern-templates.md).

Appliquez les recommandations générales suivantes pour les dispositions :

*   Évitez les marges étroites ou larges sur vos conteneurs HTML. 20 pixels est une grande valeur par défaut.
*   Alignez les éléments intentionnellement. Les retraits supplémentaires et les nouveaux points d’alignement doivent aider la hiérarchie visuelle.
*   Les interfaces Office se trouvent sur une grille 4px. Essayez de conserver votre marge intérieure entre les éléments à des multiples de 4.
*   Une interface surchargée peut être source de confusion et ne pas être utilisée facilement avec les interactions tactiles.
*   Vérifiez que les dispositions sont cohérentes entre les écrans. Les modifications de disposition inattendues ressemblent à des bogues visuels qui contribuent à un manque de confiance en votre solution.
*   Suivez les modèles de disposition courants. Les conventions permettent aux utilisateurs de comprendre comment utiliser une interface.
*   Évitez les éléments redondants comme la personnalisation ou les commandes.
*   Consolidez les contrôles et les affichages pour éviter une utilisation excessive de la souris.
*   Créez des expériences réactives qui s’adaptent aux hauteurs et largeurs du conteneur HTML.
