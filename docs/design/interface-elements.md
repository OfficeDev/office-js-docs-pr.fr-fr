---
title: Éléments d’interface utilisateur Office pour les compléments Office
description: Obtenez une vue d’ensemble des différents types d’éléments d’interface utilisateur dans un complément Office.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: f553a6ac63fa7c99d8a770a6a1127591b819935e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608515"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Éléments d’interface utilisateur Office pour les compléments Office

Vous pouvez utiliser plusieurs types d’éléments d’interface utilisateur pour étendre l’interface utilisateur d’Office, y compris des commandes de complément et des conteneurs HTML. Ces éléments d’interface utilisateur ressemblent à une extension naturelle d’Office et fonctionnent sur les plateformes. Vous pouvez insérer votre code basé sur le web personnalisé dans l’un de ces éléments.

L’image suivante montre les types d’éléments d’interface utilisateur d’Office que vous pouvez créer.

![Image qui affiche des commandes de complément sur le ruban, un volet des tâches et une boîte de dialogue dans un document Office](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>Commandes de complément

Utilisez des [commandes de complément](add-in-commands.md) pour ajouter des points d’entrée vers votre complément au ruban Office. Les commandes démarrent les actions dans votre complément en exécutant du code JavaScript ou en lançant un conteneur HTML. Vous pouvez créer deux types de commandes de complément.

|**Type de commande**|**Description**|
|:---------------|:--------------|
|Onglets, menus et boutons du ruban|Permet d’ajouter des boutons personnalisés, des menus (déroulants) ou des onglets au ruban par défaut dans Office. Utilisez les boutons et menus pour déclencher une action dans Office. Utilisez les onglets pour regrouper et organiser des boutons et menus.|
|Menus contextuels| Permet de développer le menu contextuel par défaut. Les menus contextuels s’affichent lorsque les utilisateurs cliquent avec le bouton droit de la souris sur du texte dans un document Office ou un tableau dans Excel.| 

## <a name="html-containers"></a>Conteneurs HTML

Utilisez les conteneurs HTML pour intégrer du code de l’interface utilisateur basé sur HTML dans les clients Office. Ces pages web peuvent ensuite référencer l’API JavaScript Office pour interagir avec du contenu dans le document. Vous pouvez créer trois types de conteneurs HTML.

|**Conteneur HTML**|**Description**|
|:-----------------|:--------------|
|[Volets des tâches](task-pane-add-ins.md)|Permet d’afficher l’interface utilisateur personnalisée dans le volet droit du document Office. Utilisez les volets des tâches pour permettre aux utilisateurs d’interagir côte à côte avec votre complément et le document Office.|
|[Compléments de contenu](content-add-ins.md)|Permet d’afficher l’interface utilisateur personnalisée incorporée dans les documents Office. Utilisez les compléments de contenu pour permettre aux utilisateurs d’interagir avec votre complément directement dans le document Office. Par exemple, vous pouvez afficher du contenu externe tel que des vidéos ou des visualisations de données provenant d’autres sources. |
|[Boîtes de dialogue](dialog-boxes.md)|Permet d’afficher l’interface utilisateur personnalisée dans une boîte de dialogue superposée sur le document Office. Utilisez une boîte de dialogue pour les interactions qui nécessitent de l’attention et plus de valeur et ne nécessitent pas une interaction côte-à-côte avec le document.|

## <a name="see-also"></a>Voir aussi

- [Commandes de complément pour Excel, Word et PowerPoint](add-in-commands.md)
- [Volets des tâches](task-pane-add-ins.md)
- [Compléments de contenu](content-add-ins.md)
- [Boîtes de dialogue](dialog-boxes.md)
