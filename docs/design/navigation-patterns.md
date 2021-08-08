---
title: Modèles de navigation pour les compléments Office
description: Découvrez les meilleures pratiques d’utilisation des barres de commandes, des barres de tabulation et des boutons arrière pour concevoir la navigation d’un Office de commande.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 42f217d240f1c66c69ad94d59f501a73e6d9a85d1dc24e94b5bb5307a808a6eb
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081742"
---
# <a name="navigation-patterns"></a>Modèles de navigation

Les principales fonctionnalités d’un complément sont accessibles via les types de commande spécifique et la zone de l’écran limitée. Il est important que la navigation soit intuitive, fournisse du contexte et permette à l’utilisateur de se déplacer facilement au sein du complément.

## <a name="best-practices"></a>Meilleures pratiques

| À faire    | À ne pas faire |
| :---- | :---- |
| Vérifiez que l’utilisateur dispose d’une option de navigation clairement visible. | Ne compliquez pas le processus de navigation en utilisant des éléments d’interface utilisateur non standard.
| Servez-vous des composants suivants le cas échéant pour permettre aux utilisateurs de parcourir le complément. | N’ajoutez pas de difficultés qui empêcherait l’utilisateur de savoir où il se trouve ou de comprendre le contexte au sein du complément

## <a name="command-bar"></a>Barre de commandes

CommandBar est une surface du volet Des tâches qui héberge les commandes qui opèrent sur le contenu de la fenêtre, du panneau ou de la région parente qu’il se trouve au-dessus. Exemples de fonctionnalités facultatives : point d’accès au menu « hamburger », recherche et commandes sur le côté.

![Illustration montrant une barre de commandes dans un volet Office’application de bureau. Cet exemple montre une barre de commandes juste en dessous du nom du module qui inclut un menu hamburger et une recherche.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Barre d’onglets

La barre d’onglets affiche la navigation à l’aide de boutons avec du texte et des icônes empilés verticalement. Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.

![Illustration montrant une barre d’onglets dans Office volet Des tâches de l’application de bureau. Cet exemple montre une barre d’onglets immédiatement en dessous du nom du module avec les onglets « Accueil », « Paramètres », « Favoris » et « Compte ».](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Bouton Précédent

Le bouton Retour permet aux utilisateurs de récupérer à partir d’une action de navigation d’accès vers le bas. Ce modèle permet de vous assurer que les utilisateurs suivent une série d’étapes ordonnées.

![Illustration montrant un bouton Retour dans un volet Office’application de bureau. Cet exemple montre un bouton Retour juste en dessous du nom du module, en haut à gauche.](../images/add-in-back-button.png)
