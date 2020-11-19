---
title: Modèles de navigation pour les compléments Office
description: Découvrez les meilleures pratiques pour l’utilisation des barres de commandes, des barres d’onglets et des boutons de retour, pour concevoir la navigation d’un complément Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132031"
---
# <a name="navigation-patterns"></a>Modèles de navigation

Les principales fonctionnalités d’un complément sont accessibles via les types de commande spécifique et la zone de l’écran limitée. Il est important que la navigation soit intuitive, fournisse du contexte et permette à l’utilisateur de se déplacer facilement au sein du complément.

## <a name="best-practices"></a>Meilleures pratiques

| À faire    | À ne pas faire |
| :---- | :---- |
| Vérifiez que l’utilisateur dispose d’une option de navigation clairement visible. | Ne compliquez pas le processus de navigation en utilisant des éléments d’interface utilisateur non standard.
| Servez-vous des composants suivants le cas échéant pour permettre aux utilisateurs de parcourir le complément. | N’ajoutez pas de difficultés qui empêcherait l’utilisateur de savoir où il se trouve ou de comprendre le contexte au sein du complément

## <a name="command-bar"></a>Barre de commandes

La barre de commandes est une surface dans le volet Office qui héberge des commandes qui fonctionnent sur le contenu de la fenêtre, du panneau ou de la région parent qu’elle contient. Exemples de fonctionnalités facultatives : point d’accès au menu « hamburger », recherche et commandes sur le côté.

![Illustration d’une barre de commandes dans un volet Office d’une application de bureau Office. Cet exemple montre comment afficher une barre de commandes immédiatement sous le nom du complément, qui comprend un menu hamburger et une recherche.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Barre d’onglets

La barre d’onglets affiche la navigation à l’aide de boutons avec du texte et des icônes verticalement empilés. Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.

![Illustration d’une barre d’onglets dans le volet Office d’une application de bureau Office. Cet exemple montre une barre d’onglets immédiatement sous le nom du complément avec les onglets « Accueil », « paramètres », « favoris » et « compte ».](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Bouton Précédent

Le bouton précédent permet aux utilisateurs de récupérer à partir d’une action de navigation d’exploration. Ce modèle permet de vous assurer que les utilisateurs suivent une série d’étapes ordonnées.

![Illustration illustrant un bouton retour dans le volet Office d’une application de bureau Office. Cet exemple montre un bouton précédent immédiatement sous le nom du complément, dans la partie supérieure gauche.](../images/add-in-back-button.png)
