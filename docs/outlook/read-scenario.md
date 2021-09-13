---
title: Créer des compléments Outlook pour des formulaires de lecture
description: Les compléments de lecture sont des compléments Outlook qui sont activés dans le volet de lecture ou l’inspecteur de lecture dans Outlook.
ms.date: 03/19/2021
ms.localizationpriority: high
ms.openlocfilehash: 1bc65c64d0076a9a3b60aac3a18c950acff048d2
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150420"
---
# <a name="create-outlook-add-ins-for-read-forms"></a>Créer des compléments Outlook pour des formulaires de lecture

Les compléments de lecture sont des compléments Outlook activés dans le volet de lecture ou l’inspecteur de lecture d’Outlook. Contrairement aux compléments de composition (qui sont des compléments Outlook activés lorsqu’un utilisateur crée un message ou un rendez-vous), les compléments de lecture sont disponibles dans les scénarios suivants :

- Affichage d’un message électronique, d’une demande de réunion, d’une réponse à une demande de réunion ou d’une annulation de réunion.

   > [!NOTE]
   > Outlook n’active pas les compléments dans un formulaire de lecture pour certains types de messages, y compris les éléments qui sont les pièces jointes d’un autre message, les éléments du dossier Brouillons, ou encore ceux chiffrés ou protégés d’autres façons.

- Affichage d’un élément de réunion dans lequel l’utilisateur est un participant.

- Affichage d’un élément de réunion dans lequel l’utilisateur est l’organisateur (version RTM d’Outlook 2013 et d’Exchange 2013 uniquement).

   > [!NOTE]
   > À partir de la version Office 2013 SP1, si l’utilisateur visualise un élément de réunion dont il est l’organisateur, seuls les compléments de composition peuvent être activés et disponibles. Les compléments de lecture ne sont plus disponibles dans ce scénario.

Dans chacun de ces scénarios de lecture, Outlook active les compléments lorsque leurs conditions d’activation sont respectées. Les utilisateurs peuvent ensuite choisir et ouvrir les compléments activés dans la barre de compléments du volet de lecture ou de l’inspecteur de lecture. La figure suivante montre le complément **Bing Cartes** qui a été activé et ouvert alors que l’utilisateur lit un message contenant une adresse géographique.

**Volet de complément montrant le complément Bing Cartes en action pour le message Outlook sélectionné qui contient une adresse**

![Application de courrier Bing Map dans Outlook.](../images/outlook-detected-entity-card.png)

## <a name="types-of-add-ins-available-in-read-mode"></a>Types de complément disponibles en mode de lecture

Les compléments de lecture peuvent correspondre à n’importe quelle combinaison des types suivants.

- [Commandes de complément pour Outlook](add-in-commands-for-outlook.md)
- [Compléments Outlook contextuels](contextual-outlook-add-ins.md)

## <a name="api-features-available-to-read-add-ins"></a>Fonctionnalités de l’API disponibles pour les compléments de lecture

- Pour activer les compléments dans les formulaires de lecture, voir le tableau 1 dans [Spécifier des règles d’activation dans un manifeste](activation-rules.md#specify-activation-rules-in-a-manifest).
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
- [Extraire des chaînes d’entité d’un élément Outlook](extract-entity-strings-from-an-item.md)
- [Obtenir des pièces jointes d’un élément Outlook à partir du serveur](get-attachments-of-an-outlook-item.md)

## <a name="see-also"></a>Voir aussi

- [Créer votre premier complément Outlook](../quickstarts/outlook-quickstart.md)
