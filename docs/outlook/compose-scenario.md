---
title: Créer des compléments Outlook pour les formulaires de composition
description: Découvrez les scénarios et fonctionnalités des compléments Outlook pour les formulaires de composition.
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: ea85194eb74e0eb57addecddab32fe157cf2a88ef05604dfe1b1992678973996
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090865"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>Créer des compléments Outlook pour les formulaires de composition

À partir de la version 1.1 du schéma pour les manifestes des Compléments Office et v1.1 Office.js, vous pouvez créer des compléments de composition, qui sont des compléments Outlook activés dans les formulaires de composition. Contrairement aux compléments de lecture (qui sont des compléments Outlook activés en mode lecture lorsqu’un utilisateur visualise un message ou un rendez-vous), les compléments de composition sont disponibles dans les scénarios suivants.

- Composition d’un nouveau message, d’une demande de réunion ou d’un rendez-vous dans un formulaire de composition.

- Affichage ou modification d’un rendez-vous existant, ou d’un élément de réunion dans lequel l’utilisateur est l’organisateur.

   > [!NOTE]
   > Si l’utilisateur utilise la version RTM d’Outlook 2013 et d’Exchange 2013 et qu’il affiche un élément de réunion organisé par l’utilisateur, l’utilisateur peut rechercher les compléments de lecture disponibles. À partir de la version d’Office 2013 SP1, une modification a été apportée. Dans le même scénario, seuls les compléments de composition peuvent être activés et être disponibles.

- Composition d’un message de réponse inline ou réponse à un message dans un formulaire de composition individuel.

- Modification d’une réponse ( **Accepter**,  **Provisoire** ou **Refuser**) à une demande de réunion ou à un élément de réunion.

- Proposition d’une nouvelle heure pour un élément de réunion.

- Transfert d’une demande de réunion ou d’un élément de réunion, ou réponse à une demande de réunion ou un élément de réunion.

Dans chacun de ces scénarios de composition, tous les boutons de commande de complément sont affichés. Pour les compléments plus anciens qui n’implémentent pas les commandes de complément, les utilisateurs peuvent sélectionner **Compléments Office** dans le ruban pour ouvrir le volet de sélection des compléments, puis choisir et lancer un complément de composition. La figure suivante présente les commandes de complément dans un formulaire de composition.

![Affiche un formulaire de composition Outlook avec les commandes de complément.](../images/compose-form-commands.png)

La figure suivante présente le volet de sélection des compléments constitué de deux compléments de composition qui n’implémentent pas les commandes de complément, activés quand l’utilisateur compose une réponse instantanée dans Outlook.

![Application de messagerie de modèles activée pour l’élément composé.](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>Types de complément disponibles en mode composition

Les compléments de composition sont implémentés en tant que [Commandes de complément pour Outlook](add-in-commands-for-outlook.md). Pour activer les compléments pour la rédaction d’un e-mail ou de réponses à une demande de réunion, les compléments incluent un [élément de point d’extension MessageComposeCommandSurface](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) dans le manifeste. Pour activer les compléments pour composer ou modifier des rendez-vous ou des réunions dans lesquels l’utilisateur est l’organisateur, les compléments incluent un [élément de point d’extension AppointmentOrganizerCommandSurface](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).

> [!NOTE]
> Les compléments développés pour des serveurs ou des clients ne prenant pas en charge les commandes de complément se servent de [règles d’activation](activation-rules.md) dans un élément [Règle](../reference/manifest/rule.md) contenu dans l’élément [OfficeApp](../reference/manifest/officeapp.md). À moins que le complément ne soit développé spécifiquement pour des serveurs et clients plus anciens, les nouveaux compléments doivent utiliser les commandes de complément.

## <a name="api-features-available-to-compose-add-ins"></a>Fonctionnalités de l’API disponibles pour les compléments de composition

- [Ajouter et supprimer des pièces jointes à un élément dans un formulaire de composition dans Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-set-or-add-recipients.md)
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-or-set-the-subject.md)
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](insert-data-in-the-body.md)
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md)
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a>Voir aussi

- [Prise en main des compléments Outlook pour Office](../quickstarts/outlook-quickstart.md)
