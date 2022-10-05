---
title: Créer des compléments Outlook pour les formulaires de composition
description: Découvrez les scénarios et fonctionnalités des compléments Outlook pour les formulaires de composition.
ms.date: 10/03/2022
ms.localizationpriority: high
ms.openlocfilehash: ef81b21eaa0bc63a5bf38757cb188e8850ade443
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467250"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>Créer des compléments Outlook pour les formulaires de composition

Vous pouvez créer des compléments de composition, qui sont des compléments Outlook activés dans les formulaires de composition. Contrairement aux compléments en lecture (compléments Outlook activés en mode lecture lorsqu’un utilisateur affiche un message ou un rendez-vous), les compléments de composition sont disponibles dans les scénarios utilisateur suivants.

- Composition d’un nouveau message, d’une demande de réunion ou d’un rendez-vous dans un formulaire de composition.

- Affichage ou modification d’un rendez-vous existant, ou d’un élément de réunion dans lequel l’utilisateur est l’organisateur.

   > [!NOTE]
   > If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.

- Composition d’un message de réponse inline ou réponse à un message dans un formulaire de composition individuel.

- Modification d’une réponse ( **Accepter**,  **Provisoire** ou **Refuser**) à une demande de réunion ou à un élément de réunion.

- Proposition d’une nouvelle heure pour un élément de réunion.

- Transfert d’une demande de réunion ou d’un élément de réunion, ou réponse à une demande de réunion ou un élément de réunion.

In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.

![Affiche un formulaire de composition Outlook avec les commandes de complément.](../images/compose-form-commands.png)

La figure suivante présente le volet de sélection des compléments constitué de deux compléments de composition qui n’implémentent pas les commandes de complément, activés quand l’utilisateur compose une réponse instantanée dans Outlook.

![Application de messagerie de modèles activée pour l’élément composé.](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>Types de complément disponibles en mode composition

Les compléments de composition sont implémentés en tant que [Commandes de complément pour Outlook](add-in-commands-for-outlook.md). Pour activer les compléments pour la rédaction d’un e-mail ou de réponses à une demande de réunion, les compléments incluent un [élément de point d’extension MessageComposeCommandSurface](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface) dans le manifeste. Pour activer les compléments pour composer ou modifier des rendez-vous ou des réunions dans lesquels l’utilisateur est l’organisateur, les compléments incluent un [élément de point d’extension AppointmentOrganizerCommandSurface](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface).

> [!NOTE]
> Les compléments développés pour des serveurs ou des clients ne prenant pas en charge les commandes de complément se servent de [règles d’activation](activation-rules.md) dans un élément [Règle](/javascript/api/manifest/rule) contenu dans l’élément [OfficeApp](/javascript/api/manifest/officeapp). À moins que le complément ne soit développé spécifiquement pour des serveurs et clients plus anciens, les nouveaux compléments doivent utiliser les commandes de complément.

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
