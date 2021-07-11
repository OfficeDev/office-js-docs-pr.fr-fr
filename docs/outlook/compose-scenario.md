---
title: Créer des compléments Outlook pour les formulaires de composition
description: Découvrez les scénarios et fonctionnalités des compléments Outlook pour les formulaires de composition.
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 59ccebafbb3991ff3edb241596f44b5939d73693
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348530"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a><span data-ttu-id="c26bd-103">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="c26bd-103">Create Outlook add-ins for compose forms</span></span>

<span data-ttu-id="c26bd-p101">À partir de la version 1.1 du schéma pour les manifestes des Compléments Office et v1.1 Office.js, vous pouvez créer des compléments de composition, qui sont des compléments Outlook activés dans les formulaires de composition. Contrairement aux compléments de lecture (qui sont des compléments Outlook activés en mode lecture lorsqu’un utilisateur visualise un message ou un rendez-vous), les compléments de composition sont disponibles dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="c26bd-p101">Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of Office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms. In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios.</span></span>

- <span data-ttu-id="c26bd-106">Composition d’un nouveau message, d’une demande de réunion ou d’un rendez-vous dans un formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="c26bd-106">Composing a new message, meeting request, or appointment in a compose form.</span></span>

- <span data-ttu-id="c26bd-107">Affichage ou modification d’un rendez-vous existant, ou d’un élément de réunion dans lequel l’utilisateur est l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="c26bd-107">Viewing or editing an existing appointment, or meeting item in which the user is the organizer.</span></span>

   > [!NOTE]
   > <span data-ttu-id="c26bd-p102">Si l’utilisateur utilise la version RTM d’Outlook 2013 et d’Exchange 2013 et qu’il affiche un élément de réunion organisé par l’utilisateur, l’utilisateur peut rechercher les compléments de lecture disponibles. À partir de la version d’Office 2013 SP1, une modification a été apportée. Dans le même scénario, seuls les compléments de composition peuvent être activés et être disponibles.</span><span class="sxs-lookup"><span data-stu-id="c26bd-p102">If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.</span></span>

- <span data-ttu-id="c26bd-110">Composition d’un message de réponse inline ou réponse à un message dans un formulaire de composition individuel.</span><span class="sxs-lookup"><span data-stu-id="c26bd-110">Composing an inline response message or replying to a message in a separate compose form.</span></span>

- <span data-ttu-id="c26bd-111">Modification d’une réponse ( **Accepter**,  **Provisoire** ou **Refuser**) à une demande de réunion ou à un élément de réunion.</span><span class="sxs-lookup"><span data-stu-id="c26bd-111">Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.</span></span>

- <span data-ttu-id="c26bd-112">Proposition d’une nouvelle heure pour un élément de réunion.</span><span class="sxs-lookup"><span data-stu-id="c26bd-112">Proposing a new time for a meeting item.</span></span>

- <span data-ttu-id="c26bd-113">Transfert d’une demande de réunion ou d’un élément de réunion, ou réponse à une demande de réunion ou un élément de réunion.</span><span class="sxs-lookup"><span data-stu-id="c26bd-113">Forwarding or replying to a meeting request or meeting item.</span></span>

<span data-ttu-id="c26bd-p103">Dans chacun de ces scénarios de composition, tous les boutons de commande de complément sont affichés. Pour les compléments plus anciens qui n’implémentent pas les commandes de complément, les utilisateurs peuvent sélectionner **Compléments Office** dans le ruban pour ouvrir le volet de sélection des compléments, puis choisir et lancer un complément de composition. La figure suivante présente les commandes de complément dans un formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="c26bd-p103">In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.</span></span>

![Affiche un formulaire de composition Outlook avec les commandes de complément.](../images/compose-form-commands.png)

<span data-ttu-id="c26bd-118">La figure suivante présente le volet de sélection des compléments constitué de deux compléments de composition qui n’implémentent pas les commandes de complément, activés quand l’utilisateur compose une réponse instantanée dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="c26bd-118">The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.</span></span>

![Application de messagerie de modèles activée pour l’élément composé.](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a><span data-ttu-id="c26bd-120">Types de complément disponibles en mode composition</span><span class="sxs-lookup"><span data-stu-id="c26bd-120">Types of add-ins available in compose mode</span></span>

<span data-ttu-id="c26bd-121">Les compléments de composition sont implémentés en tant que [Commandes de complément pour Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="c26bd-121">Compose add-ins are implemented as [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span> <span data-ttu-id="c26bd-122">Pour activer les compléments pour la rédaction d’un e-mail ou de réponses à une demande de réunion, les compléments incluent un [élément de point d’extension MessageComposeCommandSurface](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="c26bd-122">To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) in the manifest.</span></span> <span data-ttu-id="c26bd-123">Pour activer les compléments pour composer ou modifier des rendez-vous ou des réunions dans lesquels l’utilisateur est l’organisateur, les compléments incluent un [élément de point d’extension AppointmentOrganizerCommandSurface](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span><span class="sxs-lookup"><span data-stu-id="c26bd-123">To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span></span>

> [!NOTE]
> <span data-ttu-id="c26bd-124">Les compléments développés pour des serveurs ou des clients ne prenant pas en charge les commandes de complément se servent de [règles d’activation](activation-rules.md) dans un élément [Règle](../reference/manifest/rule.md) contenu dans l’élément [OfficeApp](../reference/manifest/officeapp.md).</span><span class="sxs-lookup"><span data-stu-id="c26bd-124">Add-ins developed for servers or clients that do not support add-in commands use [activation rules](activation-rules.md) in a [Rule](../reference/manifest/rule.md) element contained in the [OfficeApp](../reference/manifest/officeapp.md) element.</span></span> <span data-ttu-id="c26bd-125">À moins que le complément ne soit développé spécifiquement pour des serveurs et clients plus anciens, les nouveaux compléments doivent utiliser les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="c26bd-125">Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.</span></span>

## <a name="api-features-available-to-compose-add-ins"></a><span data-ttu-id="c26bd-126">Fonctionnalités de l’API disponibles pour les compléments de composition</span><span class="sxs-lookup"><span data-stu-id="c26bd-126">API features available to compose add-ins</span></span>

- [<span data-ttu-id="c26bd-127">Ajouter et supprimer des pièces jointes à un élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-127">Add and remove attachments to an item in a compose form in Outlook</span></span>](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [<span data-ttu-id="c26bd-128">Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-128">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="c26bd-129">Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-129">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)
- [<span data-ttu-id="c26bd-130">Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-130">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="c26bd-131">Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-131">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="c26bd-132">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-132">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="c26bd-133">Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="c26bd-133">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a><span data-ttu-id="c26bd-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c26bd-134">See also</span></span>

- [<span data-ttu-id="c26bd-135">Prise en main des compléments Outlook pour Office</span><span class="sxs-lookup"><span data-stu-id="c26bd-135">Get Started with Outlook add-ins for Office</span></span>](../quickstarts/outlook-quickstart.md)
