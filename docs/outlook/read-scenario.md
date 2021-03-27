---
title: Créer des compléments Outlook pour des formulaires de lecture
description: Les compléments de lecture sont des compléments Outlook qui sont activés dans le volet de lecture ou l’inspecteur de lecture dans Outlook.
ms.date: 03/19/2021
localization_priority: Priority
ms.openlocfilehash: 495b4d947ec965481859c3262d3b67b93f57a5c0
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51177998"
---
# <a name="create-outlook-add-ins-for-read-forms"></a><span data-ttu-id="8dca1-103">Créer des compléments Outlook pour des formulaires de lecture</span><span class="sxs-lookup"><span data-stu-id="8dca1-103">Create Outlook add-ins for read forms</span></span>

<span data-ttu-id="8dca1-p101">Les compléments de lecture sont des compléments Outlook activés dans le volet de lecture ou l’inspecteur de lecture d’Outlook. Contrairement aux compléments de composition (qui sont des compléments Outlook activés lorsqu’un utilisateur crée un message ou un rendez-vous), les compléments de lecture sont disponibles dans les scénarios suivants :</span><span class="sxs-lookup"><span data-stu-id="8dca1-p101">Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:</span></span>

- <span data-ttu-id="8dca1-106">Affichage d’un message électronique, d’une demande de réunion, d’une réponse à une demande de réunion ou d’une annulation de réunion.</span><span class="sxs-lookup"><span data-stu-id="8dca1-106">View an email message, meeting request, meeting response, or meeting cancellation.</span></span>

   > [!NOTE]
   > <span data-ttu-id="8dca1-107">Outlook n’active pas les compléments dans un formulaire de lecture pour certains types de messages, y compris les éléments qui sont les pièces jointes d’un autre message, les éléments du dossier Brouillons, ou encore ceux chiffrés ou protégés d’autres façons.</span><span class="sxs-lookup"><span data-stu-id="8dca1-107">Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.</span></span>

- <span data-ttu-id="8dca1-108">Affichage d’un élément de réunion dans lequel l’utilisateur est un participant.</span><span class="sxs-lookup"><span data-stu-id="8dca1-108">View a meeting item in which the user is an attendee.</span></span>

- <span data-ttu-id="8dca1-109">Affichage d’un élément de réunion dans lequel l’utilisateur est l’organisateur (version RTM d’Outlook 2013 et d’Exchange 2013 uniquement).</span><span class="sxs-lookup"><span data-stu-id="8dca1-109">View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).</span></span>

   > [!NOTE]
   > <span data-ttu-id="8dca1-p102">À partir de la version Office 2013 SP1, si l’utilisateur visualise un élément de réunion dont il est l’organisateur, seuls les compléments de composition peuvent être activés et disponibles. Les compléments de lecture ne sont plus disponibles dans ce scénario.</span><span class="sxs-lookup"><span data-stu-id="8dca1-p102">Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.</span></span>

<span data-ttu-id="8dca1-p103">Dans chacun de ces scénarios de lecture, Outlook active les compléments lorsque leurs conditions d’activation sont respectées. Les utilisateurs peuvent ensuite choisir et ouvrir les compléments activés dans la barre de compléments du volet de lecture ou de l’inspecteur de lecture. La figure suivante montre le complément **Bing Cartes** qui a été activé et ouvert alors que l’utilisateur lit un message contenant une adresse géographique.</span><span class="sxs-lookup"><span data-stu-id="8dca1-p103">In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.</span></span>

<span data-ttu-id="8dca1-114">**Volet de complément montrant le complément Bing Cartes en action pour le message Outlook sélectionné qui contient une adresse**</span><span class="sxs-lookup"><span data-stu-id="8dca1-114">**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**</span></span>

![Application de messagerie avec carte Bing dans Outlook](../images/outlook-detected-entity-card.png)

## <a name="types-of-add-ins-available-in-read-mode"></a><span data-ttu-id="8dca1-116">Types de complément disponibles en mode de lecture</span><span class="sxs-lookup"><span data-stu-id="8dca1-116">Types of add-ins available in read mode</span></span>

<span data-ttu-id="8dca1-117">Les compléments de lecture peuvent correspondre à n’importe quelle combinaison des types suivants.</span><span class="sxs-lookup"><span data-stu-id="8dca1-117">Read add-ins can be any combination of the following types.</span></span>

- [<span data-ttu-id="8dca1-118">Commandes de complément pour Outlook</span><span class="sxs-lookup"><span data-stu-id="8dca1-118">Add-in commands for Outlook</span></span>](add-in-commands-for-outlook.md)
- [<span data-ttu-id="8dca1-119">Compléments Outlook contextuels</span><span class="sxs-lookup"><span data-stu-id="8dca1-119">Contextual Outlook add-ins</span></span>](contextual-outlook-add-ins.md)

## <a name="api-features-available-to-read-add-ins"></a><span data-ttu-id="8dca1-120">Fonctionnalités de l’API disponibles pour les compléments de lecture</span><span class="sxs-lookup"><span data-stu-id="8dca1-120">API features available to read add-ins</span></span>

- <span data-ttu-id="8dca1-121">Pour activer les compléments dans les formulaires de lecture, voir le tableau 1 dans [Spécifier des règles d’activation dans un manifeste](activation-rules.md#specify-activation-rules-in-a-manifest).</span><span class="sxs-lookup"><span data-stu-id="8dca1-121">For activating add-ins in read forms, see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).</span></span>
- [<span data-ttu-id="8dca1-122">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="8dca1-122">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="8dca1-123">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="8dca1-123">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="8dca1-124">Extraire des chaînes d’entité d’un élément Outlook</span><span class="sxs-lookup"><span data-stu-id="8dca1-124">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)
- [<span data-ttu-id="8dca1-125">Obtenir des pièces jointes d’un élément Outlook à partir du serveur</span><span class="sxs-lookup"><span data-stu-id="8dca1-125">Get attachments of an Outlook item from the server</span></span>](get-attachments-of-an-outlook-item.md)

## <a name="see-also"></a><span data-ttu-id="8dca1-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8dca1-126">See also</span></span>

- [<span data-ttu-id="8dca1-127">Créer votre premier complément Outlook</span><span class="sxs-lookup"><span data-stu-id="8dca1-127">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
