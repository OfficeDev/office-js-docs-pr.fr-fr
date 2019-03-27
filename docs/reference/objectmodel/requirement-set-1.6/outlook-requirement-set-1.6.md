---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871325"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="200dd-102">Ensemble de conditions requises de l’API du complément Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="200dd-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="200dd-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="200dd-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="200dd-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="200dd-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="200dd-105">Nouveautés de la version 1.6</span><span class="sxs-lookup"><span data-stu-id="200dd-105">What's new in 1.6?</span></span>

<span data-ttu-id="200dd-106">L’ensemble de conditions requises de la version 1.6 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="200dd-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="200dd-107">Les fonctionnalités suivantes ont été ajoutées.</span><span class="sxs-lookup"><span data-stu-id="200dd-107">It added the following features.</span></span>

- <span data-ttu-id="200dd-108">Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.</span><span class="sxs-lookup"><span data-stu-id="200dd-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="200dd-109">La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="200dd-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="200dd-110">La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="200dd-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="200dd-111">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="200dd-111">Change log</span></span>

- <span data-ttu-id="200dd-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="200dd-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="200dd-113">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="200dd-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="200dd-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="200dd-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="200dd-115">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="200dd-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="200dd-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="200dd-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="200dd-117">[Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="200dd-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="200dd-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="200dd-118">See also</span></span>

- [<span data-ttu-id="200dd-119">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="200dd-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="200dd-120">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="200dd-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="200dd-121">Prise en main</span><span class="sxs-lookup"><span data-stu-id="200dd-121">Get started</span></span>](/outlook/add-ins/quick-start)
