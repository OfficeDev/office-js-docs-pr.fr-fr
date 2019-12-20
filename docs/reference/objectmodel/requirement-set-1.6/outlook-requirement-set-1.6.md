---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 22702448b82a108c401f9f81d3b8a321e14ead63
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814660"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="8702c-102">Ensemble de conditions requises de l’API du complément Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="8702c-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="8702c-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="8702c-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8702c-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="8702c-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="8702c-105">Nouveautés de la version 1.6</span><span class="sxs-lookup"><span data-stu-id="8702c-105">What's new in 1.6?</span></span>

<span data-ttu-id="8702c-106">L’ensemble de conditions requises de la version 1.6 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="8702c-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="8702c-107">Les fonctionnalités suivantes ont été ajoutées.</span><span class="sxs-lookup"><span data-stu-id="8702c-107">It added the following features.</span></span>

- <span data-ttu-id="8702c-108">Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.</span><span class="sxs-lookup"><span data-stu-id="8702c-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="8702c-109">La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="8702c-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="8702c-110">La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8702c-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="8702c-111">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="8702c-111">Change log</span></span>

- <span data-ttu-id="8702c-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8702c-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="8702c-113">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="8702c-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="8702c-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8702c-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="8702c-115">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="8702c-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="8702c-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="8702c-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="8702c-117">[Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8702c-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="8702c-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8702c-118">See also</span></span>

- [<span data-ttu-id="8702c-119">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="8702c-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="8702c-120">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="8702c-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="8702c-121">Prise en main</span><span class="sxs-lookup"><span data-stu-id="8702c-121">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="8702c-122">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="8702c-122">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
