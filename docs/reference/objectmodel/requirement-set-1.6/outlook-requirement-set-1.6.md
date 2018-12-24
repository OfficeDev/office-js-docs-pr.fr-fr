---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: e780cff1a4cfe0751fccc9192784d143ab9c483f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433681"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="bcacf-102">Ensemble de conditions requises de l’API du complément Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="bcacf-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="bcacf-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bcacf-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="bcacf-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="bcacf-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="bcacf-105">Nouveautés de la version 1.6</span><span class="sxs-lookup"><span data-stu-id="bcacf-105">What's new in 1.1?</span></span>

<span data-ttu-id="bcacf-106">L’ensemble de conditions requises de la version 1.6 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="bcacf-106">The Preview Requirement set includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="bcacf-107">Les fonctionnalités suivantes ont été ajoutées.</span><span class="sxs-lookup"><span data-stu-id="bcacf-107">It added the following features.</span></span>

- <span data-ttu-id="bcacf-108">Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.</span><span class="sxs-lookup"><span data-stu-id="bcacf-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="bcacf-109">La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="bcacf-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="bcacf-110">La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bcacf-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="bcacf-111">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="bcacf-111">Change log</span></span>

- <span data-ttu-id="bcacf-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bcacf-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities) - Added a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="bcacf-113">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="bcacf-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="bcacf-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="bcacf-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) - Added a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="bcacf-115">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="bcacf-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="bcacf-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="bcacf-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters) - Added a new function that opens a new message form.</span></span>
- <span data-ttu-id="bcacf-117">[Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bcacf-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="bcacf-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bcacf-118">See also</span></span>

- [<span data-ttu-id="bcacf-119">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="bcacf-119">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="bcacf-120">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="bcacf-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="bcacf-121">Prise en main</span><span class="sxs-lookup"><span data-stu-id="bcacf-121">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)