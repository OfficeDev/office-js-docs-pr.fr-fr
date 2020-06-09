---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,6.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 024b5ab992b146a1958653c38941434da00e1a03
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611448"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="fe329-103">Ensemble de conditions requises de l’API du complément Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="fe329-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="fe329-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe329-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="fe329-105">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="fe329-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="fe329-106">Nouveautés de la version 1.6</span><span class="sxs-lookup"><span data-stu-id="fe329-106">What's new in 1.6?</span></span>

<span data-ttu-id="fe329-107">L’ensemble de conditions requises de la version 1.6 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="fe329-107">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="fe329-108">Les fonctionnalités suivantes ont été ajoutées.</span><span class="sxs-lookup"><span data-stu-id="fe329-108">It added the following features.</span></span>

- <span data-ttu-id="fe329-109">Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.</span><span class="sxs-lookup"><span data-stu-id="fe329-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="fe329-110">La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="fe329-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="fe329-111">La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fe329-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="fe329-112">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="fe329-112">Change log</span></span>

- <span data-ttu-id="fe329-113">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fe329-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="fe329-114">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="fe329-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="fe329-115">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fe329-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="fe329-116">Les correspondances en surbrillance s’appliquent aux compléments contextuels.</span><span class="sxs-lookup"><span data-stu-id="fe329-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="fe329-117">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="fe329-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="fe329-118">[Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fe329-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="fe329-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fe329-119">See also</span></span>

- [<span data-ttu-id="fe329-120">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="fe329-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="fe329-121">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="fe329-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="fe329-122">Prise en main</span><span class="sxs-lookup"><span data-stu-id="fe329-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="fe329-123">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="fe329-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
