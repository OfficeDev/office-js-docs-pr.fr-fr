---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,2.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 493a360a675a98ee6a364c67e83a60dc00bf35df
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610481"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="7250d-103">Ensemble de conditions requises de l’API du complément Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="7250d-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="7250d-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="7250d-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7250d-105">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="7250d-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="7250d-106">Nouveautés de la version 1.2</span><span class="sxs-lookup"><span data-stu-id="7250d-106">What's new in 1.2?</span></span>

<span data-ttu-id="7250d-p101">L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.</span><span class="sxs-lookup"><span data-stu-id="7250d-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="7250d-109">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="7250d-109">Change log</span></span>

- <span data-ttu-id="7250d-110">Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="7250d-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="7250d-111">Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7250d-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="7250d-112">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="7250d-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="7250d-113">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="7250d-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="7250d-114">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7250d-114">See also</span></span>

- [<span data-ttu-id="7250d-115">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="7250d-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="7250d-116">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="7250d-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7250d-117">Prise en main</span><span class="sxs-lookup"><span data-stu-id="7250d-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="7250d-118">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="7250d-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
