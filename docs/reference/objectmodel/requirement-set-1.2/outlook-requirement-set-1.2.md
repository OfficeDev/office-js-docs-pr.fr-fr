---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: d4fa18f3ab12e22ff30ef841d921f5dac89fd064
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325212"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="29ff7-102">Ensemble de conditions requises de l’API du complément Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="29ff7-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="29ff7-103">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="29ff7-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="29ff7-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="29ff7-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="29ff7-105">Nouveautés de la version 1.2</span><span class="sxs-lookup"><span data-stu-id="29ff7-105">What's new in 1.2?</span></span>

<span data-ttu-id="29ff7-p101">L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.</span><span class="sxs-lookup"><span data-stu-id="29ff7-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="29ff7-108">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="29ff7-108">Change log</span></span>

- <span data-ttu-id="29ff7-109">Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="29ff7-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="29ff7-110">Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="29ff7-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="29ff7-111">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="29ff7-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="29ff7-112">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="29ff7-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="29ff7-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="29ff7-113">See also</span></span>

- [<span data-ttu-id="29ff7-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="29ff7-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="29ff7-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="29ff7-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="29ff7-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="29ff7-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="29ff7-117">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="29ff7-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
