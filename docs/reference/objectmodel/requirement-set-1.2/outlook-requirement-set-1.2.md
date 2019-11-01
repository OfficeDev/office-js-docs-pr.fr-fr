---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 898e768dfc1828ba44f29e9da5c4baa61de186cb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902094"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="4a149-102">Ensemble de conditions requises de l’API du complément Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="4a149-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="4a149-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4a149-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4a149-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="4a149-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="4a149-105">Nouveautés de la version 1.2</span><span class="sxs-lookup"><span data-stu-id="4a149-105">What's new in 1.2?</span></span>

<span data-ttu-id="4a149-p101">L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4a149-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="4a149-108">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="4a149-108">Change log</span></span>

- <span data-ttu-id="4a149-109">Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4a149-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="4a149-110">Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4a149-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="4a149-111">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="4a149-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="4a149-112">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="4a149-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="4a149-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4a149-113">See also</span></span>

- [<span data-ttu-id="4a149-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="4a149-114">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="4a149-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="4a149-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="4a149-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="4a149-116">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="4a149-117">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="4a149-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
