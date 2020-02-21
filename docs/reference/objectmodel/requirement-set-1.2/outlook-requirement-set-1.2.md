---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: d46b705c79283049b3dbdff19b8348aa1b3c7bb0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163849"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="890ef-102">Ensemble de conditions requises de l’API du complément Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="890ef-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="890ef-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="890ef-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="890ef-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="890ef-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="890ef-105">Nouveautés de la version 1.2</span><span class="sxs-lookup"><span data-stu-id="890ef-105">What's new in 1.2?</span></span>

<span data-ttu-id="890ef-p101">L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.</span><span class="sxs-lookup"><span data-stu-id="890ef-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="890ef-108">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="890ef-108">Change log</span></span>

- <span data-ttu-id="890ef-109">Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="890ef-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="890ef-110">Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="890ef-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="890ef-111">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="890ef-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="890ef-112">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="890ef-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="890ef-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="890ef-113">See also</span></span>

- [<span data-ttu-id="890ef-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="890ef-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="890ef-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="890ef-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="890ef-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="890ef-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="890ef-117">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="890ef-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
