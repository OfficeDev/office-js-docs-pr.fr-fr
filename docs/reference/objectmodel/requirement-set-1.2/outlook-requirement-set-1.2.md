---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 1767b1b93f13de2c8a0731d2f08a1141b709b734
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068027"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="feeb6-102">Ensemble de conditions requises de l’API du complément Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="feeb6-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="feeb6-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="feeb6-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="feeb6-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="feeb6-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="feeb6-105">Nouveautés de la version 1.2</span><span class="sxs-lookup"><span data-stu-id="feeb6-105">What's new in 1.2?</span></span>

<span data-ttu-id="feeb6-p101">L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.</span><span class="sxs-lookup"><span data-stu-id="feeb6-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="feeb6-108">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="feeb6-108">Change log</span></span>

- <span data-ttu-id="feeb6-109">Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="feeb6-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="feeb6-110">Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="feeb6-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="feeb6-111">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="feeb6-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="feeb6-112">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback) : Ajout de la propriété `attachments` dans le paramètre `formData`.</span><span class="sxs-lookup"><span data-stu-id="feeb6-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="feeb6-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="feeb6-113">See also</span></span>

- [<span data-ttu-id="feeb6-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="feeb6-114">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="feeb6-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="feeb6-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="feeb6-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="feeb6-116">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
