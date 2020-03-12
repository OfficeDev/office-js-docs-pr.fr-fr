---
title: Ensemble de conditions requises de l’API du complément Outlook 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f56cf4e13bdf3518ef14da6eca83b51abe82e50c
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597046"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="4c5b3-102">Ensemble de conditions requises de l’API du complément Outlook 1.5</span><span class="sxs-lookup"><span data-stu-id="4c5b3-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="4c5b3-103">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4c5b3-104">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="4c5b3-105">Nouveautés de la version 1.5</span><span class="sxs-lookup"><span data-stu-id="4c5b3-105">What's new in 1.5?</span></span>

<span data-ttu-id="4c5b3-p101">L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées :</span><span class="sxs-lookup"><span data-stu-id="4c5b3-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="4c5b3-108">Prise en charge des [volets Office épinglables](../../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="4c5b3-108">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="4c5b3-109">Prise en charge de l’appel des [API REST](../../../outlook/use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="4c5b3-109">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="4c5b3-110">Possibilité de marquer une pièce jointe comme élément incorporé.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="4c5b3-111">Possibilité de fermer un volet Office ou une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="4c5b3-112">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="4c5b3-112">Change log</span></span>

- <span data-ttu-id="4c5b3-113">Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) : ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="4c5b3-114">Ajout de la méthode [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="4c5b3-115">Ajout de l’énumération [Office.EventType](office.md#eventtype-string) : spécifie l’événement associé à un gestionnaire d’événements et prend en charge l’événement ItemChanger.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="4c5b3-116">Ajout de la propriété [Office.context.mailbox.restUrl](office.context.mailbox.md#properties) : obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="4c5b3-p102">Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="4c5b3-119">Ajout de la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="4c5b3-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="4c5b3-120">Modification de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="4c5b3-121">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="4c5b3-122">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4c5b3-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="4c5b3-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4c5b3-123">See also</span></span>

- [<span data-ttu-id="4c5b3-124">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="4c5b3-124">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="4c5b3-125">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="4c5b3-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="4c5b3-126">Prise en main</span><span class="sxs-lookup"><span data-stu-id="4c5b3-126">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="4c5b3-127">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="4c5b3-127">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
