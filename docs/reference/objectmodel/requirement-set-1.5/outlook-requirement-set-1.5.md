---
title: Ensemble de conditions requises de l’API du complément Outlook 1.5
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,5.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: bc91ea93a6c3653dd326306139ee460132412a81
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612037"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="2a2dd-103">Ensemble de conditions requises de l’API du complément Outlook 1.5</span><span class="sxs-lookup"><span data-stu-id="2a2dd-103">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="2a2dd-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2a2dd-105">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="2a2dd-106">Nouveautés de la version 1.5</span><span class="sxs-lookup"><span data-stu-id="2a2dd-106">What's new in 1.5?</span></span>

<span data-ttu-id="2a2dd-p101">L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées :</span><span class="sxs-lookup"><span data-stu-id="2a2dd-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="2a2dd-109">Prise en charge des [volets Office épinglables](../../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="2a2dd-109">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="2a2dd-110">Prise en charge de l’appel des [API REST](../../../outlook/use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="2a2dd-110">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="2a2dd-111">Possibilité de marquer une pièce jointe comme élément incorporé.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-111">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="2a2dd-112">Possibilité de fermer un volet Office ou une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-112">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="2a2dd-113">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="2a2dd-113">Change log</span></span>

- <span data-ttu-id="2a2dd-114">Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) : ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-114">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="2a2dd-115">Ajout de la méthode [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-115">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="2a2dd-116">Ajout de l’énumération [Office.EventType](office.md#eventtype-string) : spécifie l’événement associé à un gestionnaire d’événements et prend en charge l’événement ItemChanger.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-116">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="2a2dd-117">Ajout de la propriété [Office.context.mailbox.restUrl](office.context.mailbox.md#properties) : obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-117">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="2a2dd-p102">Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="2a2dd-120">Ajout de la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="2a2dd-120">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="2a2dd-121">Modification de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-121">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="2a2dd-122">Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-122">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="2a2dd-123">Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="2a2dd-123">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="2a2dd-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2a2dd-124">See also</span></span>

- [<span data-ttu-id="2a2dd-125">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2a2dd-125">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2a2dd-126">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2a2dd-126">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="2a2dd-127">Prise en main</span><span class="sxs-lookup"><span data-stu-id="2a2dd-127">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="2a2dd-128">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="2a2dd-128">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
