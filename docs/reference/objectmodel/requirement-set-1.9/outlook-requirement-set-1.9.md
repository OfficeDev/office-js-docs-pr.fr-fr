---
title: Ensemble de conditions requises de l’API du complément Outlook 1,9
description: Ensemble de conditions requises 1,9 pour l’API de complément Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628052"
---
# <a name="outlook-add-in-api-requirement-set-19"></a><span data-ttu-id="6830c-103">Ensemble de conditions requises de l’API du complément Outlook 1,9</span><span class="sxs-lookup"><span data-stu-id="6830c-103">Outlook add-in API requirement set 1.9</span></span>

<span data-ttu-id="6830c-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="6830c-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties and events that you can use in an Outlook add-in.</span></span>

## <a name="whats-new-in-19"></a><span data-ttu-id="6830c-105">Quelles sont les nouveautés de 1,9 ?</span><span class="sxs-lookup"><span data-stu-id="6830c-105">What's new in 1.9?</span></span>

<span data-ttu-id="6830c-106">L’ensemble de conditions requises 1,9 inclut toutes les fonctionnalités de l' [ensemble de conditions requises 1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="6830c-106">Requirement set 1.9 includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="6830c-107">Les fonctionnalités suivantes ont été ajoutées.</span><span class="sxs-lookup"><span data-stu-id="6830c-107">It added the following features.</span></span>

- <span data-ttu-id="6830c-108">Ajout de nouvelles API pour les fonctionnalités ajout d’envoi, propriétés personnalisées et formulaire d’affichage.</span><span class="sxs-lookup"><span data-stu-id="6830c-108">Added new APIs for append-on-send, custom properties, and display form features.</span></span>
- <span data-ttu-id="6830c-109">Prise en charge supplémentaire de `Dialog.messageChild` .</span><span class="sxs-lookup"><span data-stu-id="6830c-109">Added support for `Dialog.messageChild`.</span></span>

### <a name="change-log"></a><span data-ttu-id="6830c-110">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="6830c-110">Change log</span></span>

- <span data-ttu-id="6830c-111">Ajout de la méthode [CustomProperties. GetAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): ajoute une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="6830c-111">Added [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): Adds a new function to the `CustomProperties` object that gets all custom properties.</span></span>
- <span data-ttu-id="6830c-112">Ajout de la méthode [Dialog. messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): ajoute une nouvelle méthode qui remet un message à partir de la page hôte, telle qu’un volet de tâches ou un fichier de fonctions sans interface utilisateur, à une boîte de dialogue ouverte à partir de la page.</span><span class="sxs-lookup"><span data-stu-id="6830c-112">Added [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): Adds a new method that delivers a message from the host page, such as a task pane or a UI-less function file, to a dialog that was opened from the page.</span></span>
- <span data-ttu-id="6830c-113">Ajout de l' [élément de manifeste ExtendedPermissions](../../manifest/extendedpermissions.md): ajoute un élément enfant à l’élément de manifeste [VersionOverrides](../../manifest/versionoverrides.md) .</span><span class="sxs-lookup"><span data-stu-id="6830c-113">Added [ExtendedPermissions manifest element](../../manifest/extendedpermissions.md): Adds a child element to the [VersionOverrides](../../manifest/versionoverrides.md) manifest element.</span></span> <span data-ttu-id="6830c-114">Pour qu’un complément prenne en charge la [fonctionnalité Append-on-Send](../../../outlook/append-on-send.md), l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="6830c-114">For an add-in to support the [append-on-send feature](../../../outlook/append-on-send.md), the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>
- <span data-ttu-id="6830c-115">Ajout de la méthode [Office. Context. Mailbox. displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un rendez-vous existant.</span><span class="sxs-lookup"><span data-stu-id="6830c-115">Added [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-): Adds a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="6830c-116">Il s’agit de la version asynchrone de la `displayAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-116">This is the async version of the `displayAppointmentForm` method.</span></span>
- <span data-ttu-id="6830c-117">Ajout de la méthode [Office. Context. Mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="6830c-117">Added [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-): Adds a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="6830c-118">Il s’agit de la version asynchrone de la `displayMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-118">This is the async version of the `displayMessageForm` method.</span></span>
- <span data-ttu-id="6830c-119">Ajout de la méthode [Office. Context. Mailbox. displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6830c-119">Added [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-): Adds a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="6830c-120">Il s’agit de la version asynchrone de la `displayNewAppointmentForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-120">This is the async version of the `displayNewAppointmentForm` method.</span></span>
- <span data-ttu-id="6830c-121">Ajout de la méthode [Office. Context. Mailbox. displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de message.</span><span class="sxs-lookup"><span data-stu-id="6830c-121">Added [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-): Adds a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="6830c-122">Il s’agit de la version asynchrone de la `displayNewMessageForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-122">This is the async version of the `displayNewMessageForm` method.</span></span>
- <span data-ttu-id="6830c-123">Ajout de la méthode [Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-): ajoute une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.</span><span class="sxs-lookup"><span data-stu-id="6830c-123">Added [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-): Adds a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>
- <span data-ttu-id="6830c-124">Ajout de la méthode [Office. Context. Mailbox. Item. displayReplyAllFormAsync](office.context.mailbox.item.md#methods): ajoute une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre à tous » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="6830c-124">Added [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods): Adds a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="6830c-125">Il s’agit de la version asynchrone de la `displayReplyAllForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-125">This is the async version of the `displayReplyAllForm` method.</span></span>
- <span data-ttu-id="6830c-126">Ajout de la méthode [Office. Context. Mailbox. Item. displayReplyFormAsync](office.context.mailbox.item.md#methods): ajoute une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre » en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="6830c-126">Added [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods): Adds a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="6830c-127">Il s’agit de la version asynchrone de la `displayReplyForm` méthode.</span><span class="sxs-lookup"><span data-stu-id="6830c-127">This is the async version of the `displayReplyForm` method.</span></span>

## <a name="see-also"></a><span data-ttu-id="6830c-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6830c-128">See also</span></span>

- [<span data-ttu-id="6830c-129">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="6830c-129">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="6830c-130">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="6830c-130">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="6830c-131">Prise en main</span><span class="sxs-lookup"><span data-stu-id="6830c-131">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="6830c-132">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="6830c-132">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)