---
title: Présentation des autorisations de complément Outlook
description: Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste (Restricted, ReadItem, ReadWriteItem ou ReadWriteMailbox).
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 60b65416585b5215ed565a3689c1e7f398e001a5
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325325"
---
# <a name="understanding-outlook-add-in-permissions"></a><span data-ttu-id="e692f-103">Présentation des autorisations de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="e692f-103">Understanding Outlook add-in permissions</span></span>

<span data-ttu-id="e692f-p101">Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste. Les niveaux disponibles sont **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**. Ces niveaux d’autorisation sont cumulatifs : **Restricted** est le niveau le plus bas, et chaque niveau supérieur inclut les autorisations de tous les niveaux inférieurs. **ReadWriteMailbox** contient toutes les autorisations prises en charge.</span><span class="sxs-lookup"><span data-stu-id="e692f-p101">Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.</span></span>

<span data-ttu-id="e692f-p102">Vous pouvez voir les autorisations demandées par un complément de messagerie avant de l’installer depuis [AppSource](https://appsource.microsoft.com). Vous pouvez également voir les autorisations requises des compléments installés dans le Centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="e692f-p102">You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.</span></span>

## <a name="restricted-permission"></a><span data-ttu-id="e692f-110">Autorisation Restricted</span><span class="sxs-lookup"><span data-stu-id="e692f-110">Restricted permission</span></span>

<span data-ttu-id="e692f-p103">L’autorisation **Restricted** est la plus basique. Indiquez **Restricted** dans l’élément [Permissions](../reference/manifest/permissions.md) du manifeste pour demander cette autorisation. Outlook affecte par défaut ce niveau d’autorisation à un complément de messagerie si le complément ne demande pas d’autorisation spécifique dans son manifeste.</span><span class="sxs-lookup"><span data-stu-id="e692f-p103">The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.</span></span>

### <a name="can-do"></a><span data-ttu-id="e692f-114">Vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="e692f-114">Can do</span></span>

- <span data-ttu-id="e692f-115">[Obtenir uniquement des entités spécifiques](match-strings-in-an-item-as-well-known-entities.md) (numéro de téléphone, adresse, URL) de l’objet ou du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="e692f-115">[Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.</span></span>

- <span data-ttu-id="e692f-116">Spécifier une [règle d’activation ItemIs](activation-rules.md#itemis-rule) qui exige que l’élément actuel soit un type d’élément spécifique dans un formulaire de lecture ou de composition, ou une [règle ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md) qui correspond à l’un des sous-ensembles plus petits d’entités connues prises en charge (numéro de téléphone, adresse, URL) dans l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="e692f-116">Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.</span></span>

- <span data-ttu-id="e692f-117">Accéder aux propriétés et aux méthodes qui ne sont **pas** associées aux informations spécifiques concernant l’utilisateur ou l’élément. (Consultez la section suivante pour obtenir la liste des membres qui le sont.)</span><span class="sxs-lookup"><span data-stu-id="e692f-117">Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).</span></span>

### <a name="cant-do"></a><span data-ttu-id="e692f-118">Vous ne pouvez pas :</span><span class="sxs-lookup"><span data-stu-id="e692f-118">Can't do</span></span>

- <span data-ttu-id="e692f-119">Utiliser une règle [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) sur l’entité contact, adresse de messagerie, suggestion de réunion ou suggestion de tâche.</span><span class="sxs-lookup"><span data-stu-id="e692f-119">Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entity.</span></span>

- <span data-ttu-id="e692f-120">Utiliser la règle [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) ou [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule).</span><span class="sxs-lookup"><span data-stu-id="e692f-120">Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.</span></span>

- <span data-ttu-id="e692f-p104">Accéder aux membres de la liste suivante qui se rapportent aux informations de l’utilisateur ou de l’élément. Si vous tentez d’accéder aux membres de cette liste, vous obtenez la valeur **null** et un message d’erreur indiquant qu’Outlook requiert le complément de messagerie pour bénéficier d’autorisations élevées.</span><span class="sxs-lookup"><span data-stu-id="e692f-p104">Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.</span></span>

    - [<span data-ttu-id="e692f-123">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-123">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-124">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-124">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-125">item.attachments</span><span class="sxs-lookup"><span data-stu-id="e692f-125">item.attachments</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-126">item.bcc</span><span class="sxs-lookup"><span data-stu-id="e692f-126">item.bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-127">item.body</span><span class="sxs-lookup"><span data-stu-id="e692f-127">item.body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-128">item.cc</span><span class="sxs-lookup"><span data-stu-id="e692f-128">item.cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-129">item.from</span><span class="sxs-lookup"><span data-stu-id="e692f-129">item.from</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-130">item.getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e692f-130">item.getRegExMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-131">item.getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e692f-131">item.getRegExMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-132">item.optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e692f-132">item.optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-133">item.organizer</span><span class="sxs-lookup"><span data-stu-id="e692f-133">item.organizer</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-134">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-134">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-135">item.requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e692f-135">item.requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-136">item.sender</span><span class="sxs-lookup"><span data-stu-id="e692f-136">item.sender</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-137">item.to</span><span class="sxs-lookup"><span data-stu-id="e692f-137">item.to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="e692f-138">mailbox.getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-138">mailbox.getCallbackTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="e692f-139">mailbox.getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-139">mailbox.getUserIdentityTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="e692f-140">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-140">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="e692f-141">mailbox.userProfile</span><span class="sxs-lookup"><span data-stu-id="e692f-141">mailbox.userProfile</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - <span data-ttu-id="e692f-142">[Body](/javascript/api/outlook/office.body) et tous ses membres enfants</span><span class="sxs-lookup"><span data-stu-id="e692f-142">[Body](/javascript/api/outlook/office.body) and all its child members</span></span>
    - <span data-ttu-id="e692f-143">[Location](/javascript/api/outlook/office.location) et tous ses membres enfants</span><span class="sxs-lookup"><span data-stu-id="e692f-143">[Location](/javascript/api/outlook/office.location) and all its child members</span></span>
    - <span data-ttu-id="e692f-144">[Recipients](/javascript/api/outlook/office.recipients) et tous ses membres enfants</span><span class="sxs-lookup"><span data-stu-id="e692f-144">[Recipients](/javascript/api/outlook/office.recipients) and all its child members</span></span>
    - <span data-ttu-id="e692f-145">[Subject](/javascript/api/outlook/office.subject) et tous ses membres enfants</span><span class="sxs-lookup"><span data-stu-id="e692f-145">[Subject](/javascript/api/outlook/office.subject) and all its child members</span></span>
    - <span data-ttu-id="e692f-146">[Time](/javascript/api/outlook/office.time) et tous ses membres enfants</span><span class="sxs-lookup"><span data-stu-id="e692f-146">[Time](/javascript/api/outlook/office.time) and all its child members</span></span>

## <a name="readitem-permission"></a><span data-ttu-id="e692f-147">Autorisation ReadItem</span><span class="sxs-lookup"><span data-stu-id="e692f-147">ReadItem permission</span></span>

<span data-ttu-id="e692f-p105">L’autorisation **ReadItem** est le niveau suivant dans le modèle d’autorisations. Indiquez **ReadItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation.</span><span class="sxs-lookup"><span data-stu-id="e692f-p105">The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="e692f-150">Vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="e692f-150">Can do</span></span>

- <span data-ttu-id="e692f-151">[Lire toutes les propriétés](item-data.md) de l’élément actuel dans un formulaire de lecture ou de [composition](get-and-set-item-data-in-a-compose-form.md), par exemple, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) dans un formulaire de lecture et [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) dans un formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="e692f-151">[Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.</span></span>

- <span data-ttu-id="e692f-152">[Obtenir un jeton de rappel pour obtenir les pièces jointes de l’élément](get-attachments-of-an-outlook-item.md) ou l’élément complet avec les services Web Exchange ou les [API REST Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="e692f-152">[Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).</span></span>

- <span data-ttu-id="e692f-153">[Écrire des propriétés personnalisées](/javascript/api/outlook/office.CustomProperties) définies par le complément sur cet élément.</span><span class="sxs-lookup"><span data-stu-id="e692f-153">[Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.</span></span>

- <span data-ttu-id="e692f-154">[Obtenir toutes les entités existantes connues](match-strings-in-an-item-as-well-known-entities.md), et pas seulement un sous-ensemble, à partir de l’objet ou du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="e692f-154">[Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.</span></span>

- <span data-ttu-id="e692f-p106">Utiliser toutes les [entités connues](activation-rules.md#itemhasknownentity-rule) dans les règles [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ou les [expressions régulières](activation-rules.md#itemhasregularexpressionmatch-rule) dans les règles [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule). L’exemple suivant suit le schéma version 1.1. Il montre une règle qui active le complément si une ou plusieurs entités connues sont trouvées dans l’objet ou le corps du message sélectionné :</span><span class="sxs-lookup"><span data-stu-id="e692f-p106">Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a><span data-ttu-id="e692f-158">Vous ne pouvez pas :</span><span class="sxs-lookup"><span data-stu-id="e692f-158">Can't do</span></span>

- <span data-ttu-id="e692f-159">Utilisez le jeton fourni par **mailbox.getCallbackTokenAsync** pour les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="e692f-159">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="e692f-160">Mettre à jour ou supprimer l’élément actuel à l’aide de l’API REST Outlook ou accéder à tous les autres éléments de la boîte aux lettres de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="e692f-160">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="e692f-161">Récupérer l’élément d’événement de calendrier actuel à l’aide de l’API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="e692f-161">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="e692f-162">Utilisez l’une des API suivantes :</span><span class="sxs-lookup"><span data-stu-id="e692f-162">Use any of the following APIs:</span></span>
    - [<span data-ttu-id="e692f-163">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-163">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="e692f-164">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-164">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-165">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-165">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-166">item.bcc.addAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-166">item.bcc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-167">item.bcc.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-167">item.bcc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-168">item.body.prependAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-168">item.body.prependAsync</span></span>](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [<span data-ttu-id="e692f-169">item.body.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-169">item.body.setAsync</span></span>](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [<span data-ttu-id="e692f-170">item.body.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-170">item.body.setSelectedDataAsync</span></span>](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [<span data-ttu-id="e692f-171">item.cc.addAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-171">item.cc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-172">item.cc.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-172">item.cc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-173">item.end.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-173">item.end.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="e692f-174">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-174">item.location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [<span data-ttu-id="e692f-175">item.optionalAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-175">item.optionalAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-176">item.optionalAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-176">item.optionalAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-177">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-177">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="e692f-178">item.requiredAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-178">item.requiredAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-179">item.requiredAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-179">item.requiredAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-180">item.start.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-180">item.start.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="e692f-181">item.subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-181">item.subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [<span data-ttu-id="e692f-182">item.to.addAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-182">item.to.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="e692f-183">item.to.setAsync</span><span class="sxs-lookup"><span data-stu-id="e692f-183">item.to.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a><span data-ttu-id="e692f-184">Autorisation ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e692f-184">ReadWriteItem permission</span></span>

<span data-ttu-id="e692f-p107">Vous pouvez indiquer **ReadWriteItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation. Les compléments de messagerie activés dans des formulaires de composition et utilisant des méthodes d’écriture (par exemple, **Message.to.addAsync** ou **Message.to.setAsync**) doivent utiliser au moins ce niveau d’autorisation.</span><span class="sxs-lookup"><span data-stu-id="e692f-p107">Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="e692f-187">Vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="e692f-187">Can do</span></span>

- <span data-ttu-id="e692f-188">[Lire et écrire toutes les propriétés au niveau de l’élément](item-data.md) concernant l’élément affiché ou en cours de composition dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="e692f-188">[Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.</span></span>

- <span data-ttu-id="e692f-189">[Ajouter ou supprimer des pièces jointes](add-and-remove-attachments-to-an-item-in-a-compose-form.md) de cet élément.</span><span class="sxs-lookup"><span data-stu-id="e692f-189">[Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.</span></span>

- <span data-ttu-id="e692f-190">Utilisez tous les autres membres de l’API JavaScript pour Office qui s’appliquent aux compléments de messagerie, sauf **Mailbox. makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="e692f-190">Use all other members of the Office JavaScript API that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.</span></span>

### <a name="cant-do"></a><span data-ttu-id="e692f-191">Vous ne pouvez pas :</span><span class="sxs-lookup"><span data-stu-id="e692f-191">Can't do</span></span>

- <span data-ttu-id="e692f-192">Utilisez le jeton fourni par **mailbox.getCallbackTokenAsync** pour les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="e692f-192">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="e692f-193">Mettre à jour ou supprimer l’élément actuel à l’aide de l’API REST Outlook ou accéder à tous les autres éléments de la boîte aux lettres de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="e692f-193">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="e692f-194">Récupérer l’élément d’événement de calendrier actuel à l’aide de l’API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="e692f-194">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="e692f-195">Utiliser **mailbox.makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="e692f-195">Use **mailbox.makeEWSRequestAsync**.</span></span>

## <a name="readwritemailbox-permission"></a><span data-ttu-id="e692f-196">Autorisation ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e692f-196">ReadWriteMailbox permission</span></span>

<span data-ttu-id="e692f-p108">L’autorisation **ReadWriteMailbox** correspond au plus haut niveau d’autorisation. Indiquez **ReadWriteMailbox** dans l’élément **Permissions** du manifeste pour demander cette autorisation.</span><span class="sxs-lookup"><span data-stu-id="e692f-p108">The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.</span></span>

<span data-ttu-id="e692f-199">En plus des actions prises en charge par l’autorisation **ReadWriteItem**, le jeton fourni par **mailbox.getCallbackTokenAsync** fournit un accès aux opérations des services web Exchange ou à l’API REST Outlook pour effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="e692f-199">In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:</span></span>

- <span data-ttu-id="e692f-200">Lire et écrire toutes les propriétés d’un élément de la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e692f-200">Read and write all properties of any item in the user's mailbox.</span></span>
- <span data-ttu-id="e692f-201">Créer, lire et écrire dans tous les dossiers ou tous les éléments de cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e692f-201">Create, read, and write to any folder or item in that mailbox.</span></span>
- <span data-ttu-id="e692f-202">Envoyer un élément depuis cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e692f-202">Send an item from that mailbox</span></span>

<span data-ttu-id="e692f-203">Grâce à **mailbox.makeEWSRequestAsync**, vous pouvez accéder aux opérations des services web Exchange suivantes :</span><span class="sxs-lookup"><span data-stu-id="e692f-203">Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:</span></span>

- [<span data-ttu-id="e692f-204">CopyItem</span><span class="sxs-lookup"><span data-stu-id="e692f-204">CopyItem</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)
- [<span data-ttu-id="e692f-205">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="e692f-205">CreateFolder</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)
- [<span data-ttu-id="e692f-206">CreateItem</span><span class="sxs-lookup"><span data-stu-id="e692f-206">CreateItem</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)
- [<span data-ttu-id="e692f-207">FindConversation</span><span class="sxs-lookup"><span data-stu-id="e692f-207">FindConversation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)
- [<span data-ttu-id="e692f-208">FindFolder</span><span class="sxs-lookup"><span data-stu-id="e692f-208">FindFolder</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)
- [<span data-ttu-id="e692f-209">FindItem</span><span class="sxs-lookup"><span data-stu-id="e692f-209">FindItem</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)
- [<span data-ttu-id="e692f-210">GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="e692f-210">GetConversationItems</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [<span data-ttu-id="e692f-211">GetFolder</span><span class="sxs-lookup"><span data-stu-id="e692f-211">GetFolder</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)
- [<span data-ttu-id="e692f-212">GetItem</span><span class="sxs-lookup"><span data-stu-id="e692f-212">GetItem</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)
- [<span data-ttu-id="e692f-213">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="e692f-213">MarkAsJunk</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [<span data-ttu-id="e692f-214">MoveItem</span><span class="sxs-lookup"><span data-stu-id="e692f-214">MoveItem</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)
- [<span data-ttu-id="e692f-215">SendItem</span><span class="sxs-lookup"><span data-stu-id="e692f-215">SendItem</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)
- [<span data-ttu-id="e692f-216">UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="e692f-216">UpdateFolder</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [<span data-ttu-id="e692f-217">UpdateItem</span><span class="sxs-lookup"><span data-stu-id="e692f-217">UpdateItem</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)

<span data-ttu-id="e692f-218">Toute tentative d’utilisation d’une opération non prise en charge entraînera une réponse d’erreur.</span><span class="sxs-lookup"><span data-stu-id="e692f-218">Attempting to use an unsupported operation will result in an error response.</span></span>

## <a name="see-also"></a><span data-ttu-id="e692f-219">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e692f-219">See also</span></span>

- [<span data-ttu-id="e692f-220">Confidentialité, autorisations et sécurité pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="e692f-220">Privacy, permissions, and security for Outlook add-ins</span></span>](../develop/privacy-and-security.md)
- [<span data-ttu-id="e692f-221">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="e692f-221">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
