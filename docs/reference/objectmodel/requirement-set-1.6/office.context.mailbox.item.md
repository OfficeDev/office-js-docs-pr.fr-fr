---
title: Office.Context.Mailbox.Item - exigence défini 1.6
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 23f27a2949ddcdaa17ffe3f4711002d47d699458
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387316"
---
# <a name="item"></a><span data-ttu-id="4bd70-102">élément</span><span class="sxs-lookup"><span data-stu-id="4bd70-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4bd70-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4bd70-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4bd70-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4bd70-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-106">Requirements</span></span>

|<span data-ttu-id="4bd70-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-107">Requirement</span></span>| <span data-ttu-id="4bd70-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-110">1.0</span></span>|
|[<span data-ttu-id="4bd70-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4bd70-112">Restricted</span></span>|
|[<span data-ttu-id="4bd70-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4bd70-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4bd70-115">Members and methods</span></span>

| <span data-ttu-id="4bd70-116">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-116">Member</span></span> | <span data-ttu-id="4bd70-117">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4bd70-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4bd70-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="4bd70-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-119">Member</span></span> |
| [<span data-ttu-id="4bd70-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4bd70-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="4bd70-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-121">Member</span></span> |
| [<span data-ttu-id="4bd70-122">body</span><span class="sxs-lookup"><span data-stu-id="4bd70-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="4bd70-123">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-123">Member</span></span> |
| [<span data-ttu-id="4bd70-124">cc</span><span class="sxs-lookup"><span data-stu-id="4bd70-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="4bd70-125">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-125">Member</span></span> |
| [<span data-ttu-id="4bd70-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4bd70-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4bd70-127">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-127">Member</span></span> |
| [<span data-ttu-id="4bd70-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4bd70-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4bd70-129">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-129">Member</span></span> |
| [<span data-ttu-id="4bd70-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4bd70-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4bd70-131">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-131">Member</span></span> |
| [<span data-ttu-id="4bd70-132">end</span><span class="sxs-lookup"><span data-stu-id="4bd70-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="4bd70-133">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-133">Member</span></span> |
| [<span data-ttu-id="4bd70-134">from</span><span class="sxs-lookup"><span data-stu-id="4bd70-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="4bd70-135">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-135">Member</span></span> |
| [<span data-ttu-id="4bd70-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4bd70-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4bd70-137">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-137">Member</span></span> |
| [<span data-ttu-id="4bd70-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4bd70-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4bd70-139">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-139">Member</span></span> |
| [<span data-ttu-id="4bd70-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4bd70-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4bd70-141">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-141">Member</span></span> |
| [<span data-ttu-id="4bd70-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4bd70-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="4bd70-143">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-143">Member</span></span> |
| [<span data-ttu-id="4bd70-144">location</span><span class="sxs-lookup"><span data-stu-id="4bd70-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="4bd70-145">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-145">Member</span></span> |
| [<span data-ttu-id="4bd70-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4bd70-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4bd70-147">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-147">Member</span></span> |
| [<span data-ttu-id="4bd70-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4bd70-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="4bd70-149">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-149">Member</span></span> |
| [<span data-ttu-id="4bd70-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4bd70-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="4bd70-151">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-151">Member</span></span> |
| [<span data-ttu-id="4bd70-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4bd70-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="4bd70-153">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-153">Member</span></span> |
| [<span data-ttu-id="4bd70-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4bd70-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="4bd70-155">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-155">Member</span></span> |
| [<span data-ttu-id="4bd70-156">sender</span><span class="sxs-lookup"><span data-stu-id="4bd70-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="4bd70-157">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-157">Member</span></span> |
| [<span data-ttu-id="4bd70-158">start</span><span class="sxs-lookup"><span data-stu-id="4bd70-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="4bd70-159">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-159">Member</span></span> |
| [<span data-ttu-id="4bd70-160">subject</span><span class="sxs-lookup"><span data-stu-id="4bd70-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="4bd70-161">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-161">Member</span></span> |
| [<span data-ttu-id="4bd70-162">to</span><span class="sxs-lookup"><span data-stu-id="4bd70-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="4bd70-163">Membre</span><span class="sxs-lookup"><span data-stu-id="4bd70-163">Member</span></span> |
| [<span data-ttu-id="4bd70-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4bd70-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-165">Method</span></span> |
| [<span data-ttu-id="4bd70-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4bd70-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-167">Method</span></span> |
| [<span data-ttu-id="4bd70-168">close</span><span class="sxs-lookup"><span data-stu-id="4bd70-168">close</span></span>](#close) | <span data-ttu-id="4bd70-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-169">Method</span></span> |
| [<span data-ttu-id="4bd70-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4bd70-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="4bd70-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-171">Method</span></span> |
| [<span data-ttu-id="4bd70-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4bd70-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="4bd70-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-173">Method</span></span> |
| [<span data-ttu-id="4bd70-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="4bd70-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="4bd70-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-175">Method</span></span> |
| [<span data-ttu-id="4bd70-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4bd70-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="4bd70-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-177">Method</span></span> |
| [<span data-ttu-id="4bd70-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4bd70-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="4bd70-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-179">Method</span></span> |
| [<span data-ttu-id="4bd70-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4bd70-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4bd70-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-181">Method</span></span> |
| [<span data-ttu-id="4bd70-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4bd70-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4bd70-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-183">Method</span></span> |
| [<span data-ttu-id="4bd70-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4bd70-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-185">Method</span></span> |
| [<span data-ttu-id="4bd70-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="4bd70-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="4bd70-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-187">Method</span></span> |
| [<span data-ttu-id="4bd70-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4bd70-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4bd70-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-189">Method</span></span> |
| [<span data-ttu-id="4bd70-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4bd70-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-191">Method</span></span> |
| [<span data-ttu-id="4bd70-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4bd70-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-193">Method</span></span> |
| [<span data-ttu-id="4bd70-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4bd70-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-195">Method</span></span> |
| [<span data-ttu-id="4bd70-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bd70-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4bd70-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="4bd70-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4bd70-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-198">Example</span></span>

<span data-ttu-id="4bd70-199">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bd70-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="4bd70-200">Membres</span><span class="sxs-lookup"><span data-stu-id="4bd70-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="4bd70-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bd70-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="4bd70-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-204">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="4bd70-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4bd70-205">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4bd70-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-206">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-206">Type:</span></span>

*   <span data-ttu-id="4bd70-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bd70-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-208">Requirements</span></span>

|<span data-ttu-id="4bd70-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-209">Requirement</span></span>| <span data-ttu-id="4bd70-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-212">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-212">1.0</span></span>|
|[<span data-ttu-id="4bd70-213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-214">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-216">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-217">Example</span></span>

<span data-ttu-id="4bd70-218">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="4bd70-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="4bd70-220">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4bd70-221">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-222">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-222">Type:</span></span>

*   [<span data-ttu-id="4bd70-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="4bd70-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4bd70-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-224">Requirements</span></span>

|<span data-ttu-id="4bd70-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-225">Requirement</span></span>| <span data-ttu-id="4bd70-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-228">1.1</span><span class="sxs-lookup"><span data-stu-id="4bd70-228">1.1</span></span>|
|[<span data-ttu-id="4bd70-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-230">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-232">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="4bd70-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="4bd70-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="4bd70-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-236">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-236">Type:</span></span>

*   [<span data-ttu-id="4bd70-237">Corps</span><span class="sxs-lookup"><span data-stu-id="4bd70-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="4bd70-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-238">Requirements</span></span>

|<span data-ttu-id="4bd70-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-239">Requirement</span></span>| <span data-ttu-id="4bd70-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-242">1.1</span><span class="sxs-lookup"><span data-stu-id="4bd70-242">1.1</span></span>|
|[<span data-ttu-id="4bd70-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-244">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-246">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-246">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="4bd70-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="4bd70-248">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-248">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4bd70-249">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-249">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-250">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-250">Read mode</span></span>

<span data-ttu-id="4bd70-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-253">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-253">Compose mode</span></span>

<span data-ttu-id="4bd70-254">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-254">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-255">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-255">Type:</span></span>

*   <span data-ttu-id="4bd70-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-257">Requirements</span></span>

|<span data-ttu-id="4bd70-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-258">Requirement</span></span>| <span data-ttu-id="4bd70-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-261">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-261">1.0</span></span>|
|[<span data-ttu-id="4bd70-262">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-263">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-264">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-265">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-265">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-266">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-266">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4bd70-267">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4bd70-267">(nullable) conversationId :String</span></span>

<span data-ttu-id="4bd70-268">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="4bd70-268">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4bd70-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4bd70-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-273">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-273">Type:</span></span>

*   <span data-ttu-id="4bd70-274">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-275">Requirements</span></span>

|<span data-ttu-id="4bd70-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-276">Requirement</span></span>| <span data-ttu-id="4bd70-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-279">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-279">1.0</span></span>|
|[<span data-ttu-id="4bd70-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-281">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-283">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-283">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="4bd70-284">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4bd70-284">dateTimeCreated :Date</span></span>

<span data-ttu-id="4bd70-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-287">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-287">Type:</span></span>

*   <span data-ttu-id="4bd70-288">Date</span><span class="sxs-lookup"><span data-stu-id="4bd70-288">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-289">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-289">Requirements</span></span>

|<span data-ttu-id="4bd70-290">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-290">Requirement</span></span>| <span data-ttu-id="4bd70-291">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-291">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-292">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-293">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-293">1.0</span></span>|
|[<span data-ttu-id="4bd70-294">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-295">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-296">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-297">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-297">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-298">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4bd70-299">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4bd70-299">dateTimeModified :Date</span></span>

<span data-ttu-id="4bd70-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-302">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-302">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-303">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-303">Type:</span></span>

*   <span data-ttu-id="4bd70-304">Date</span><span class="sxs-lookup"><span data-stu-id="4bd70-304">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-305">Requirements</span></span>

|<span data-ttu-id="4bd70-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-306">Requirement</span></span>| <span data-ttu-id="4bd70-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-309">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-309">1.0</span></span>|
|[<span data-ttu-id="4bd70-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-311">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-313">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-314">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-314">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="4bd70-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bd70-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="4bd70-316">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-316">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4bd70-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-319">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-319">Read mode</span></span>

<span data-ttu-id="4bd70-320">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-320">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-321">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-321">Compose mode</span></span>

<span data-ttu-id="4bd70-322">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-322">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4bd70-323">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-323">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-324">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-324">Type:</span></span>

*   <span data-ttu-id="4bd70-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bd70-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-326">Requirements</span></span>

|<span data-ttu-id="4bd70-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-327">Requirement</span></span>| <span data-ttu-id="4bd70-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-330">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-330">1.0</span></span>|
|[<span data-ttu-id="4bd70-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-332">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-334">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-334">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-335">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-335">Example</span></span>

<span data-ttu-id="4bd70-336">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-336">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="4bd70-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4bd70-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="4bd70-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4bd70-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-342">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-342">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-343">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-343">Type:</span></span>

*   [<span data-ttu-id="4bd70-344">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bd70-344">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4bd70-345">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-345">Requirements</span></span>

|<span data-ttu-id="4bd70-346">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-346">Requirement</span></span>| <span data-ttu-id="4bd70-347">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-348">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-349">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-349">1.0</span></span>|
|[<span data-ttu-id="4bd70-350">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-351">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-352">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-353">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-353">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="4bd70-354">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4bd70-354">internetMessageId :String</span></span>

<span data-ttu-id="4bd70-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-357">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-357">Type:</span></span>

*   <span data-ttu-id="4bd70-358">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-358">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-359">Requirements</span></span>

|<span data-ttu-id="4bd70-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-360">Requirement</span></span>| <span data-ttu-id="4bd70-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-363">1.0</span></span>|
|[<span data-ttu-id="4bd70-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-365">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-367">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-368">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4bd70-369">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4bd70-369">itemClass :String</span></span>

<span data-ttu-id="4bd70-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4bd70-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4bd70-374">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-374">Type</span></span> | <span data-ttu-id="4bd70-375">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-375">Description</span></span> | <span data-ttu-id="4bd70-376">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="4bd70-376">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4bd70-377">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="4bd70-377">Appointment items</span></span> | <span data-ttu-id="4bd70-378">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-378">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="4bd70-379">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="4bd70-379">Message items</span></span> | <span data-ttu-id="4bd70-380">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="4bd70-380">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4bd70-381">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-381">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-382">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-382">Type:</span></span>

*   <span data-ttu-id="4bd70-383">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-383">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-384">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-384">Requirements</span></span>

|<span data-ttu-id="4bd70-385">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-385">Requirement</span></span>| <span data-ttu-id="4bd70-386">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-386">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-387">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-387">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-388">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-388">1.0</span></span>|
|[<span data-ttu-id="4bd70-389">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-389">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-390">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-390">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-391">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-391">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-392">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-392">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-393">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-393">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4bd70-394">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4bd70-394">(nullable) itemId :String</span></span>

<span data-ttu-id="4bd70-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-397">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4bd70-397">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4bd70-398">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bd70-398">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4bd70-399">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4bd70-399">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4bd70-400">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4bd70-400">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4bd70-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-403">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-403">Type:</span></span>

*   <span data-ttu-id="4bd70-404">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-404">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-405">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-405">Requirements</span></span>

|<span data-ttu-id="4bd70-406">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-406">Requirement</span></span>| <span data-ttu-id="4bd70-407">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-408">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-409">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-409">1.0</span></span>|
|[<span data-ttu-id="4bd70-410">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-411">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-412">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-413">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-414">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-414">Example</span></span>

<span data-ttu-id="4bd70-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="4bd70-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4bd70-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4bd70-418">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="4bd70-418">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4bd70-419">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-419">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-420">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-420">Type:</span></span>

*   [<span data-ttu-id="4bd70-421">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4bd70-421">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4bd70-422">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-422">Requirements</span></span>

|<span data-ttu-id="4bd70-423">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-423">Requirement</span></span>| <span data-ttu-id="4bd70-424">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-425">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-426">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-426">1.0</span></span>|
|[<span data-ttu-id="4bd70-427">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-428">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-429">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-430">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-430">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-431">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-431">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="4bd70-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="4bd70-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="4bd70-433">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-433">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-434">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-434">Read mode</span></span>

<span data-ttu-id="4bd70-435">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-435">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-436">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-436">Compose mode</span></span>

<span data-ttu-id="4bd70-437">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-437">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-438">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-438">Type:</span></span>

*   <span data-ttu-id="4bd70-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="4bd70-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-440">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-440">Requirements</span></span>

|<span data-ttu-id="4bd70-441">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-441">Requirement</span></span>| <span data-ttu-id="4bd70-442">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-443">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-444">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-444">1.0</span></span>|
|[<span data-ttu-id="4bd70-445">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-446">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-447">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-448">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-449">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-449">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4bd70-450">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4bd70-450">normalizedSubject :String</span></span>

<span data-ttu-id="4bd70-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4bd70-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="4bd70-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-455">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-455">Type:</span></span>

*   <span data-ttu-id="4bd70-456">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-457">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-457">Requirements</span></span>

|<span data-ttu-id="4bd70-458">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-458">Requirement</span></span>| <span data-ttu-id="4bd70-459">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-460">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-461">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-461">1.0</span></span>|
|[<span data-ttu-id="4bd70-462">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-463">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-464">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-465">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-466">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="4bd70-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4bd70-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="4bd70-468">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-468">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-469">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-469">Type:</span></span>

*   [<span data-ttu-id="4bd70-470">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4bd70-470">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4bd70-471">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-471">Requirements</span></span>

|<span data-ttu-id="4bd70-472">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-472">Requirement</span></span>| <span data-ttu-id="4bd70-473">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-474">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-475">1.3</span><span class="sxs-lookup"><span data-stu-id="4bd70-475">1.3</span></span>|
|[<span data-ttu-id="4bd70-476">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-477">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-478">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-479">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-479">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="4bd70-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="4bd70-481">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-481">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4bd70-482">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-482">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-483">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-483">Read mode</span></span>

<span data-ttu-id="4bd70-484">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4bd70-484">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-485">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-485">Compose mode</span></span>

<span data-ttu-id="4bd70-486">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="4bd70-486">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-487">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-487">Type:</span></span>

*   <span data-ttu-id="4bd70-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-489">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-489">Requirements</span></span>

|<span data-ttu-id="4bd70-490">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-490">Requirement</span></span>| <span data-ttu-id="4bd70-491">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-492">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-493">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-493">1.0</span></span>|
|[<span data-ttu-id="4bd70-494">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-495">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-496">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-497">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-497">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-498">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-498">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="4bd70-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4bd70-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="4bd70-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-502">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-502">Type:</span></span>

*   [<span data-ttu-id="4bd70-503">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bd70-503">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4bd70-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-504">Requirements</span></span>

|<span data-ttu-id="4bd70-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-505">Requirement</span></span>| <span data-ttu-id="4bd70-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-508">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-508">1.0</span></span>|
|[<span data-ttu-id="4bd70-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-510">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-512">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-512">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-513">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="4bd70-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="4bd70-515">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-515">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4bd70-516">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-516">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-517">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-517">Read mode</span></span>

<span data-ttu-id="4bd70-518">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4bd70-518">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-519">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-519">Compose mode</span></span>

<span data-ttu-id="4bd70-520">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="4bd70-520">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-521">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-521">Type:</span></span>

*   <span data-ttu-id="4bd70-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-523">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-523">Requirements</span></span>

|<span data-ttu-id="4bd70-524">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-524">Requirement</span></span>| <span data-ttu-id="4bd70-525">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-526">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-527">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-527">1.0</span></span>|
|[<span data-ttu-id="4bd70-528">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-529">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-530">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-530">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-531">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-531">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-532">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-532">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="4bd70-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4bd70-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="4bd70-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4bd70-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-538">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-538">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-539">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-539">Type:</span></span>

*   [<span data-ttu-id="4bd70-540">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bd70-540">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4bd70-541">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-541">Requirements</span></span>

|<span data-ttu-id="4bd70-542">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-542">Requirement</span></span>| <span data-ttu-id="4bd70-543">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-543">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-544">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-545">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-545">1.0</span></span>|
|[<span data-ttu-id="4bd70-546">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-546">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-547">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-547">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-548">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-548">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-549">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-549">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-550">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-550">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="4bd70-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bd70-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="4bd70-552">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-552">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4bd70-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-555">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-555">Read mode</span></span>

<span data-ttu-id="4bd70-556">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-556">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-557">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-557">Compose mode</span></span>

<span data-ttu-id="4bd70-558">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-558">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4bd70-559">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-559">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-560">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-560">Type:</span></span>

*   <span data-ttu-id="4bd70-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bd70-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-562">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-562">Requirements</span></span>

|<span data-ttu-id="4bd70-563">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-563">Requirement</span></span>| <span data-ttu-id="4bd70-564">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-565">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-566">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-566">1.0</span></span>|
|[<span data-ttu-id="4bd70-567">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-568">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-569">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-570">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-570">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-571">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-571">Example</span></span>

<span data-ttu-id="4bd70-572">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-572">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="4bd70-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4bd70-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="4bd70-574">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4bd70-575">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4bd70-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-576">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-576">Read mode</span></span>

<span data-ttu-id="4bd70-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-579">Compose mode</span></span>

<span data-ttu-id="4bd70-580">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="4bd70-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bd70-581">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-581">Type:</span></span>

*   <span data-ttu-id="4bd70-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4bd70-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-583">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-583">Requirements</span></span>

|<span data-ttu-id="4bd70-584">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-584">Requirement</span></span>| <span data-ttu-id="4bd70-585">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-586">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-587">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-587">1.0</span></span>|
|[<span data-ttu-id="4bd70-588">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-589">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-590">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-591">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-591">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="4bd70-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="4bd70-593">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4bd70-594">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bd70-595">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-595">Read mode</span></span>

<span data-ttu-id="4bd70-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bd70-598">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-598">Compose mode</span></span>

<span data-ttu-id="4bd70-599">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4bd70-600">Type :</span><span class="sxs-lookup"><span data-stu-id="4bd70-600">Type:</span></span>

*   <span data-ttu-id="4bd70-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bd70-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-602">Requirements</span></span>

|<span data-ttu-id="4bd70-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-603">Requirement</span></span>| <span data-ttu-id="4bd70-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-606">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-606">1.0</span></span>|
|[<span data-ttu-id="4bd70-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-608">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-610">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-611">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-611">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="4bd70-612">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4bd70-612">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4bd70-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bd70-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4bd70-614">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="4bd70-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4bd70-615">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="4bd70-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4bd70-616">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4bd70-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-617">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-617">Parameters:</span></span>

|<span data-ttu-id="4bd70-618">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-618">Name</span></span>| <span data-ttu-id="4bd70-619">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-619">Type</span></span>| <span data-ttu-id="4bd70-620">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-620">Attributes</span></span>| <span data-ttu-id="4bd70-621">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4bd70-622">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-622">String</span></span>||<span data-ttu-id="4bd70-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4bd70-625">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-625">String</span></span>||<span data-ttu-id="4bd70-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4bd70-628">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-628">Object</span></span>| <span data-ttu-id="4bd70-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-629">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-630">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-630">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="4bd70-631">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-631">Object</span></span> | <span data-ttu-id="4bd70-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-632">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-633">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="4bd70-634">Boolean</span><span class="sxs-lookup"><span data-stu-id="4bd70-634">Boolean</span></span> | <span data-ttu-id="4bd70-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-635">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-636">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-636">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="4bd70-637">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-637">function</span></span>| <span data-ttu-id="4bd70-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-638">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-639">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bd70-640">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-640">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4bd70-641">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-641">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bd70-642">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4bd70-642">Errors</span></span>

| <span data-ttu-id="4bd70-643">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4bd70-643">Error code</span></span> | <span data-ttu-id="4bd70-644">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-644">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4bd70-645">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-645">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4bd70-646">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-646">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4bd70-647">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-647">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-648">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-648">Requirements</span></span>

|<span data-ttu-id="4bd70-649">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-649">Requirement</span></span>| <span data-ttu-id="4bd70-650">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-651">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-652">1.1</span><span class="sxs-lookup"><span data-stu-id="4bd70-652">1.1</span></span>|
|[<span data-ttu-id="4bd70-653">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-653">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-654">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-654">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-655">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-655">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-656">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-656">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bd70-657">Exemples</span><span class="sxs-lookup"><span data-stu-id="4bd70-657">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="4bd70-658">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-658">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4bd70-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bd70-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4bd70-660">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-660">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4bd70-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4bd70-664">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4bd70-664">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4bd70-665">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-665">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-666">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-666">Parameters:</span></span>

|<span data-ttu-id="4bd70-667">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-667">Name</span></span>| <span data-ttu-id="4bd70-668">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-668">Type</span></span>| <span data-ttu-id="4bd70-669">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-669">Attributes</span></span>| <span data-ttu-id="4bd70-670">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-670">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4bd70-671">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-671">String</span></span>||<span data-ttu-id="4bd70-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4bd70-674">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-674">String</span></span>||<span data-ttu-id="4bd70-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4bd70-677">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-677">Object</span></span>| <span data-ttu-id="4bd70-678">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-678">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-679">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-679">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4bd70-680">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-680">Object</span></span>| <span data-ttu-id="4bd70-681">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-681">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-682">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-682">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4bd70-683">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-683">function</span></span>| <span data-ttu-id="4bd70-684">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-684">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-685">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-685">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bd70-686">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-686">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4bd70-687">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-687">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bd70-688">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4bd70-688">Errors</span></span>

| <span data-ttu-id="4bd70-689">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4bd70-689">Error code</span></span> | <span data-ttu-id="4bd70-690">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-690">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4bd70-691">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-691">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-692">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-692">Requirements</span></span>

|<span data-ttu-id="4bd70-693">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-693">Requirement</span></span>| <span data-ttu-id="4bd70-694">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-695">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-696">1.1</span><span class="sxs-lookup"><span data-stu-id="4bd70-696">1.1</span></span>|
|[<span data-ttu-id="4bd70-697">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-697">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-698">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-698">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-699">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-699">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-700">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-700">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-701">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-701">Example</span></span>

<span data-ttu-id="4bd70-702">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-702">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="4bd70-703">close()</span><span class="sxs-lookup"><span data-stu-id="4bd70-703">close()</span></span>

<span data-ttu-id="4bd70-704">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="4bd70-704">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4bd70-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-707">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-707">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4bd70-708">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="4bd70-708">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-709">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-709">Requirements</span></span>

|<span data-ttu-id="4bd70-710">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-710">Requirement</span></span>| <span data-ttu-id="4bd70-711">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-712">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-713">1.3</span><span class="sxs-lookup"><span data-stu-id="4bd70-713">1.3</span></span>|
|[<span data-ttu-id="4bd70-714">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-714">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-715">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4bd70-715">Restricted</span></span>|
|[<span data-ttu-id="4bd70-716">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-716">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-717">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-717">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="4bd70-718">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4bd70-718">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="4bd70-719">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4bd70-719">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-720">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-720">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-721">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-721">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4bd70-722">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4bd70-722">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4bd70-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-726">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-726">Parameters:</span></span>

| <span data-ttu-id="4bd70-727">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-727">Name</span></span> | <span data-ttu-id="4bd70-728">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-728">Type</span></span> | <span data-ttu-id="4bd70-729">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-729">Attributes</span></span> | <span data-ttu-id="4bd70-730">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-730">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4bd70-731">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4bd70-731">String &#124; Object</span></span>| |<span data-ttu-id="4bd70-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4bd70-734">**OU**</span><span class="sxs-lookup"><span data-stu-id="4bd70-734">**OR**</span></span><br/><span data-ttu-id="4bd70-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4bd70-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4bd70-737">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-737">String</span></span> | <span data-ttu-id="4bd70-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-738">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4bd70-741">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-741">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4bd70-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-742">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-743">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-743">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4bd70-744">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-744">String</span></span> | | <span data-ttu-id="4bd70-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4bd70-747">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-747">String</span></span> | | <span data-ttu-id="4bd70-748">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-748">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4bd70-749">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-749">String</span></span> | | <span data-ttu-id="4bd70-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4bd70-752">Booléen</span><span class="sxs-lookup"><span data-stu-id="4bd70-752">Boolean</span></span> | | <span data-ttu-id="4bd70-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4bd70-755">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-755">String</span></span> | | <span data-ttu-id="4bd70-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4bd70-759">function</span><span class="sxs-lookup"><span data-stu-id="4bd70-759">function</span></span> | <span data-ttu-id="4bd70-760">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-760">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-761">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-761">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-762">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-762">Requirements</span></span>

|<span data-ttu-id="4bd70-763">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-763">Requirement</span></span>| <span data-ttu-id="4bd70-764">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-765">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-766">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-766">1.0</span></span>|
|[<span data-ttu-id="4bd70-767">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-767">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-768">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-769">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-769">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-770">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-770">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bd70-771">Exemples</span><span class="sxs-lookup"><span data-stu-id="4bd70-771">Examples</span></span>

<span data-ttu-id="4bd70-772">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-772">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4bd70-773">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4bd70-773">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4bd70-774">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4bd70-774">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4bd70-775">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4bd70-775">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4bd70-776">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-776">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4bd70-777">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-777">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="4bd70-778">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4bd70-778">displayReplyForm(formData)</span></span>

<span data-ttu-id="4bd70-779">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4bd70-779">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-780">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-780">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-781">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-781">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4bd70-782">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4bd70-782">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4bd70-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-786">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-786">Parameters:</span></span>

| <span data-ttu-id="4bd70-787">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-787">Name</span></span> | <span data-ttu-id="4bd70-788">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-788">Type</span></span> | <span data-ttu-id="4bd70-789">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-789">Attributes</span></span> | <span data-ttu-id="4bd70-790">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-790">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4bd70-791">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4bd70-791">String &#124; Object</span></span>| | <span data-ttu-id="4bd70-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4bd70-794">**OU**</span><span class="sxs-lookup"><span data-stu-id="4bd70-794">**OR**</span></span><br/><span data-ttu-id="4bd70-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4bd70-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4bd70-797">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-797">String</span></span> | <span data-ttu-id="4bd70-798">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-798">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4bd70-801">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-801">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4bd70-802">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-802">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-803">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-803">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4bd70-804">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-804">String</span></span> | | <span data-ttu-id="4bd70-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4bd70-807">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-807">String</span></span> | | <span data-ttu-id="4bd70-808">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-808">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4bd70-809">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-809">String</span></span> | | <span data-ttu-id="4bd70-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4bd70-812">Booléen</span><span class="sxs-lookup"><span data-stu-id="4bd70-812">Boolean</span></span> | | <span data-ttu-id="4bd70-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4bd70-815">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-815">String</span></span> | | <span data-ttu-id="4bd70-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4bd70-819">function</span><span class="sxs-lookup"><span data-stu-id="4bd70-819">function</span></span> | <span data-ttu-id="4bd70-820">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-820">&lt;optional&gt;</span></span> | <span data-ttu-id="4bd70-821">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-822">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-822">Requirements</span></span>

|<span data-ttu-id="4bd70-823">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-823">Requirement</span></span>| <span data-ttu-id="4bd70-824">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-825">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-826">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-826">1.0</span></span>|
|[<span data-ttu-id="4bd70-827">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-827">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-828">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-828">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-829">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-829">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-830">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-830">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bd70-831">Exemples</span><span class="sxs-lookup"><span data-stu-id="4bd70-831">Examples</span></span>

<span data-ttu-id="4bd70-832">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-832">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4bd70-833">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4bd70-833">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4bd70-834">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4bd70-834">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4bd70-835">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4bd70-835">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4bd70-836">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-836">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4bd70-837">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-837">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="4bd70-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4bd70-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="4bd70-839">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4bd70-839">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-840">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-840">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-841">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-841">Requirements</span></span>

|<span data-ttu-id="4bd70-842">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-842">Requirement</span></span>| <span data-ttu-id="4bd70-843">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-843">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-844">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-844">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-845">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-845">1.0</span></span>|
|[<span data-ttu-id="4bd70-846">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-846">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-847">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-847">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-848">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-848">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-849">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-849">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-850">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-850">Returns:</span></span>

<span data-ttu-id="4bd70-851">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4bd70-851">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4bd70-852">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-852">Example</span></span>

<span data-ttu-id="4bd70-853">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-853">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="4bd70-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4bd70-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4bd70-855">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4bd70-855">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-856">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-856">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-857">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-857">Parameters:</span></span>

|<span data-ttu-id="4bd70-858">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-858">Name</span></span>| <span data-ttu-id="4bd70-859">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-859">Type</span></span>| <span data-ttu-id="4bd70-860">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-860">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4bd70-861">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4bd70-861">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="4bd70-862">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="4bd70-862">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-863">Requirements</span></span>

|<span data-ttu-id="4bd70-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-864">Requirement</span></span>| <span data-ttu-id="4bd70-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-867">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-867">1.0</span></span>|
|[<span data-ttu-id="4bd70-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-869">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4bd70-869">Restricted</span></span>|
|[<span data-ttu-id="4bd70-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-872">Returns:</span></span>

<span data-ttu-id="4bd70-873">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="4bd70-873">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4bd70-874">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4bd70-874">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4bd70-875">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-875">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4bd70-876">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="4bd70-876">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4bd70-877">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4bd70-877">Value of `entityType`</span></span> | <span data-ttu-id="4bd70-878">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="4bd70-878">Type of objects in returned array</span></span> | <span data-ttu-id="4bd70-879">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="4bd70-879">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4bd70-880">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-880">String</span></span> | <span data-ttu-id="4bd70-881">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bd70-881">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4bd70-882">Contact</span><span class="sxs-lookup"><span data-stu-id="4bd70-882">Contact</span></span> | <span data-ttu-id="4bd70-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bd70-883">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4bd70-884">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-884">String</span></span> | <span data-ttu-id="4bd70-885">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bd70-885">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4bd70-886">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4bd70-886">MeetingSuggestion</span></span> | <span data-ttu-id="4bd70-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bd70-887">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4bd70-888">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4bd70-888">PhoneNumber</span></span> | <span data-ttu-id="4bd70-889">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bd70-889">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4bd70-890">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4bd70-890">TaskSuggestion</span></span> | <span data-ttu-id="4bd70-891">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bd70-891">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4bd70-892">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-892">String</span></span> | <span data-ttu-id="4bd70-893">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bd70-893">**Restricted**</span></span> |

<span data-ttu-id="4bd70-894">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4bd70-894">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4bd70-895">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-895">Example</span></span>

<span data-ttu-id="4bd70-896">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4bd70-896">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="4bd70-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4bd70-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4bd70-898">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-898">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-899">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-899">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-900">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4bd70-900">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-901">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-901">Parameters:</span></span>

|<span data-ttu-id="4bd70-902">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-902">Name</span></span>| <span data-ttu-id="4bd70-903">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-903">Type</span></span>| <span data-ttu-id="4bd70-904">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-904">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4bd70-905">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-905">String</span></span>|<span data-ttu-id="4bd70-906">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4bd70-906">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-907">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-907">Requirements</span></span>

|<span data-ttu-id="4bd70-908">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-908">Requirement</span></span>| <span data-ttu-id="4bd70-909">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-909">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-910">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-910">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-911">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-911">1.0</span></span>|
|[<span data-ttu-id="4bd70-912">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-912">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-913">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-913">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-914">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-914">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-915">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-915">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-916">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-916">Returns:</span></span>

<span data-ttu-id="4bd70-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4bd70-919">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4bd70-919">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="4bd70-920">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4bd70-920">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4bd70-921">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-921">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-922">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-922">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4bd70-926">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4bd70-926">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4bd70-927">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-927">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4bd70-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-931">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-931">Requirements</span></span>

|<span data-ttu-id="4bd70-932">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-932">Requirement</span></span>| <span data-ttu-id="4bd70-933">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-934">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-935">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-935">1.0</span></span>|
|[<span data-ttu-id="4bd70-936">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-936">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-937">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-938">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-938">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-939">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-939">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-940">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-940">Returns:</span></span>

<span data-ttu-id="4bd70-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4bd70-943">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4bd70-943">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4bd70-944">Object</span><span class="sxs-lookup"><span data-stu-id="4bd70-944">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4bd70-945">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-945">Example</span></span>

<span data-ttu-id="4bd70-946">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4bd70-946">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4bd70-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4bd70-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4bd70-948">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-948">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-949">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-949">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-950">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4bd70-950">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4bd70-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-953">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-953">Parameters:</span></span>

|<span data-ttu-id="4bd70-954">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-954">Name</span></span>| <span data-ttu-id="4bd70-955">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-955">Type</span></span>| <span data-ttu-id="4bd70-956">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-956">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4bd70-957">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-957">String</span></span>|<span data-ttu-id="4bd70-958">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4bd70-958">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-959">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-959">Requirements</span></span>

|<span data-ttu-id="4bd70-960">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-960">Requirement</span></span>| <span data-ttu-id="4bd70-961">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-962">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-963">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-963">1.0</span></span>|
|[<span data-ttu-id="4bd70-964">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-965">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-965">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-966">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-967">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-968">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-968">Returns:</span></span>

<span data-ttu-id="4bd70-969">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-969">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4bd70-970">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4bd70-970">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4bd70-971">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4bd70-971">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4bd70-972">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-972">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4bd70-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4bd70-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4bd70-974">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bd70-974">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4bd70-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-977">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-977">Parameters:</span></span>

|<span data-ttu-id="4bd70-978">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-978">Name</span></span>| <span data-ttu-id="4bd70-979">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-979">Type</span></span>| <span data-ttu-id="4bd70-980">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-980">Attributes</span></span>| <span data-ttu-id="4bd70-981">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-981">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4bd70-982">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bd70-982">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4bd70-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4bd70-986">Object</span><span class="sxs-lookup"><span data-stu-id="4bd70-986">Object</span></span>| <span data-ttu-id="4bd70-987">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-987">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-988">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-988">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4bd70-989">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-989">Object</span></span>| <span data-ttu-id="4bd70-990">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-990">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-991">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-991">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4bd70-992">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-992">function</span></span>||<span data-ttu-id="4bd70-993">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-993">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bd70-994">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-994">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4bd70-995">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-995">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-996">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-996">Requirements</span></span>

|<span data-ttu-id="4bd70-997">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-997">Requirement</span></span>| <span data-ttu-id="4bd70-998">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-998">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-999">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-999">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1000">1.2</span><span class="sxs-lookup"><span data-stu-id="4bd70-1000">1.2</span></span>|
|[<span data-ttu-id="4bd70-1001">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1001">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1002">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1002">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-1003">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1003">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1004">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-1004">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-1005">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1005">Returns:</span></span>

<span data-ttu-id="4bd70-1006">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1006">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4bd70-1007">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4bd70-1007">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4bd70-1008">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bd70-1008">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4bd70-1009">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1009">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="4bd70-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4bd70-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="4bd70-p163">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4bd70-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-1013">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-1014">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1014">Requirements</span></span>

|<span data-ttu-id="4bd70-1015">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1015">Requirement</span></span>| <span data-ttu-id="4bd70-1016">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1017">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1018">1.6</span><span class="sxs-lookup"><span data-stu-id="4bd70-1018">1.6</span></span> |
|[<span data-ttu-id="4bd70-1019">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1019">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1020">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1020">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-1021">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1021">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1022">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-1022">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-1023">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1023">Returns:</span></span>

<span data-ttu-id="4bd70-1024">Type : [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4bd70-1024">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4bd70-1025">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1025">Example</span></span>

<span data-ttu-id="4bd70-1026">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1026">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4bd70-1027">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4bd70-1027">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4bd70-p164">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4bd70-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-1030">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1030">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4bd70-p165">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4bd70-1034">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1034">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4bd70-1035">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1035">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4bd70-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bd70-1039">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1039">Requirements</span></span>

|<span data-ttu-id="4bd70-1040">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1040">Requirement</span></span>| <span data-ttu-id="4bd70-1041">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1042">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1043">1.6</span><span class="sxs-lookup"><span data-stu-id="4bd70-1043">1.6</span></span> |
|[<span data-ttu-id="4bd70-1044">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1044">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1045">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1045">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-1046">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1046">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1047">Lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-1047">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bd70-1048">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1048">Returns:</span></span>

<span data-ttu-id="4bd70-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4bd70-1051">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1051">Example</span></span>

<span data-ttu-id="4bd70-1052">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1052">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4bd70-1053">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4bd70-1053">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4bd70-1054">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1054">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4bd70-p168">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-1058">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1058">Parameters:</span></span>

|<span data-ttu-id="4bd70-1059">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-1059">Name</span></span>| <span data-ttu-id="4bd70-1060">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-1060">Type</span></span>| <span data-ttu-id="4bd70-1061">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-1061">Attributes</span></span>| <span data-ttu-id="4bd70-1062">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-1062">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4bd70-1063">function</span><span class="sxs-lookup"><span data-stu-id="4bd70-1063">function</span></span>||<span data-ttu-id="4bd70-1064">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-1064">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bd70-1065">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1065">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4bd70-1066">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1066">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4bd70-1067">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1067">Object</span></span>| <span data-ttu-id="4bd70-1068">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1069">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1069">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4bd70-1070">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1070">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-1071">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1071">Requirements</span></span>

|<span data-ttu-id="4bd70-1072">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1072">Requirement</span></span>| <span data-ttu-id="4bd70-1073">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1074">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1075">1.0</span><span class="sxs-lookup"><span data-stu-id="4bd70-1075">1.0</span></span>|
|[<span data-ttu-id="4bd70-1076">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1077">ReadItem</span></span>|
|[<span data-ttu-id="4bd70-1078">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1079">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bd70-1079">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-1080">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1080">Example</span></span>

<span data-ttu-id="4bd70-p171">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4bd70-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bd70-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4bd70-1085">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1085">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4bd70-p172">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-1090">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1090">Parameters:</span></span>

|<span data-ttu-id="4bd70-1091">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-1091">Name</span></span>| <span data-ttu-id="4bd70-1092">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-1092">Type</span></span>| <span data-ttu-id="4bd70-1093">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-1093">Attributes</span></span>| <span data-ttu-id="4bd70-1094">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-1094">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4bd70-1095">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-1095">String</span></span>||<span data-ttu-id="4bd70-1096">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1096">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4bd70-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1097">Object</span></span>| <span data-ttu-id="4bd70-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1099">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4bd70-1100">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1100">Object</span></span>| <span data-ttu-id="4bd70-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1102">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4bd70-1103">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-1103">function</span></span>| <span data-ttu-id="4bd70-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1105">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bd70-1106">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bd70-1107">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4bd70-1107">Errors</span></span>

| <span data-ttu-id="4bd70-1108">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1108">Error code</span></span> | <span data-ttu-id="4bd70-1109">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4bd70-1110">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-1111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1111">Requirements</span></span>

|<span data-ttu-id="4bd70-1112">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1112">Requirement</span></span>| <span data-ttu-id="4bd70-1113">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1114">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="4bd70-1115">1.1</span></span>|
|[<span data-ttu-id="4bd70-1116">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-1118">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1119">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-1120">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1120">Example</span></span>

<span data-ttu-id="4bd70-1121">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="4bd70-1121">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="4bd70-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4bd70-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="4bd70-1123">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="4bd70-p173">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-1127">si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4bd70-1128">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4bd70-p175">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4bd70-1132">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4bd70-1133">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="4bd70-1134">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="4bd70-1135">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-1136">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1136">Parameters:</span></span>

|<span data-ttu-id="4bd70-1137">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-1137">Name</span></span>| <span data-ttu-id="4bd70-1138">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-1138">Type</span></span>| <span data-ttu-id="4bd70-1139">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-1139">Attributes</span></span>| <span data-ttu-id="4bd70-1140">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="4bd70-1141">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1141">Object</span></span>| <span data-ttu-id="4bd70-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1143">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4bd70-1144">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1144">Object</span></span>| <span data-ttu-id="4bd70-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1146">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4bd70-1147">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-1147">function</span></span>||<span data-ttu-id="4bd70-1148">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bd70-1149">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bd70-1150">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1150">Requirements</span></span>

|<span data-ttu-id="4bd70-1151">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1151">Requirement</span></span>| <span data-ttu-id="4bd70-1152">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1153">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="4bd70-1154">1.3</span></span>|
|[<span data-ttu-id="4bd70-1155">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-1157">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1158">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bd70-1159">範例</span><span class="sxs-lookup"><span data-stu-id="4bd70-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="4bd70-p177">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4bd70-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4bd70-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4bd70-1163">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4bd70-p178">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bd70-1167">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4bd70-1167">Parameters:</span></span>

|<span data-ttu-id="4bd70-1168">Nom</span><span class="sxs-lookup"><span data-stu-id="4bd70-1168">Name</span></span>| <span data-ttu-id="4bd70-1169">Type</span><span class="sxs-lookup"><span data-stu-id="4bd70-1169">Type</span></span>| <span data-ttu-id="4bd70-1170">Attributs</span><span class="sxs-lookup"><span data-stu-id="4bd70-1170">Attributes</span></span>| <span data-ttu-id="4bd70-1171">Description</span><span class="sxs-lookup"><span data-stu-id="4bd70-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4bd70-1172">String</span><span class="sxs-lookup"><span data-stu-id="4bd70-1172">String</span></span>||<span data-ttu-id="4bd70-p179">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4bd70-1176">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1176">Object</span></span>| <span data-ttu-id="4bd70-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1178">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4bd70-1179">Objet</span><span class="sxs-lookup"><span data-stu-id="4bd70-1179">Object</span></span>| <span data-ttu-id="4bd70-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-1181">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="4bd70-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bd70-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="4bd70-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4bd70-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="4bd70-p180">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4bd70-p181">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="4bd70-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4bd70-1188">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="4bd70-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4bd70-1189">fonction</span><span class="sxs-lookup"><span data-stu-id="4bd70-1189">function</span></span>||<span data-ttu-id="4bd70-1190">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bd70-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bd70-1191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4bd70-1191">Requirements</span></span>

|<span data-ttu-id="4bd70-1192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bd70-1192">Requirement</span></span>| <span data-ttu-id="4bd70-1193">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bd70-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bd70-1194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bd70-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bd70-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="4bd70-1195">1.2</span></span>|
|[<span data-ttu-id="4bd70-1196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4bd70-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bd70-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bd70-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bd70-1198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bd70-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bd70-1199">Composition</span><span class="sxs-lookup"><span data-stu-id="4bd70-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bd70-1200">Exemple</span><span class="sxs-lookup"><span data-stu-id="4bd70-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
