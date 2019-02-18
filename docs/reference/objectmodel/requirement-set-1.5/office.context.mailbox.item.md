---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Priority
ms.openlocfilehash: b95985f7ed76b9952e5698e9190ff4c1fa00a7cb
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068237"
---
# <a name="item"></a><span data-ttu-id="b9d54-102">élément</span><span class="sxs-lookup"><span data-stu-id="b9d54-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b9d54-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b9d54-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b9d54-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b9d54-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-106">Requirements</span></span>

|<span data-ttu-id="b9d54-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-107">Requirement</span></span>| <span data-ttu-id="b9d54-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-110">1.0</span></span>|
|[<span data-ttu-id="b9d54-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b9d54-112">Restricted</span></span>|
|[<span data-ttu-id="b9d54-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b9d54-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="b9d54-115">Members and methods</span></span>

| <span data-ttu-id="b9d54-116">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-116">Member</span></span> | <span data-ttu-id="b9d54-117">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b9d54-118">attachments</span><span class="sxs-lookup"><span data-stu-id="b9d54-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="b9d54-119">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-119">Member</span></span> |
| [<span data-ttu-id="b9d54-120">bcc</span><span class="sxs-lookup"><span data-stu-id="b9d54-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="b9d54-121">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-121">Member</span></span> |
| [<span data-ttu-id="b9d54-122">body</span><span class="sxs-lookup"><span data-stu-id="b9d54-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="b9d54-123">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-123">Member</span></span> |
| [<span data-ttu-id="b9d54-124">cc</span><span class="sxs-lookup"><span data-stu-id="b9d54-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="b9d54-125">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-125">Member</span></span> |
| [<span data-ttu-id="b9d54-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="b9d54-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="b9d54-127">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-127">Member</span></span> |
| [<span data-ttu-id="b9d54-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="b9d54-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="b9d54-129">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-129">Member</span></span> |
| [<span data-ttu-id="b9d54-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b9d54-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="b9d54-131">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-131">Member</span></span> |
| [<span data-ttu-id="b9d54-132">end</span><span class="sxs-lookup"><span data-stu-id="b9d54-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="b9d54-133">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-133">Member</span></span> |
| [<span data-ttu-id="b9d54-134">from</span><span class="sxs-lookup"><span data-stu-id="b9d54-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="b9d54-135">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-135">Member</span></span> |
| [<span data-ttu-id="b9d54-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="b9d54-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="b9d54-137">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-137">Member</span></span> |
| [<span data-ttu-id="b9d54-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="b9d54-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="b9d54-139">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-139">Member</span></span> |
| [<span data-ttu-id="b9d54-140">itemId</span><span class="sxs-lookup"><span data-stu-id="b9d54-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="b9d54-141">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-141">Member</span></span> |
| [<span data-ttu-id="b9d54-142">itemType</span><span class="sxs-lookup"><span data-stu-id="b9d54-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="b9d54-143">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-143">Member</span></span> |
| [<span data-ttu-id="b9d54-144">location</span><span class="sxs-lookup"><span data-stu-id="b9d54-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="b9d54-145">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-145">Member</span></span> |
| [<span data-ttu-id="b9d54-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="b9d54-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="b9d54-147">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-147">Member</span></span> |
| [<span data-ttu-id="b9d54-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="b9d54-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="b9d54-149">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-149">Member</span></span> |
| [<span data-ttu-id="b9d54-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b9d54-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="b9d54-151">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-151">Member</span></span> |
| [<span data-ttu-id="b9d54-152">organizer</span><span class="sxs-lookup"><span data-stu-id="b9d54-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="b9d54-153">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-153">Member</span></span> |
| [<span data-ttu-id="b9d54-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b9d54-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="b9d54-155">Member</span><span class="sxs-lookup"><span data-stu-id="b9d54-155">Member</span></span> |
| [<span data-ttu-id="b9d54-156">sender</span><span class="sxs-lookup"><span data-stu-id="b9d54-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="b9d54-157">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-157">Member</span></span> |
| [<span data-ttu-id="b9d54-158">start</span><span class="sxs-lookup"><span data-stu-id="b9d54-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="b9d54-159">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-159">Member</span></span> |
| [<span data-ttu-id="b9d54-160">subject</span><span class="sxs-lookup"><span data-stu-id="b9d54-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="b9d54-161">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-161">Member</span></span> |
| [<span data-ttu-id="b9d54-162">to</span><span class="sxs-lookup"><span data-stu-id="b9d54-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="b9d54-163">Membre</span><span class="sxs-lookup"><span data-stu-id="b9d54-163">Member</span></span> |
| [<span data-ttu-id="b9d54-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="b9d54-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-165">Method</span></span> |
| [<span data-ttu-id="b9d54-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="b9d54-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-167">Method</span></span> |
| [<span data-ttu-id="b9d54-168">close</span><span class="sxs-lookup"><span data-stu-id="b9d54-168">close</span></span>](#close) | <span data-ttu-id="b9d54-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-169">Method</span></span> |
| [<span data-ttu-id="b9d54-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b9d54-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="b9d54-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-171">Method</span></span> |
| [<span data-ttu-id="b9d54-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b9d54-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="b9d54-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-173">Method</span></span> |
| [<span data-ttu-id="b9d54-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="b9d54-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="b9d54-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-175">Method</span></span> |
| [<span data-ttu-id="b9d54-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b9d54-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="b9d54-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-177">Method</span></span> |
| [<span data-ttu-id="b9d54-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b9d54-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="b9d54-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-179">Method</span></span> |
| [<span data-ttu-id="b9d54-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b9d54-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="b9d54-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-181">Method</span></span> |
| [<span data-ttu-id="b9d54-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b9d54-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="b9d54-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-183">Method</span></span> |
| [<span data-ttu-id="b9d54-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="b9d54-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-185">Method</span></span> |
| [<span data-ttu-id="b9d54-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="b9d54-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-187">Method</span></span> |
| [<span data-ttu-id="b9d54-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="b9d54-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-189">Method</span></span> |
| [<span data-ttu-id="b9d54-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="b9d54-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-191">Method</span></span> |
| [<span data-ttu-id="b9d54-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b9d54-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="b9d54-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="b9d54-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="b9d54-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-194">Example</span></span>

<span data-ttu-id="b9d54-195">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9d54-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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
};
```

### <a name="members"></a><span data-ttu-id="b9d54-196">Membres</span><span class="sxs-lookup"><span data-stu-id="b9d54-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="b9d54-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b9d54-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="b9d54-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-200">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="b9d54-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b9d54-201">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b9d54-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-202">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-202">Type</span></span>

*   <span data-ttu-id="b9d54-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b9d54-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-204">Requirements</span></span>

|<span data-ttu-id="b9d54-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-205">Requirement</span></span>| <span data-ttu-id="b9d54-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-208">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-208">1.0</span></span>|
|[<span data-ttu-id="b9d54-209">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-210">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-212">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-213">Example</span></span>

<span data-ttu-id="b9d54-214">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="b9d54-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="b9d54-216">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b9d54-217">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-218">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-218">Type</span></span>

*   [<span data-ttu-id="b9d54-219">Destinataires</span><span class="sxs-lookup"><span data-stu-id="b9d54-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b9d54-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-220">Requirements</span></span>

|<span data-ttu-id="b9d54-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-221">Requirement</span></span>| <span data-ttu-id="b9d54-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-224">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d54-224">1.1</span></span>|
|[<span data-ttu-id="b9d54-225">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-226">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-228">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="b9d54-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="b9d54-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="b9d54-231">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-232">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-232">Type</span></span>

*   [<span data-ttu-id="b9d54-233">Body</span><span class="sxs-lookup"><span data-stu-id="b9d54-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="b9d54-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-234">Requirements</span></span>

|<span data-ttu-id="b9d54-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-235">Requirement</span></span>| <span data-ttu-id="b9d54-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-238">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d54-238">1.1</span></span>|
|[<span data-ttu-id="b9d54-239">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-240">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-242">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-242">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-243">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-243">Example</span></span>

<span data-ttu-id="b9d54-244">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="b9d54-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="b9d54-245">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="b9d54-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="b9d54-247">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b9d54-248">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-249">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-249">Read mode</span></span>

<span data-ttu-id="b9d54-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-252">Compose mode</span></span>

<span data-ttu-id="b9d54-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-254">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-254">Type</span></span>

*   <span data-ttu-id="b9d54-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-256">Requirements</span></span>

|<span data-ttu-id="b9d54-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-257">Requirement</span></span>| <span data-ttu-id="b9d54-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-260">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-260">1.0</span></span>|
|[<span data-ttu-id="b9d54-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-262">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-264">Compose or read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b9d54-265">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b9d54-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="b9d54-266">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="b9d54-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b9d54-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b9d54-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-271">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-271">Type</span></span>

*   <span data-ttu-id="b9d54-272">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-273">Requirements</span></span>

|<span data-ttu-id="b9d54-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-274">Requirement</span></span>| <span data-ttu-id="b9d54-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-277">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-277">1.0</span></span>|
|[<span data-ttu-id="b9d54-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-278">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-279">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-280">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-281">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-282">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="b9d54-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b9d54-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="b9d54-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-286">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-286">Type</span></span>

*   <span data-ttu-id="b9d54-287">Date</span><span class="sxs-lookup"><span data-stu-id="b9d54-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-288">Requirements</span></span>

|<span data-ttu-id="b9d54-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-289">Requirement</span></span>| <span data-ttu-id="b9d54-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-292">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-292">1.0</span></span>|
|[<span data-ttu-id="b9d54-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-294">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b9d54-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b9d54-298">dateTimeModified :Date</span></span>

<span data-ttu-id="b9d54-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-301">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-302">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-302">Type</span></span>

*   <span data-ttu-id="b9d54-303">Date</span><span class="sxs-lookup"><span data-stu-id="b9d54-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-304">Requirements</span></span>

|<span data-ttu-id="b9d54-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-305">Requirement</span></span>| <span data-ttu-id="b9d54-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-308">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-308">1.0</span></span>|
|[<span data-ttu-id="b9d54-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-310">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="b9d54-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="b9d54-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="b9d54-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b9d54-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-318">Read mode</span></span>

<span data-ttu-id="b9d54-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-320">Compose mode</span></span>

<span data-ttu-id="b9d54-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b9d54-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="b9d54-323">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-323">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="b9d54-324">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-324">Type</span></span>

*   <span data-ttu-id="b9d54-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="b9d54-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-326">Requirements</span></span>

|<span data-ttu-id="b9d54-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-327">Requirement</span></span>| <span data-ttu-id="b9d54-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-330">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-330">1.0</span></span>|
|[<span data-ttu-id="b9d54-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-332">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-334">Compose or read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="b9d54-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b9d54-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="b9d54-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b9d54-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-340">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-341">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-341">Type</span></span>

*   [<span data-ttu-id="b9d54-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b9d54-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b9d54-343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-343">Requirements</span></span>

|<span data-ttu-id="b9d54-344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-344">Requirement</span></span>| <span data-ttu-id="b9d54-345">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-347">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-347">1.0</span></span>|
|[<span data-ttu-id="b9d54-348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-349">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-351">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-352">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="b9d54-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b9d54-353">internetMessageId :String</span></span>

<span data-ttu-id="b9d54-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-356">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-356">Type</span></span>

*   <span data-ttu-id="b9d54-357">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-358">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-358">Requirements</span></span>

|<span data-ttu-id="b9d54-359">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-359">Requirement</span></span>| <span data-ttu-id="b9d54-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-362">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-362">1.0</span></span>|
|[<span data-ttu-id="b9d54-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-364">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b9d54-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b9d54-368">itemClass :String</span></span>

<span data-ttu-id="b9d54-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b9d54-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b9d54-373">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-373">Type</span></span> | <span data-ttu-id="b9d54-374">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-374">Description</span></span> | <span data-ttu-id="b9d54-375">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="b9d54-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b9d54-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="b9d54-376">Appointment items</span></span> | <span data-ttu-id="b9d54-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="b9d54-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="b9d54-378">Message items</span></span> | <span data-ttu-id="b9d54-379">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="b9d54-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b9d54-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-381">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-381">Type</span></span>

*   <span data-ttu-id="b9d54-382">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-383">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-383">Requirements</span></span>

|<span data-ttu-id="b9d54-384">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-384">Requirement</span></span>| <span data-ttu-id="b9d54-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-386">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-387">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-387">1.0</span></span>|
|[<span data-ttu-id="b9d54-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-389">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b9d54-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b9d54-393">(nullable) itemId :String</span></span>

<span data-ttu-id="b9d54-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-396">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9d54-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b9d54-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9d54-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b9d54-398">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b9d54-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b9d54-399">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b9d54-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b9d54-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-402">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-402">Type</span></span>

*   <span data-ttu-id="b9d54-403">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-404">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-404">Requirements</span></span>

|<span data-ttu-id="b9d54-405">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-405">Requirement</span></span>| <span data-ttu-id="b9d54-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-407">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-408">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-408">1.0</span></span>|
|[<span data-ttu-id="b9d54-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-410">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-412">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-413">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-413">Example</span></span>

<span data-ttu-id="b9d54-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="b9d54-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b9d54-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b9d54-417">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="b9d54-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b9d54-418">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-419">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-419">Type</span></span>

*   [<span data-ttu-id="b9d54-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b9d54-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b9d54-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-421">Requirements</span></span>

|<span data-ttu-id="b9d54-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-422">Requirement</span></span>| <span data-ttu-id="b9d54-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-425">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-425">1.0</span></span>|
|[<span data-ttu-id="b9d54-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-427">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-429">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="b9d54-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="b9d54-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="b9d54-432">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-433">Read mode</span></span>

<span data-ttu-id="b9d54-434">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-435">Compose mode</span></span>

<span data-ttu-id="b9d54-436">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-437">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-437">Type</span></span>

*   <span data-ttu-id="b9d54-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="b9d54-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-439">Requirements</span></span>

|<span data-ttu-id="b9d54-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-440">Requirement</span></span>| <span data-ttu-id="b9d54-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-443">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-443">1.0</span></span>|
|[<span data-ttu-id="b9d54-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-445">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-447">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-447">Compose or read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b9d54-448">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b9d54-448">normalizedSubject :String</span></span>

<span data-ttu-id="b9d54-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b9d54-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="b9d54-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-453">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-453">Type</span></span>

*   <span data-ttu-id="b9d54-454">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-455">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-455">Requirements</span></span>

|<span data-ttu-id="b9d54-456">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-456">Requirement</span></span>| <span data-ttu-id="b9d54-457">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-458">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-459">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-459">1.0</span></span>|
|[<span data-ttu-id="b9d54-460">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-461">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-462">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-463">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-464">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="b9d54-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b9d54-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="b9d54-466">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-467">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-467">Type</span></span>

*   [<span data-ttu-id="b9d54-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b9d54-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b9d54-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-469">Requirements</span></span>

|<span data-ttu-id="b9d54-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-470">Requirement</span></span>| <span data-ttu-id="b9d54-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-473">1.3</span><span class="sxs-lookup"><span data-stu-id="b9d54-473">1.3</span></span>|
|[<span data-ttu-id="b9d54-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-475">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-477">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-477">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="b9d54-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="b9d54-480">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b9d54-481">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-482">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-482">Read mode</span></span>

<span data-ttu-id="b9d54-483">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="b9d54-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-484">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-484">Compose mode</span></span>

<span data-ttu-id="b9d54-485">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="b9d54-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-486">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-486">Type</span></span>

*   <span data-ttu-id="b9d54-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-488">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-488">Requirements</span></span>

|<span data-ttu-id="b9d54-489">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-489">Requirement</span></span>| <span data-ttu-id="b9d54-490">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-491">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-492">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-492">1.0</span></span>|
|[<span data-ttu-id="b9d54-493">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-494">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-495">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-496">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-496">Compose or read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="b9d54-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b9d54-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="b9d54-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-500">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-500">Type</span></span>

*   [<span data-ttu-id="b9d54-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b9d54-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b9d54-502">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-502">Requirements</span></span>

|<span data-ttu-id="b9d54-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-503">Requirement</span></span>| <span data-ttu-id="b9d54-504">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-505">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-506">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-506">1.0</span></span>|
|[<span data-ttu-id="b9d54-507">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-508">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-509">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-510">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-511">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="b9d54-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="b9d54-513">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b9d54-514">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-515">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-515">Read mode</span></span>

<span data-ttu-id="b9d54-516">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="b9d54-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-517">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-517">Compose mode</span></span>

<span data-ttu-id="b9d54-518">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="b9d54-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-519">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-519">Type</span></span>

*   <span data-ttu-id="b9d54-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-521">Requirements</span></span>

|<span data-ttu-id="b9d54-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-522">Requirement</span></span>| <span data-ttu-id="b9d54-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-525">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-525">1.0</span></span>|
|[<span data-ttu-id="b9d54-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-527">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-529">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-529">Compose or read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="b9d54-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b9d54-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="b9d54-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b9d54-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-535">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d54-536">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-536">Type</span></span>

*   [<span data-ttu-id="b9d54-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b9d54-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b9d54-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-538">Requirements</span></span>

|<span data-ttu-id="b9d54-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-539">Requirement</span></span>| <span data-ttu-id="b9d54-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-542">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-542">1.0</span></span>|
|[<span data-ttu-id="b9d54-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-543">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-544">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-545">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="b9d54-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="b9d54-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="b9d54-549">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b9d54-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-552">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-552">Read mode</span></span>

<span data-ttu-id="b9d54-553">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-554">Compose mode</span></span>

<span data-ttu-id="b9d54-555">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b9d54-556">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="b9d54-557">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="b9d54-558">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-558">Type</span></span>

*   <span data-ttu-id="b9d54-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="b9d54-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-560">Requirements</span></span>

|<span data-ttu-id="b9d54-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-561">Requirement</span></span>| <span data-ttu-id="b9d54-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-564">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-564">1.0</span></span>|
|[<span data-ttu-id="b9d54-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-566">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-568">Compose or read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="b9d54-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b9d54-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="b9d54-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b9d54-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="b9d54-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-572">Read mode</span></span>

<span data-ttu-id="b9d54-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-575">Compose mode</span></span>

<span data-ttu-id="b9d54-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="b9d54-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-577">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-577">Type</span></span>

*   <span data-ttu-id="b9d54-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b9d54-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-579">Requirements</span></span>

|<span data-ttu-id="b9d54-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-580">Requirement</span></span>| <span data-ttu-id="b9d54-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-583">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-583">1.0</span></span>|
|[<span data-ttu-id="b9d54-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-585">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-587">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="b9d54-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="b9d54-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b9d54-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b9d54-591">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-591">Read mode</span></span>

<span data-ttu-id="b9d54-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="b9d54-594">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-594">Compose mode</span></span>

<span data-ttu-id="b9d54-595">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b9d54-596">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-596">Type</span></span>

*   <span data-ttu-id="b9d54-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b9d54-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-598">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-598">Requirements</span></span>

|<span data-ttu-id="b9d54-599">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-599">Requirement</span></span>| <span data-ttu-id="b9d54-600">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-601">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-602">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-602">1.0</span></span>|
|[<span data-ttu-id="b9d54-603">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-604">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-605">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-606">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-606">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="b9d54-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="b9d54-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b9d54-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b9d54-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b9d54-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="b9d54-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b9d54-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="b9d54-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b9d54-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="b9d54-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-612">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-612">Parameters</span></span>

|<span data-ttu-id="b9d54-613">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-613">Name</span></span>| <span data-ttu-id="b9d54-614">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-614">Type</span></span>| <span data-ttu-id="b9d54-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-615">Attributes</span></span>| <span data-ttu-id="b9d54-616">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b9d54-617">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-617">String</span></span>||<span data-ttu-id="b9d54-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b9d54-620">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-620">String</span></span>||<span data-ttu-id="b9d54-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b9d54-623">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-623">Object</span></span>| <span data-ttu-id="b9d54-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-624">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="b9d54-626">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-626">Object</span></span> | <span data-ttu-id="b9d54-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-627">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="b9d54-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="b9d54-629">Boolean</span></span> | <span data-ttu-id="b9d54-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-630">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-631">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="b9d54-632">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-632">function</span></span>| <span data-ttu-id="b9d54-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-633">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-634">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b9d54-635">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b9d54-636">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b9d54-637">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b9d54-637">Errors</span></span>

| <span data-ttu-id="b9d54-638">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b9d54-638">Error code</span></span> | <span data-ttu-id="b9d54-639">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b9d54-640">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b9d54-641">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b9d54-642">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-643">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-643">Requirements</span></span>

|<span data-ttu-id="b9d54-644">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-644">Requirement</span></span>| <span data-ttu-id="b9d54-645">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-646">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-647">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d54-647">1.1</span></span>|
|[<span data-ttu-id="b9d54-648">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-650">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-651">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b9d54-652">Exemples</span><span class="sxs-lookup"><span data-stu-id="b9d54-652">Examples</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="b9d54-653">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b9d54-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b9d54-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b9d54-655">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b9d54-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b9d54-659">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="b9d54-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b9d54-660">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-661">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-661">Parameters</span></span>

|<span data-ttu-id="b9d54-662">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-662">Name</span></span>| <span data-ttu-id="b9d54-663">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-663">Type</span></span>| <span data-ttu-id="b9d54-664">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-664">Attributes</span></span>| <span data-ttu-id="b9d54-665">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b9d54-666">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-666">String</span></span>||<span data-ttu-id="b9d54-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b9d54-669">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-669">String</span></span>||<span data-ttu-id="b9d54-670">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="b9d54-670">The sujbect of the item to be attached.</span></span> <span data-ttu-id="b9d54-671">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b9d54-672">Object</span><span class="sxs-lookup"><span data-stu-id="b9d54-672">Object</span></span>| <span data-ttu-id="b9d54-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-673">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-674">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b9d54-675">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-675">Object</span></span>| <span data-ttu-id="b9d54-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-676">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-677">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b9d54-678">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-678">function</span></span>| <span data-ttu-id="b9d54-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-679">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-680">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b9d54-681">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b9d54-682">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b9d54-683">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b9d54-683">Errors</span></span>

| <span data-ttu-id="b9d54-684">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b9d54-684">Error code</span></span> | <span data-ttu-id="b9d54-685">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b9d54-686">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-687">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-687">Requirements</span></span>

|<span data-ttu-id="b9d54-688">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-688">Requirement</span></span>| <span data-ttu-id="b9d54-689">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-690">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-691">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d54-691">1.1</span></span>|
|[<span data-ttu-id="b9d54-692">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-694">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-695">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-696">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-696">Example</span></span>

<span data-ttu-id="b9d54-697">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="b9d54-698">close()</span><span class="sxs-lookup"><span data-stu-id="b9d54-698">close()</span></span>

<span data-ttu-id="b9d54-699">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="b9d54-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b9d54-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-702">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b9d54-703">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="b9d54-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-704">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-704">Requirements</span></span>

|<span data-ttu-id="b9d54-705">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-705">Requirement</span></span>| <span data-ttu-id="b9d54-706">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-707">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-708">1.3</span><span class="sxs-lookup"><span data-stu-id="b9d54-708">1.3</span></span>|
|[<span data-ttu-id="b9d54-709">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-710">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b9d54-710">Restricted</span></span>|
|[<span data-ttu-id="b9d54-711">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-712">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="b9d54-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="b9d54-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="b9d54-714">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b9d54-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-715">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b9d54-716">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b9d54-717">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="b9d54-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b9d54-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-721">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-721">Parameters</span></span>

| <span data-ttu-id="b9d54-722">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-722">Name</span></span> | <span data-ttu-id="b9d54-723">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-723">Type</span></span> | <span data-ttu-id="b9d54-724">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-724">Attributes</span></span> | <span data-ttu-id="b9d54-725">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b9d54-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b9d54-726">String &#124; Object</span></span>| |<span data-ttu-id="b9d54-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b9d54-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="b9d54-729">**OR**</span></span><br/><span data-ttu-id="b9d54-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="b9d54-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b9d54-732">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-732">String</span></span> | <span data-ttu-id="b9d54-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-733">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b9d54-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b9d54-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-737">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-738">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b9d54-739">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9d54-739">String</span></span> | | <span data-ttu-id="b9d54-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b9d54-742">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-742">String</span></span> | | <span data-ttu-id="b9d54-743">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b9d54-744">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9d54-744">String</span></span> | | <span data-ttu-id="b9d54-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b9d54-747">Booléen</span><span class="sxs-lookup"><span data-stu-id="b9d54-747">Boolean</span></span> | | <span data-ttu-id="b9d54-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b9d54-750">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-750">String</span></span> | | <span data-ttu-id="b9d54-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b9d54-754">function</span><span class="sxs-lookup"><span data-stu-id="b9d54-754">function</span></span> | <span data-ttu-id="b9d54-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-755">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-756">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-757">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-757">Requirements</span></span>

|<span data-ttu-id="b9d54-758">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-758">Requirement</span></span>| <span data-ttu-id="b9d54-759">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-760">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-761">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-761">1.0</span></span>|
|[<span data-ttu-id="b9d54-762">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-763">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-764">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-765">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b9d54-766">Exemples</span><span class="sxs-lookup"><span data-stu-id="b9d54-766">Examples</span></span>

<span data-ttu-id="b9d54-767">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b9d54-768">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="b9d54-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b9d54-769">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="b9d54-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b9d54-770">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="b9d54-770">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="b9d54-771">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-771">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="b9d54-772">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="b9d54-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="b9d54-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="b9d54-774">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b9d54-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-775">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b9d54-776">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b9d54-777">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="b9d54-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b9d54-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-781">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-781">Parameters</span></span>

| <span data-ttu-id="b9d54-782">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-782">Name</span></span> | <span data-ttu-id="b9d54-783">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-783">Type</span></span> | <span data-ttu-id="b9d54-784">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-784">Attributes</span></span> | <span data-ttu-id="b9d54-785">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b9d54-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b9d54-786">String &#124; Object</span></span>| | <span data-ttu-id="b9d54-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b9d54-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="b9d54-789">**OR**</span></span><br/><span data-ttu-id="b9d54-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="b9d54-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b9d54-792">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-792">String</span></span> | <span data-ttu-id="b9d54-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-793">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b9d54-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b9d54-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-797">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-798">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b9d54-799">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9d54-799">String</span></span> | | <span data-ttu-id="b9d54-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b9d54-802">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-802">String</span></span> | | <span data-ttu-id="b9d54-803">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b9d54-804">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9d54-804">String</span></span> | | <span data-ttu-id="b9d54-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b9d54-807">Booléen</span><span class="sxs-lookup"><span data-stu-id="b9d54-807">Boolean</span></span> | | <span data-ttu-id="b9d54-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b9d54-810">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-810">String</span></span> | | <span data-ttu-id="b9d54-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b9d54-814">function</span><span class="sxs-lookup"><span data-stu-id="b9d54-814">function</span></span> | <span data-ttu-id="b9d54-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-815">&lt;optional&gt;</span></span> | <span data-ttu-id="b9d54-816">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-817">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-817">Requirements</span></span>

|<span data-ttu-id="b9d54-818">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-818">Requirement</span></span>| <span data-ttu-id="b9d54-819">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-820">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-821">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-821">1.0</span></span>|
|[<span data-ttu-id="b9d54-822">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-823">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-824">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-825">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b9d54-826">Exemples</span><span class="sxs-lookup"><span data-stu-id="b9d54-826">Examples</span></span>

<span data-ttu-id="b9d54-827">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b9d54-828">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="b9d54-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b9d54-829">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="b9d54-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b9d54-830">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="b9d54-830">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="b9d54-831">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-831">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="b9d54-832">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="b9d54-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b9d54-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="b9d54-834">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b9d54-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-835">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-836">Requirements</span></span>

|<span data-ttu-id="b9d54-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-837">Requirement</span></span>| <span data-ttu-id="b9d54-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-840">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-840">1.0</span></span>|
|[<span data-ttu-id="b9d54-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-842">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-845">Returns:</span></span>

<span data-ttu-id="b9d54-846">Type : [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b9d54-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b9d54-847">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-847">Example</span></span>

<span data-ttu-id="b9d54-848">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="b9d54-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b9d54-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b9d54-850">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b9d54-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-851">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-852">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-852">Parameters</span></span>

|<span data-ttu-id="b9d54-853">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-853">Name</span></span>| <span data-ttu-id="b9d54-854">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-854">Type</span></span>| <span data-ttu-id="b9d54-855">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b9d54-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b9d54-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="b9d54-857">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="b9d54-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-858">Requirements</span></span>

|<span data-ttu-id="b9d54-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-859">Requirement</span></span>| <span data-ttu-id="b9d54-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-862">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-862">1.0</span></span>|
|[<span data-ttu-id="b9d54-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-864">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b9d54-864">Restricted</span></span>|
|[<span data-ttu-id="b9d54-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-867">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-867">Returns:</span></span>

<span data-ttu-id="b9d54-868">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="b9d54-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b9d54-869">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="b9d54-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b9d54-870">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b9d54-871">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="b9d54-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b9d54-872">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="b9d54-872">Value of `entityType`</span></span> | <span data-ttu-id="b9d54-873">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="b9d54-873">Type of objects in returned array</span></span> | <span data-ttu-id="b9d54-874">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="b9d54-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b9d54-875">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-875">String</span></span> | <span data-ttu-id="b9d54-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b9d54-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b9d54-877">Contact</span><span class="sxs-lookup"><span data-stu-id="b9d54-877">Contact</span></span> | <span data-ttu-id="b9d54-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b9d54-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b9d54-879">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-879">String</span></span> | <span data-ttu-id="b9d54-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b9d54-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b9d54-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b9d54-881">MeetingSuggestion</span></span> | <span data-ttu-id="b9d54-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b9d54-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b9d54-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b9d54-883">PhoneNumber</span></span> | <span data-ttu-id="b9d54-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b9d54-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b9d54-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b9d54-885">TaskSuggestion</span></span> | <span data-ttu-id="b9d54-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b9d54-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b9d54-887">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-887">String</span></span> | <span data-ttu-id="b9d54-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b9d54-888">**Restricted**</span></span> |

<span data-ttu-id="b9d54-889">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b9d54-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b9d54-890">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-890">Example</span></span>

<span data-ttu-id="b9d54-891">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b9d54-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="b9d54-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b9d54-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b9d54-893">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-894">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b9d54-895">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="b9d54-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-896">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-896">Parameters</span></span>

|<span data-ttu-id="b9d54-897">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-897">Name</span></span>| <span data-ttu-id="b9d54-898">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-898">Type</span></span>| <span data-ttu-id="b9d54-899">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b9d54-900">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-900">String</span></span>|<span data-ttu-id="b9d54-901">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="b9d54-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-902">Requirements</span></span>

|<span data-ttu-id="b9d54-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-903">Requirement</span></span>| <span data-ttu-id="b9d54-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-906">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-906">1.0</span></span>|
|[<span data-ttu-id="b9d54-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-908">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-910">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-911">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-911">Returns:</span></span>

<span data-ttu-id="b9d54-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b9d54-914">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b9d54-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b9d54-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b9d54-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b9d54-916">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-917">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b9d54-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b9d54-921">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="b9d54-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b9d54-922">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b9d54-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d54-926">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-926">Requirements</span></span>

|<span data-ttu-id="b9d54-927">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-927">Requirement</span></span>| <span data-ttu-id="b9d54-928">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-929">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-930">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-930">1.0</span></span>|
|[<span data-ttu-id="b9d54-931">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-932">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-933">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-934">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-935">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-935">Returns:</span></span>

<span data-ttu-id="b9d54-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b9d54-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b9d54-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b9d54-939">Object</span><span class="sxs-lookup"><span data-stu-id="b9d54-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b9d54-940">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-940">Example</span></span>

<span data-ttu-id="b9d54-941">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="b9d54-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b9d54-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b9d54-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b9d54-943">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-944">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b9d54-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b9d54-945">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="b9d54-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b9d54-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-948">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-948">Parameters</span></span>

|<span data-ttu-id="b9d54-949">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-949">Name</span></span>| <span data-ttu-id="b9d54-950">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-950">Type</span></span>| <span data-ttu-id="b9d54-951">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b9d54-952">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-952">String</span></span>|<span data-ttu-id="b9d54-953">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="b9d54-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-954">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-954">Requirements</span></span>

|<span data-ttu-id="b9d54-955">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-955">Requirement</span></span>| <span data-ttu-id="b9d54-956">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-957">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-958">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-958">1.0</span></span>|
|[<span data-ttu-id="b9d54-959">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-960">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-961">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-962">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-963">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-963">Returns:</span></span>

<span data-ttu-id="b9d54-964">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b9d54-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b9d54-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b9d54-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b9d54-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b9d54-967">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b9d54-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b9d54-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b9d54-969">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="b9d54-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b9d54-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-972">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-972">Parameters</span></span>

|<span data-ttu-id="b9d54-973">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-973">Name</span></span>| <span data-ttu-id="b9d54-974">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-974">Type</span></span>| <span data-ttu-id="b9d54-975">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-975">Attributes</span></span>| <span data-ttu-id="b9d54-976">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b9d54-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b9d54-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b9d54-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b9d54-981">Object</span><span class="sxs-lookup"><span data-stu-id="b9d54-981">Object</span></span>| <span data-ttu-id="b9d54-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-982">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-983">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b9d54-984">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-984">Object</span></span>| <span data-ttu-id="b9d54-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-985">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-986">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b9d54-987">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-987">function</span></span>||<span data-ttu-id="b9d54-988">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b9d54-989">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b9d54-990">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-991">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-991">Requirements</span></span>

|<span data-ttu-id="b9d54-992">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-992">Requirement</span></span>| <span data-ttu-id="b9d54-993">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-994">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-995">1.2</span><span class="sxs-lookup"><span data-stu-id="b9d54-995">1.2</span></span>|
|[<span data-ttu-id="b9d54-996">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-998">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-999">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b9d54-1000">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b9d54-1000">Returns:</span></span>

<span data-ttu-id="b9d54-1001">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b9d54-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b9d54-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b9d54-1003">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b9d54-1004">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-1004">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b9d54-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b9d54-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b9d54-1006">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b9d54-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-1010">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1010">Parameters</span></span>

|<span data-ttu-id="b9d54-1011">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-1011">Name</span></span>| <span data-ttu-id="b9d54-1012">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-1012">Type</span></span>| <span data-ttu-id="b9d54-1013">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-1013">Attributes</span></span>| <span data-ttu-id="b9d54-1014">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b9d54-1015">function</span><span class="sxs-lookup"><span data-stu-id="b9d54-1015">function</span></span>||<span data-ttu-id="b9d54-1016">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b9d54-1017">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b9d54-1018">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b9d54-1019">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1019">Object</span></span>| <span data-ttu-id="b9d54-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1021">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b9d54-1022">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-1023">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-1023">Requirements</span></span>

|<span data-ttu-id="b9d54-1024">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-1024">Requirement</span></span>| <span data-ttu-id="b9d54-1025">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-1026">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="b9d54-1027">1.0</span></span>|
|[<span data-ttu-id="b9d54-1028">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-1029">ReadItem</span></span>|
|[<span data-ttu-id="b9d54-1030">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-1031">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9d54-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-1032">Example</span></span>

<span data-ttu-id="b9d54-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b9d54-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b9d54-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b9d54-1037">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b9d54-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-1042">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1042">Parameters</span></span>

|<span data-ttu-id="b9d54-1043">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-1043">Name</span></span>| <span data-ttu-id="b9d54-1044">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-1044">Type</span></span>| <span data-ttu-id="b9d54-1045">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-1045">Attributes</span></span>| <span data-ttu-id="b9d54-1046">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b9d54-1047">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-1047">String</span></span>||<span data-ttu-id="b9d54-1048">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="b9d54-1049">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1049">Object</span></span>| <span data-ttu-id="b9d54-1050">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1051">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b9d54-1052">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1052">Object</span></span>| <span data-ttu-id="b9d54-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1054">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b9d54-1055">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-1055">function</span></span>| <span data-ttu-id="b9d54-1056">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1057">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b9d54-1058">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b9d54-1059">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b9d54-1059">Errors</span></span>

| <span data-ttu-id="b9d54-1060">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b9d54-1060">Error code</span></span> | <span data-ttu-id="b9d54-1061">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b9d54-1062">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-1063">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-1063">Requirements</span></span>

|<span data-ttu-id="b9d54-1064">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-1064">Requirement</span></span>| <span data-ttu-id="b9d54-1065">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-1066">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d54-1067">1.1</span></span>|
|[<span data-ttu-id="b9d54-1068">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-1068">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-1070">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-1070">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-1071">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-1072">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-1072">Example</span></span>

<span data-ttu-id="b9d54-1073">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="b9d54-1073">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b9d54-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b9d54-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="b9d54-1075">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="b9d54-p168">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-1079">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b9d54-1080">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b9d54-p170">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d54-1084">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="b9d54-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b9d54-1085">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1085">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b9d54-1086">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1086">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b9d54-1087">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1087">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-1088">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1088">Parameters</span></span>

|<span data-ttu-id="b9d54-1089">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-1089">Name</span></span>| <span data-ttu-id="b9d54-1090">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-1090">Type</span></span>| <span data-ttu-id="b9d54-1091">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-1091">Attributes</span></span>| <span data-ttu-id="b9d54-1092">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-1092">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="b9d54-1093">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1093">Object</span></span>| <span data-ttu-id="b9d54-1094">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1095">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b9d54-1096">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1096">Object</span></span>| <span data-ttu-id="b9d54-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1098">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b9d54-1099">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-1099">function</span></span>||<span data-ttu-id="b9d54-1100">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-1100">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b9d54-1101">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1101">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9d54-1102">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-1102">Requirements</span></span>

|<span data-ttu-id="b9d54-1103">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-1103">Requirement</span></span>| <span data-ttu-id="b9d54-1104">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-1104">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-1105">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1105">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-1106">1.3</span><span class="sxs-lookup"><span data-stu-id="b9d54-1106">1.3</span></span>|
|[<span data-ttu-id="b9d54-1107">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-1107">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-1108">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-1108">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-1109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-1109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-1110">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-1110">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b9d54-1111">範例</span><span class="sxs-lookup"><span data-stu-id="b9d54-1111">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="b9d54-p172">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b9d54-1114">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b9d54-1114">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b9d54-1115">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1115">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b9d54-p173">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b9d54-1119">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1119">Parameters</span></span>

|<span data-ttu-id="b9d54-1120">Nom</span><span class="sxs-lookup"><span data-stu-id="b9d54-1120">Name</span></span>| <span data-ttu-id="b9d54-1121">Type</span><span class="sxs-lookup"><span data-stu-id="b9d54-1121">Type</span></span>| <span data-ttu-id="b9d54-1122">Attributs</span><span class="sxs-lookup"><span data-stu-id="b9d54-1122">Attributes</span></span>| <span data-ttu-id="b9d54-1123">Description</span><span class="sxs-lookup"><span data-stu-id="b9d54-1123">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b9d54-1124">String</span><span class="sxs-lookup"><span data-stu-id="b9d54-1124">String</span></span>||<span data-ttu-id="b9d54-p174">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b9d54-1128">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1128">Object</span></span>| <span data-ttu-id="b9d54-1129">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1129">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1130">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1130">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b9d54-1131">Objet</span><span class="sxs-lookup"><span data-stu-id="b9d54-1131">Object</span></span>| <span data-ttu-id="b9d54-1132">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1132">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-1133">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1133">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b9d54-1134">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b9d54-1134">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b9d54-1135">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b9d54-1135">&lt;optional&gt;</span></span>|<span data-ttu-id="b9d54-p175">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b9d54-p176">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="b9d54-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b9d54-1140">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="b9d54-1140">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b9d54-1141">fonction</span><span class="sxs-lookup"><span data-stu-id="b9d54-1141">function</span></span>||<span data-ttu-id="b9d54-1142">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b9d54-1142">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9d54-1143">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9d54-1143">Requirements</span></span>

|<span data-ttu-id="b9d54-1144">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9d54-1144">Requirement</span></span>| <span data-ttu-id="b9d54-1145">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9d54-1145">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d54-1146">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9d54-1146">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9d54-1147">1.2</span><span class="sxs-lookup"><span data-stu-id="b9d54-1147">1.2</span></span>|
|[<span data-ttu-id="b9d54-1148">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9d54-1148">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9d54-1149">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b9d54-1149">ReadWriteItem</span></span>|
|[<span data-ttu-id="b9d54-1150">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9d54-1150">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9d54-1151">Composition</span><span class="sxs-lookup"><span data-stu-id="b9d54-1151">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d54-1152">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9d54-1152">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
