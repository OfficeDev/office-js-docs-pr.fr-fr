---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,6
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 980135223414b58bb048dce54a9fe1446a26086c
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167360"
---
# <a name="item"></a><span data-ttu-id="f912e-102">élément</span><span class="sxs-lookup"><span data-stu-id="f912e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f912e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f912e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f912e-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f912e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-106">Requirements</span></span>

|<span data-ttu-id="f912e-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-107">Requirement</span></span>| <span data-ttu-id="f912e-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-110">1.0</span></span>|
|[<span data-ttu-id="f912e-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f912e-112">Restricted</span></span>|
|[<span data-ttu-id="f912e-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f912e-115">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="f912e-115">Members and methods</span></span>

| <span data-ttu-id="f912e-116">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-116">Member</span></span> | <span data-ttu-id="f912e-117">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f912e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="f912e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="f912e-119">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-119">Member</span></span> |
| [<span data-ttu-id="f912e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="f912e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="f912e-121">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-121">Member</span></span> |
| [<span data-ttu-id="f912e-122">body</span><span class="sxs-lookup"><span data-stu-id="f912e-122">body</span></span>](#body-body) | <span data-ttu-id="f912e-123">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-123">Member</span></span> |
| [<span data-ttu-id="f912e-124">cc</span><span class="sxs-lookup"><span data-stu-id="f912e-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f912e-125">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-125">Member</span></span> |
| [<span data-ttu-id="f912e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="f912e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f912e-127">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-127">Member</span></span> |
| [<span data-ttu-id="f912e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f912e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f912e-129">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-129">Member</span></span> |
| [<span data-ttu-id="f912e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f912e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f912e-131">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-131">Member</span></span> |
| [<span data-ttu-id="f912e-132">end</span><span class="sxs-lookup"><span data-stu-id="f912e-132">end</span></span>](#end-datetime) | <span data-ttu-id="f912e-133">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-133">Member</span></span> |
| [<span data-ttu-id="f912e-134">from</span><span class="sxs-lookup"><span data-stu-id="f912e-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="f912e-135">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-135">Member</span></span> |
| [<span data-ttu-id="f912e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f912e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f912e-137">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-137">Member</span></span> |
| [<span data-ttu-id="f912e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="f912e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f912e-139">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-139">Member</span></span> |
| [<span data-ttu-id="f912e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="f912e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f912e-141">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-141">Member</span></span> |
| [<span data-ttu-id="f912e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="f912e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="f912e-143">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-143">Member</span></span> |
| [<span data-ttu-id="f912e-144">location</span><span class="sxs-lookup"><span data-stu-id="f912e-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="f912e-145">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-145">Member</span></span> |
| [<span data-ttu-id="f912e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f912e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f912e-147">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-147">Member</span></span> |
| [<span data-ttu-id="f912e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f912e-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="f912e-149">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-149">Member</span></span> |
| [<span data-ttu-id="f912e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f912e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f912e-151">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-151">Member</span></span> |
| [<span data-ttu-id="f912e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="f912e-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="f912e-153">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-153">Member</span></span> |
| [<span data-ttu-id="f912e-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f912e-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f912e-155">Member</span><span class="sxs-lookup"><span data-stu-id="f912e-155">Member</span></span> |
| [<span data-ttu-id="f912e-156">sender</span><span class="sxs-lookup"><span data-stu-id="f912e-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="f912e-157">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-157">Member</span></span> |
| [<span data-ttu-id="f912e-158">start</span><span class="sxs-lookup"><span data-stu-id="f912e-158">start</span></span>](#start-datetime) | <span data-ttu-id="f912e-159">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-159">Member</span></span> |
| [<span data-ttu-id="f912e-160">subject</span><span class="sxs-lookup"><span data-stu-id="f912e-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="f912e-161">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-161">Member</span></span> |
| [<span data-ttu-id="f912e-162">to</span><span class="sxs-lookup"><span data-stu-id="f912e-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f912e-163">Membre</span><span class="sxs-lookup"><span data-stu-id="f912e-163">Member</span></span> |
| [<span data-ttu-id="f912e-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f912e-165">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-165">Method</span></span> |
| [<span data-ttu-id="f912e-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f912e-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-167">Method</span></span> |
| [<span data-ttu-id="f912e-168">close</span><span class="sxs-lookup"><span data-stu-id="f912e-168">close</span></span>](#close) | <span data-ttu-id="f912e-169">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-169">Method</span></span> |
| [<span data-ttu-id="f912e-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f912e-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="f912e-171">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-171">Method</span></span> |
| [<span data-ttu-id="f912e-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f912e-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="f912e-173">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-173">Method</span></span> |
| [<span data-ttu-id="f912e-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="f912e-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="f912e-175">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-175">Method</span></span> |
| [<span data-ttu-id="f912e-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f912e-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f912e-177">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-177">Method</span></span> |
| [<span data-ttu-id="f912e-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f912e-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f912e-179">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-179">Method</span></span> |
| [<span data-ttu-id="f912e-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f912e-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f912e-181">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-181">Method</span></span> |
| [<span data-ttu-id="f912e-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f912e-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f912e-183">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-183">Method</span></span> |
| [<span data-ttu-id="f912e-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f912e-185">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-185">Method</span></span> |
| [<span data-ttu-id="f912e-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="f912e-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="f912e-187">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-187">Method</span></span> |
| [<span data-ttu-id="f912e-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f912e-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f912e-189">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-189">Method</span></span> |
| [<span data-ttu-id="f912e-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f912e-191">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-191">Method</span></span> |
| [<span data-ttu-id="f912e-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f912e-193">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-193">Method</span></span> |
| [<span data-ttu-id="f912e-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f912e-195">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-195">Method</span></span> |
| [<span data-ttu-id="f912e-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f912e-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f912e-197">Méthode</span><span class="sxs-lookup"><span data-stu-id="f912e-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f912e-198">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-198">Example</span></span>

<span data-ttu-id="f912e-199">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f912e-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="f912e-200">Membres</span><span class="sxs-lookup"><span data-stu-id="f912e-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="f912e-201">pièces jointes : tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="f912e-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="f912e-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-204">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="f912e-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f912e-205">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f912e-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-206">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-206">Type</span></span>

*   <span data-ttu-id="f912e-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="f912e-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-208">Requirements</span></span>

|<span data-ttu-id="f912e-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-209">Requirement</span></span>| <span data-ttu-id="f912e-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-212">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-212">1.0</span></span>|
|[<span data-ttu-id="f912e-213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-214">ReadItem</span></span>|
|[<span data-ttu-id="f912e-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-216">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-217">Example</span></span>

<span data-ttu-id="f912e-218">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f912e-219">CCI : [destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-220">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f912e-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f912e-221">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-222">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-222">Type</span></span>

*   [<span data-ttu-id="f912e-223">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f912e-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-224">Requirements</span></span>

|<span data-ttu-id="f912e-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-225">Requirement</span></span>| <span data-ttu-id="f912e-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-228">1.1</span><span class="sxs-lookup"><span data-stu-id="f912e-228">1.1</span></span>|
|[<span data-ttu-id="f912e-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-230">ReadItem</span></span>|
|[<span data-ttu-id="f912e-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-232">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-233">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="f912e-234">Body : [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-235">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-236">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-236">Type</span></span>

*   [<span data-ttu-id="f912e-237">Body</span><span class="sxs-lookup"><span data-stu-id="f912e-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-238">Requirements</span></span>

|<span data-ttu-id="f912e-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-239">Requirement</span></span>| <span data-ttu-id="f912e-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-242">1.1</span><span class="sxs-lookup"><span data-stu-id="f912e-242">1.1</span></span>|
|[<span data-ttu-id="f912e-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-244">ReadItem</span></span>|
|[<span data-ttu-id="f912e-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-247">Example</span></span>

<span data-ttu-id="f912e-248">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="f912e-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f912e-249">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f912e-250">CC : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-251">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f912e-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f912e-252">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-253">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-253">Read mode</span></span>

<span data-ttu-id="f912e-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f912e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-256">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-256">Compose mode</span></span>

<span data-ttu-id="f912e-257">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="f912e-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f912e-258">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-258">Type</span></span>

*   <span data-ttu-id="f912e-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-260">Requirements</span></span>

|<span data-ttu-id="f912e-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-261">Requirement</span></span>| <span data-ttu-id="f912e-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-264">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-264">1.0</span></span>|
|[<span data-ttu-id="f912e-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-266">ReadItem</span></span>|
|[<span data-ttu-id="f912e-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f912e-269">(Nullable) conversationId : chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="f912e-270">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="f912e-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f912e-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="f912e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f912e-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="f912e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-275">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-275">Type</span></span>

*   <span data-ttu-id="f912e-276">String</span><span class="sxs-lookup"><span data-stu-id="f912e-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-277">Requirements</span></span>

|<span data-ttu-id="f912e-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-278">Requirement</span></span>| <span data-ttu-id="f912e-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-281">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-281">1.0</span></span>|
|[<span data-ttu-id="f912e-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-283">ReadItem</span></span>|
|[<span data-ttu-id="f912e-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="f912e-287">dateTimeCreated : date</span><span class="sxs-lookup"><span data-stu-id="f912e-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="f912e-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-290">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-290">Type</span></span>

*   <span data-ttu-id="f912e-291">Date</span><span class="sxs-lookup"><span data-stu-id="f912e-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-292">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-292">Requirements</span></span>

|<span data-ttu-id="f912e-293">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-293">Requirement</span></span>| <span data-ttu-id="f912e-294">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-295">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-296">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-296">1.0</span></span>|
|[<span data-ttu-id="f912e-297">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-298">ReadItem</span></span>|
|[<span data-ttu-id="f912e-299">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-300">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-301">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="f912e-302">dateTimeModified : date</span><span class="sxs-lookup"><span data-stu-id="f912e-302">dateTimeModified: Date</span></span>

<span data-ttu-id="f912e-303">Obtient la date et l’heure de la dernière modification d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="f912e-304">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-305">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-306">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-306">Type</span></span>

*   <span data-ttu-id="f912e-307">Date</span><span class="sxs-lookup"><span data-stu-id="f912e-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-308">Requirements</span></span>

|<span data-ttu-id="f912e-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-309">Requirement</span></span>| <span data-ttu-id="f912e-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-312">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-312">1.0</span></span>|
|[<span data-ttu-id="f912e-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-314">ReadItem</span></span>|
|[<span data-ttu-id="f912e-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-316">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="f912e-318">fin : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-319">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f912e-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f912e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-322">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-322">Read mode</span></span>

<span data-ttu-id="f912e-323">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f912e-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-324">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-324">Compose mode</span></span>

<span data-ttu-id="f912e-325">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f912e-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f912e-326">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f912e-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f912e-327">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f912e-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="f912e-328">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-328">Type</span></span>

*   <span data-ttu-id="f912e-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-330">Requirements</span></span>

|<span data-ttu-id="f912e-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-331">Requirement</span></span>| <span data-ttu-id="f912e-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-334">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-334">1.0</span></span>|
|[<span data-ttu-id="f912e-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-336">ReadItem</span></span>|
|[<span data-ttu-id="f912e-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-338">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f912e-339">de : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f912e-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f912e-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-344">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f912e-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-345">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-345">Type</span></span>

*   [<span data-ttu-id="f912e-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f912e-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="f912e-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="f912e-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-348">Requirements</span></span>

|<span data-ttu-id="f912e-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-349">Requirement</span></span>| <span data-ttu-id="f912e-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-352">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-352">1.0</span></span>|
|[<span data-ttu-id="f912e-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-354">ReadItem</span></span>|
|[<span data-ttu-id="f912e-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-356">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="f912e-357">internetMessageId : chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-357">internetMessageId: String</span></span>

<span data-ttu-id="f912e-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-360">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-360">Type</span></span>

*   <span data-ttu-id="f912e-361">String</span><span class="sxs-lookup"><span data-stu-id="f912e-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-362">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-362">Requirements</span></span>

|<span data-ttu-id="f912e-363">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-363">Requirement</span></span>| <span data-ttu-id="f912e-364">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-365">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-366">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-366">1.0</span></span>|
|[<span data-ttu-id="f912e-367">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-368">ReadItem</span></span>|
|[<span data-ttu-id="f912e-369">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-370">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-371">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="f912e-372">itemClass : chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-372">itemClass: String</span></span>

<span data-ttu-id="f912e-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f912e-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f912e-377">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-377">Type</span></span> | <span data-ttu-id="f912e-378">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-378">Description</span></span> | <span data-ttu-id="f912e-379">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="f912e-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f912e-380">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f912e-380">Appointment items</span></span> | <span data-ttu-id="f912e-381">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="f912e-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="f912e-382">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="f912e-382">Message items</span></span> | <span data-ttu-id="f912e-383">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="f912e-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f912e-384">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="f912e-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-385">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-385">Type</span></span>

*   <span data-ttu-id="f912e-386">String</span><span class="sxs-lookup"><span data-stu-id="f912e-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-387">Requirements</span></span>

|<span data-ttu-id="f912e-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-388">Requirement</span></span>| <span data-ttu-id="f912e-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-391">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-391">1.0</span></span>|
|[<span data-ttu-id="f912e-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-393">ReadItem</span></span>|
|[<span data-ttu-id="f912e-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-395">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-396">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f912e-397">(Nullable) itemId : String</span><span class="sxs-lookup"><span data-stu-id="f912e-397">(nullable) itemId: String</span></span>

<span data-ttu-id="f912e-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-400">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="f912e-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f912e-401">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="f912e-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f912e-402">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f912e-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f912e-403">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f912e-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f912e-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-406">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-406">Type</span></span>

*   <span data-ttu-id="f912e-407">String</span><span class="sxs-lookup"><span data-stu-id="f912e-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-408">Requirements</span></span>

|<span data-ttu-id="f912e-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-409">Requirement</span></span>| <span data-ttu-id="f912e-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-412">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-412">1.0</span></span>|
|[<span data-ttu-id="f912e-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-414">ReadItem</span></span>|
|[<span data-ttu-id="f912e-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-417">Example</span></span>

<span data-ttu-id="f912e-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f912e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="f912e-420">itemType : [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-421">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="f912e-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f912e-422">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-423">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-423">Type</span></span>

*   [<span data-ttu-id="f912e-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f912e-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-425">Requirements</span></span>

|<span data-ttu-id="f912e-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-426">Requirement</span></span>| <span data-ttu-id="f912e-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-429">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-429">1.0</span></span>|
|[<span data-ttu-id="f912e-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-431">ReadItem</span></span>|
|[<span data-ttu-id="f912e-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-433">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-434">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="f912e-435">Location : String | [Emplacement](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-436">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-437">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-437">Read mode</span></span>

<span data-ttu-id="f912e-438">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-439">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-439">Compose mode</span></span>

<span data-ttu-id="f912e-440">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f912e-441">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-441">Type</span></span>

*   <span data-ttu-id="f912e-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-443">Requirements</span></span>

|<span data-ttu-id="f912e-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-444">Requirement</span></span>| <span data-ttu-id="f912e-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-447">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-447">1.0</span></span>|
|[<span data-ttu-id="f912e-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-449">ReadItem</span></span>|
|[<span data-ttu-id="f912e-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-451">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-451">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f912e-452">normalizedSubject : chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-452">normalizedSubject: String</span></span>

<span data-ttu-id="f912e-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f912e-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="f912e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-457">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-457">Type</span></span>

*   <span data-ttu-id="f912e-458">String</span><span class="sxs-lookup"><span data-stu-id="f912e-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-459">Requirements</span></span>

|<span data-ttu-id="f912e-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-460">Requirement</span></span>| <span data-ttu-id="f912e-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-463">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-463">1.0</span></span>|
|[<span data-ttu-id="f912e-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-465">ReadItem</span></span>|
|[<span data-ttu-id="f912e-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-468">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="f912e-469">notificationMessages : [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-470">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-471">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-471">Type</span></span>

*   [<span data-ttu-id="f912e-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f912e-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-473">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-473">Requirements</span></span>

|<span data-ttu-id="f912e-474">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-474">Requirement</span></span>| <span data-ttu-id="f912e-475">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-476">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-477">1.3</span><span class="sxs-lookup"><span data-stu-id="f912e-477">1.3</span></span>|
|[<span data-ttu-id="f912e-478">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-479">ReadItem</span></span>|
|[<span data-ttu-id="f912e-480">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-481">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-482">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-482">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f912e-483">optionalAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="f912e-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-484">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="f912e-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f912e-485">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-486">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-486">Read mode</span></span>

<span data-ttu-id="f912e-487">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f912e-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-488">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-488">Compose mode</span></span>

<span data-ttu-id="f912e-489">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="f912e-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f912e-490">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-490">Type</span></span>

*   <span data-ttu-id="f912e-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-492">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-492">Requirements</span></span>

|<span data-ttu-id="f912e-493">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-493">Requirement</span></span>| <span data-ttu-id="f912e-494">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-495">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-496">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-496">1.0</span></span>|
|[<span data-ttu-id="f912e-497">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-498">ReadItem</span></span>|
|[<span data-ttu-id="f912e-499">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-500">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-500">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f912e-501">Organisateur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-504">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-504">Type</span></span>

*   [<span data-ttu-id="f912e-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f912e-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-506">Requirements</span></span>

|<span data-ttu-id="f912e-507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-507">Requirement</span></span>| <span data-ttu-id="f912e-508">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-510">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-510">1.0</span></span>|
|[<span data-ttu-id="f912e-511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-512">ReadItem</span></span>|
|[<span data-ttu-id="f912e-513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-514">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-515">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f912e-516">requiredAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="f912e-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-517">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="f912e-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f912e-518">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-519">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-519">Read mode</span></span>

<span data-ttu-id="f912e-520">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f912e-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-521">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-521">Compose mode</span></span>

<span data-ttu-id="f912e-522">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="f912e-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f912e-523">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-523">Type</span></span>

*   <span data-ttu-id="f912e-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-525">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-525">Requirements</span></span>

|<span data-ttu-id="f912e-526">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-526">Requirement</span></span>| <span data-ttu-id="f912e-527">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-528">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-529">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-529">1.0</span></span>|
|[<span data-ttu-id="f912e-530">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-531">ReadItem</span></span>|
|[<span data-ttu-id="f912e-532">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-533">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-533">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f912e-534">expéditeur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f912e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f912e-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f912e-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-539">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f912e-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f912e-540">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-540">Type</span></span>

*   [<span data-ttu-id="f912e-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f912e-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f912e-542">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-542">Requirements</span></span>

|<span data-ttu-id="f912e-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-543">Requirement</span></span>| <span data-ttu-id="f912e-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-546">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-546">1.0</span></span>|
|[<span data-ttu-id="f912e-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-548">ReadItem</span></span>|
|[<span data-ttu-id="f912e-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-550">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-551">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-551">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="f912e-552">début : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-553">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f912e-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f912e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-556">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-556">Read mode</span></span>

<span data-ttu-id="f912e-557">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f912e-557">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-558">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-558">Compose mode</span></span>

<span data-ttu-id="f912e-559">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f912e-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f912e-560">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f912e-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f912e-561">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f912e-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="f912e-562">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-562">Type</span></span>

*   <span data-ttu-id="f912e-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-564">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-564">Requirements</span></span>

|<span data-ttu-id="f912e-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-565">Requirement</span></span>| <span data-ttu-id="f912e-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-568">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-568">1.0</span></span>|
|[<span data-ttu-id="f912e-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-570">ReadItem</span></span>|
|[<span data-ttu-id="f912e-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-572">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="f912e-573">Subject : String | [Objet](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-574">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f912e-575">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="f912e-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-576">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-576">Read mode</span></span>

<span data-ttu-id="f912e-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f912e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-579">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-579">Compose mode</span></span>

<span data-ttu-id="f912e-580">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="f912e-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f912e-581">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-581">Type</span></span>

*   <span data-ttu-id="f912e-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-583">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-583">Requirements</span></span>

|<span data-ttu-id="f912e-584">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-584">Requirement</span></span>| <span data-ttu-id="f912e-585">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-586">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-587">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-587">1.0</span></span>|
|[<span data-ttu-id="f912e-588">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-589">ReadItem</span></span>|
|[<span data-ttu-id="f912e-590">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-591">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-591">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f912e-592">to : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f912e-593">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="f912e-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f912e-594">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f912e-595">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-595">Read mode</span></span>

<span data-ttu-id="f912e-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f912e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f912e-598">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f912e-598">Compose mode</span></span>

<span data-ttu-id="f912e-599">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="f912e-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f912e-600">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-600">Type</span></span>

*   <span data-ttu-id="f912e-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-602">Requirements</span></span>

|<span data-ttu-id="f912e-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-603">Requirement</span></span>| <span data-ttu-id="f912e-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-606">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-606">1.0</span></span>|
|[<span data-ttu-id="f912e-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-608">ReadItem</span></span>|
|[<span data-ttu-id="f912e-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-610">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f912e-611">Méthodes</span><span class="sxs-lookup"><span data-stu-id="f912e-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f912e-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f912e-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f912e-613">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f912e-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f912e-614">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="f912e-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f912e-615">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f912e-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-616">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-616">Parameters</span></span>

|<span data-ttu-id="f912e-617">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-617">Name</span></span>| <span data-ttu-id="f912e-618">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-618">Type</span></span>| <span data-ttu-id="f912e-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-619">Attributes</span></span>| <span data-ttu-id="f912e-620">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f912e-621">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-621">String</span></span>||<span data-ttu-id="f912e-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f912e-624">String</span><span class="sxs-lookup"><span data-stu-id="f912e-624">String</span></span>||<span data-ttu-id="f912e-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f912e-627">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-627">Object</span></span>| <span data-ttu-id="f912e-628">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-628">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-629">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="f912e-630">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-630">Object</span></span> | <span data-ttu-id="f912e-631">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-631">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-632">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="f912e-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="f912e-633">Boolean</span></span> | <span data-ttu-id="f912e-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-634">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-635">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f912e-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="f912e-636">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-636">function</span></span>| <span data-ttu-id="f912e-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-637">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-638">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f912e-639">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f912e-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f912e-640">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f912e-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f912e-641">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f912e-641">Errors</span></span>

| <span data-ttu-id="f912e-642">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f912e-642">Error code</span></span> | <span data-ttu-id="f912e-643">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f912e-644">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="f912e-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f912e-645">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="f912e-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f912e-646">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f912e-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-647">Requirements</span></span>

|<span data-ttu-id="f912e-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-648">Requirement</span></span>| <span data-ttu-id="f912e-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-651">1.1</span><span class="sxs-lookup"><span data-stu-id="f912e-651">1.1</span></span>|
|[<span data-ttu-id="f912e-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f912e-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="f912e-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-655">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f912e-656">Exemples</span><span class="sxs-lookup"><span data-stu-id="f912e-656">Examples</span></span>

```js
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

<span data-ttu-id="f912e-657">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f912e-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f912e-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f912e-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f912e-659">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f912e-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f912e-663">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f912e-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f912e-664">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez ; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="f912e-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-665">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-665">Parameters</span></span>

|<span data-ttu-id="f912e-666">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-666">Name</span></span>| <span data-ttu-id="f912e-667">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-667">Type</span></span>| <span data-ttu-id="f912e-668">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-668">Attributes</span></span>| <span data-ttu-id="f912e-669">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f912e-670">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-670">String</span></span>||<span data-ttu-id="f912e-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f912e-673">String</span><span class="sxs-lookup"><span data-stu-id="f912e-673">String</span></span>||<span data-ttu-id="f912e-674">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="f912e-674">The subject of the item to be attached.</span></span> <span data-ttu-id="f912e-675">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f912e-676">Object</span><span class="sxs-lookup"><span data-stu-id="f912e-676">Object</span></span>| <span data-ttu-id="f912e-677">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-677">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-678">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f912e-679">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-679">Object</span></span>| <span data-ttu-id="f912e-680">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-680">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-681">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f912e-682">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-682">function</span></span>| <span data-ttu-id="f912e-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-683">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-684">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f912e-685">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f912e-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f912e-686">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f912e-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f912e-687">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f912e-687">Errors</span></span>

| <span data-ttu-id="f912e-688">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f912e-688">Error code</span></span> | <span data-ttu-id="f912e-689">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f912e-690">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f912e-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-691">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-691">Requirements</span></span>

|<span data-ttu-id="f912e-692">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-692">Requirement</span></span>| <span data-ttu-id="f912e-693">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-694">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-695">1.1</span><span class="sxs-lookup"><span data-stu-id="f912e-695">1.1</span></span>|
|[<span data-ttu-id="f912e-696">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f912e-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="f912e-698">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-699">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-700">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-700">Example</span></span>

<span data-ttu-id="f912e-701">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f912e-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

<br>

---
---

#### <a name="close"></a><span data-ttu-id="f912e-702">close()</span><span class="sxs-lookup"><span data-stu-id="f912e-702">close()</span></span>

<span data-ttu-id="f912e-703">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="f912e-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f912e-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="f912e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-706">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f912e-707">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="f912e-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-708">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-708">Requirements</span></span>

|<span data-ttu-id="f912e-709">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-709">Requirement</span></span>| <span data-ttu-id="f912e-710">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-711">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-712">1.3</span><span class="sxs-lookup"><span data-stu-id="f912e-712">1.3</span></span>|
|[<span data-ttu-id="f912e-713">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-714">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f912e-714">Restricted</span></span>|
|[<span data-ttu-id="f912e-715">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-716">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-716">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f912e-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f912e-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f912e-718">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-719">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-720">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="f912e-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f912e-721">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f912e-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f912e-722">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="f912e-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="f912e-723">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="f912e-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="f912e-724">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f912e-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-725">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-725">Parameters</span></span>

| <span data-ttu-id="f912e-726">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-726">Name</span></span> | <span data-ttu-id="f912e-727">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-727">Type</span></span> | <span data-ttu-id="f912e-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-728">Attributes</span></span> | <span data-ttu-id="f912e-729">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f912e-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f912e-730">String &#124; Object</span></span>| |<span data-ttu-id="f912e-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f912e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f912e-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="f912e-733">**OR**</span></span><br/><span data-ttu-id="f912e-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f912e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f912e-736">String</span><span class="sxs-lookup"><span data-stu-id="f912e-736">String</span></span> | <span data-ttu-id="f912e-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-737">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f912e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f912e-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f912e-741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-741">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-742">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f912e-743">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-743">String</span></span> | | <span data-ttu-id="f912e-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f912e-746">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-746">String</span></span> | | <span data-ttu-id="f912e-747">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f912e-748">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-748">String</span></span> | | <span data-ttu-id="f912e-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f912e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f912e-751">Booléen</span><span class="sxs-lookup"><span data-stu-id="f912e-751">Boolean</span></span> | | <span data-ttu-id="f912e-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f912e-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f912e-754">String</span><span class="sxs-lookup"><span data-stu-id="f912e-754">String</span></span> | | <span data-ttu-id="f912e-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f912e-758">function</span><span class="sxs-lookup"><span data-stu-id="f912e-758">function</span></span> | <span data-ttu-id="f912e-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-759">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-760">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-761">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-761">Requirements</span></span>

|<span data-ttu-id="f912e-762">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-762">Requirement</span></span>| <span data-ttu-id="f912e-763">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-764">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-765">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-765">1.0</span></span>|
|[<span data-ttu-id="f912e-766">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-767">ReadItem</span></span>|
|[<span data-ttu-id="f912e-768">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-769">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f912e-770">Exemples</span><span class="sxs-lookup"><span data-stu-id="f912e-770">Examples</span></span>

<span data-ttu-id="f912e-771">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f912e-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f912e-772">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f912e-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f912e-773">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f912e-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f912e-774">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f912e-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f912e-775">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f912e-776">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f912e-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f912e-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f912e-778">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-779">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-780">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="f912e-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f912e-781">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f912e-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f912e-782">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="f912e-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="f912e-783">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="f912e-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="f912e-784">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f912e-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-785">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-785">Parameters</span></span>

| <span data-ttu-id="f912e-786">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-786">Name</span></span> | <span data-ttu-id="f912e-787">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-787">Type</span></span> | <span data-ttu-id="f912e-788">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-788">Attributes</span></span> | <span data-ttu-id="f912e-789">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f912e-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f912e-790">String &#124; Object</span></span>| | <span data-ttu-id="f912e-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f912e-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f912e-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="f912e-793">**OR**</span></span><br/><span data-ttu-id="f912e-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f912e-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f912e-796">String</span><span class="sxs-lookup"><span data-stu-id="f912e-796">String</span></span> | <span data-ttu-id="f912e-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-797">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f912e-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f912e-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f912e-801">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-801">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-802">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f912e-803">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-803">String</span></span> | | <span data-ttu-id="f912e-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f912e-806">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-806">String</span></span> | | <span data-ttu-id="f912e-807">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f912e-808">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-808">String</span></span> | | <span data-ttu-id="f912e-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f912e-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f912e-811">Booléen</span><span class="sxs-lookup"><span data-stu-id="f912e-811">Boolean</span></span> | | <span data-ttu-id="f912e-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f912e-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f912e-814">String</span><span class="sxs-lookup"><span data-stu-id="f912e-814">String</span></span> | | <span data-ttu-id="f912e-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f912e-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f912e-818">function</span><span class="sxs-lookup"><span data-stu-id="f912e-818">function</span></span> | <span data-ttu-id="f912e-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-819">&lt;optional&gt;</span></span> | <span data-ttu-id="f912e-820">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-821">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-821">Requirements</span></span>

|<span data-ttu-id="f912e-822">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-822">Requirement</span></span>| <span data-ttu-id="f912e-823">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-824">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-825">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-825">1.0</span></span>|
|[<span data-ttu-id="f912e-826">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-827">ReadItem</span></span>|
|[<span data-ttu-id="f912e-828">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-829">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f912e-830">Exemples</span><span class="sxs-lookup"><span data-stu-id="f912e-830">Examples</span></span>

<span data-ttu-id="f912e-831">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f912e-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f912e-832">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f912e-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f912e-833">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f912e-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f912e-834">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f912e-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f912e-835">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f912e-836">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="f912e-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="f912e-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="f912e-838">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-839">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-840">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-840">Requirements</span></span>

|<span data-ttu-id="f912e-841">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-841">Requirement</span></span>| <span data-ttu-id="f912e-842">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-843">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-844">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-844">1.0</span></span>|
|[<span data-ttu-id="f912e-845">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-846">ReadItem</span></span>|
|[<span data-ttu-id="f912e-847">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-848">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-849">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-849">Returns:</span></span>

<span data-ttu-id="f912e-850">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="f912e-851">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-851">Example</span></span>

<span data-ttu-id="f912e-852">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-852">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="f912e-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="f912e-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="f912e-854">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-855">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-856">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-856">Parameters</span></span>

|<span data-ttu-id="f912e-857">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-857">Name</span></span>| <span data-ttu-id="f912e-858">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-858">Type</span></span>| <span data-ttu-id="f912e-859">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f912e-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f912e-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="f912e-861">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="f912e-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-862">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-862">Requirements</span></span>

|<span data-ttu-id="f912e-863">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-863">Requirement</span></span>| <span data-ttu-id="f912e-864">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-865">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-866">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-866">1.0</span></span>|
|[<span data-ttu-id="f912e-867">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-868">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f912e-868">Restricted</span></span>|
|[<span data-ttu-id="f912e-869">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-870">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-871">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-871">Returns:</span></span>

<span data-ttu-id="f912e-872">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="f912e-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f912e-873">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f912e-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f912e-874">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f912e-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f912e-875">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="f912e-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f912e-876">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f912e-876">Value of `entityType`</span></span> | <span data-ttu-id="f912e-877">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="f912e-877">Type of objects in returned array</span></span> | <span data-ttu-id="f912e-878">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="f912e-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f912e-879">String</span><span class="sxs-lookup"><span data-stu-id="f912e-879">String</span></span> | <span data-ttu-id="f912e-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f912e-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f912e-881">Contact</span><span class="sxs-lookup"><span data-stu-id="f912e-881">Contact</span></span> | <span data-ttu-id="f912e-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f912e-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f912e-883">String</span><span class="sxs-lookup"><span data-stu-id="f912e-883">String</span></span> | <span data-ttu-id="f912e-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f912e-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f912e-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f912e-885">MeetingSuggestion</span></span> | <span data-ttu-id="f912e-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f912e-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f912e-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f912e-887">PhoneNumber</span></span> | <span data-ttu-id="f912e-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f912e-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f912e-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f912e-889">TaskSuggestion</span></span> | <span data-ttu-id="f912e-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f912e-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f912e-891">String</span><span class="sxs-lookup"><span data-stu-id="f912e-891">String</span></span> | <span data-ttu-id="f912e-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f912e-892">**Restricted**</span></span> |

<span data-ttu-id="f912e-893">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="f912e-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="f912e-894">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-894">Example</span></span>

<span data-ttu-id="f912e-895">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f912e-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="f912e-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="f912e-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="f912e-897">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f912e-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-898">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-899">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f912e-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-900">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-900">Parameters</span></span>

|<span data-ttu-id="f912e-901">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-901">Name</span></span>| <span data-ttu-id="f912e-902">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-902">Type</span></span>| <span data-ttu-id="f912e-903">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f912e-904">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-904">String</span></span>|<span data-ttu-id="f912e-905">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f912e-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-906">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-906">Requirements</span></span>

|<span data-ttu-id="f912e-907">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-907">Requirement</span></span>| <span data-ttu-id="f912e-908">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-909">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-910">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-910">1.0</span></span>|
|[<span data-ttu-id="f912e-911">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-912">ReadItem</span></span>|
|[<span data-ttu-id="f912e-913">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-914">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-915">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-915">Returns:</span></span>

<span data-ttu-id="f912e-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f912e-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f912e-918">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="f912e-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="f912e-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f912e-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f912e-920">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f912e-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-921">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f912e-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f912e-925">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="f912e-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f912e-926">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f912e-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f912e-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-930">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-930">Requirements</span></span>

|<span data-ttu-id="f912e-931">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-931">Requirement</span></span>| <span data-ttu-id="f912e-932">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-933">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-934">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-934">1.0</span></span>|
|[<span data-ttu-id="f912e-935">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-936">ReadItem</span></span>|
|[<span data-ttu-id="f912e-937">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-938">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-939">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-939">Returns:</span></span>

<span data-ttu-id="f912e-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="f912e-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="f912e-942">Type : objet</span><span class="sxs-lookup"><span data-stu-id="f912e-942">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="f912e-943">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-943">Example</span></span>

<span data-ttu-id="f912e-944">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f912e-944">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f912e-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f912e-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f912e-946">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f912e-946">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-947">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-948">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f912e-948">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f912e-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="f912e-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-951">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-951">Parameters</span></span>

|<span data-ttu-id="f912e-952">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-952">Name</span></span>| <span data-ttu-id="f912e-953">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-953">Type</span></span>| <span data-ttu-id="f912e-954">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-954">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f912e-955">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f912e-955">String</span></span>|<span data-ttu-id="f912e-956">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f912e-956">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-957">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-957">Requirements</span></span>

|<span data-ttu-id="f912e-958">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-958">Requirement</span></span>| <span data-ttu-id="f912e-959">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-959">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-960">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-960">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-961">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-961">1.0</span></span>|
|[<span data-ttu-id="f912e-962">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-962">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-963">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-963">ReadItem</span></span>|
|[<span data-ttu-id="f912e-964">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-964">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-965">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-965">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-966">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-966">Returns:</span></span>

<span data-ttu-id="f912e-967">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f912e-967">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="f912e-968">Type : Array. < String ></span><span class="sxs-lookup"><span data-stu-id="f912e-968">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="f912e-969">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-969">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f912e-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f912e-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f912e-971">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f912e-971">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f912e-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f912e-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-974">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-974">Parameters</span></span>

|<span data-ttu-id="f912e-975">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-975">Name</span></span>| <span data-ttu-id="f912e-976">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-976">Type</span></span>| <span data-ttu-id="f912e-977">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-977">Attributes</span></span>| <span data-ttu-id="f912e-978">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-978">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f912e-979">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f912e-979">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f912e-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="f912e-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f912e-983">Object</span><span class="sxs-lookup"><span data-stu-id="f912e-983">Object</span></span>| <span data-ttu-id="f912e-984">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-984">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-985">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-985">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f912e-986">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-986">Object</span></span>| <span data-ttu-id="f912e-987">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-987">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-988">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-988">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f912e-989">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-989">function</span></span>||<span data-ttu-id="f912e-990">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f912e-991">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f912e-991">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f912e-992">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f912e-992">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-993">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-993">Requirements</span></span>

|<span data-ttu-id="f912e-994">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-994">Requirement</span></span>| <span data-ttu-id="f912e-995">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-995">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-996">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-996">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-997">1.2</span><span class="sxs-lookup"><span data-stu-id="f912e-997">1.2</span></span>|
|[<span data-ttu-id="f912e-998">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-998">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-999">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-999">ReadItem</span></span>|
|[<span data-ttu-id="f912e-1000">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1000">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1001">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-1001">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-1002">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-1002">Returns:</span></span>

<span data-ttu-id="f912e-1003">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f912e-1003">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="f912e-1004">Type : String</span><span class="sxs-lookup"><span data-stu-id="f912e-1004">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f912e-1005">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1005">Example</span></span>

```js
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

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="f912e-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="f912e-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="f912e-1007">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-1007">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="f912e-1008">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f912e-1008">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-1009">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-1009">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-1010">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1010">Requirements</span></span>

|<span data-ttu-id="f912e-1011">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1011">Requirement</span></span>| <span data-ttu-id="f912e-1012">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1013">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1014">1.6</span><span class="sxs-lookup"><span data-stu-id="f912e-1014">1.6</span></span> |
|[<span data-ttu-id="f912e-1015">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1016">ReadItem</span></span>|
|[<span data-ttu-id="f912e-1017">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1018">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-1018">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-1019">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-1019">Returns:</span></span>

<span data-ttu-id="f912e-1020">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f912e-1020">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="f912e-1021">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1021">Example</span></span>

<span data-ttu-id="f912e-1022">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f912e-1022">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f912e-1023">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f912e-1023">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f912e-p164">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f912e-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-1026">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="f912e-1026">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f912e-p165">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f912e-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f912e-1030">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="f912e-1030">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f912e-1031">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f912e-1031">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f912e-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f912e-1035">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1035">Requirements</span></span>

|<span data-ttu-id="f912e-1036">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1036">Requirement</span></span>| <span data-ttu-id="f912e-1037">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1037">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1038">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1038">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1039">1.6</span><span class="sxs-lookup"><span data-stu-id="f912e-1039">1.6</span></span> |
|[<span data-ttu-id="f912e-1040">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1040">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1041">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1041">ReadItem</span></span>|
|[<span data-ttu-id="f912e-1042">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1042">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1043">Lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-1043">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f912e-1044">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f912e-1044">Returns:</span></span>

<span data-ttu-id="f912e-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="f912e-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f912e-1047">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1047">Example</span></span>

<span data-ttu-id="f912e-1048">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f912e-1048">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f912e-1049">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f912e-1049">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f912e-1050">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f912e-1050">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f912e-p168">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="f912e-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-1054">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-1054">Parameters</span></span>

|<span data-ttu-id="f912e-1055">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-1055">Name</span></span>| <span data-ttu-id="f912e-1056">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-1056">Type</span></span>| <span data-ttu-id="f912e-1057">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-1057">Attributes</span></span>| <span data-ttu-id="f912e-1058">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-1058">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f912e-1059">function</span><span class="sxs-lookup"><span data-stu-id="f912e-1059">function</span></span>||<span data-ttu-id="f912e-1060">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-1060">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f912e-1061">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f912e-1061">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f912e-1062">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f912e-1062">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f912e-1063">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1063">Object</span></span>| <span data-ttu-id="f912e-1064">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1065">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1065">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f912e-1066">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1066">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-1067">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1067">Requirements</span></span>

|<span data-ttu-id="f912e-1068">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1068">Requirement</span></span>| <span data-ttu-id="f912e-1069">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1069">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1070">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1070">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1071">1.0</span><span class="sxs-lookup"><span data-stu-id="f912e-1071">1.0</span></span>|
|[<span data-ttu-id="f912e-1072">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1072">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1073">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1073">ReadItem</span></span>|
|[<span data-ttu-id="f912e-1074">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1074">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1075">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f912e-1075">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-1076">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1076">Example</span></span>

<span data-ttu-id="f912e-p171">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f912e-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f912e-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f912e-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f912e-1081">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f912e-1081">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f912e-1082">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-1082">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f912e-1083">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="f912e-1083">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f912e-1084">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="f912e-1084">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f912e-1085">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="f912e-1085">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-1086">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-1086">Parameters</span></span>

|<span data-ttu-id="f912e-1087">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-1087">Name</span></span>| <span data-ttu-id="f912e-1088">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-1088">Type</span></span>| <span data-ttu-id="f912e-1089">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-1089">Attributes</span></span>| <span data-ttu-id="f912e-1090">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-1090">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f912e-1091">String</span><span class="sxs-lookup"><span data-stu-id="f912e-1091">String</span></span>||<span data-ttu-id="f912e-1092">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="f912e-1092">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="f912e-1093">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1093">Object</span></span>| <span data-ttu-id="f912e-1094">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1095">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f912e-1096">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1096">Object</span></span>| <span data-ttu-id="f912e-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1098">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f912e-1099">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-1099">function</span></span>| <span data-ttu-id="f912e-1100">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f912e-1102">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="f912e-1102">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f912e-1103">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f912e-1103">Errors</span></span>

| <span data-ttu-id="f912e-1104">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f912e-1104">Error code</span></span> | <span data-ttu-id="f912e-1105">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-1105">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f912e-1106">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="f912e-1106">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-1107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1107">Requirements</span></span>

|<span data-ttu-id="f912e-1108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1108">Requirement</span></span>| <span data-ttu-id="f912e-1109">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1111">1.1</span><span class="sxs-lookup"><span data-stu-id="f912e-1111">1.1</span></span>|
|[<span data-ttu-id="f912e-1112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1113">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1113">ReadWriteItem</span></span>|
|[<span data-ttu-id="f912e-1114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1115">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-1115">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-1116">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1116">Example</span></span>

<span data-ttu-id="f912e-1117">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="f912e-1117">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f912e-1118">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f912e-1118">saveAsync([options], callback)</span></span>

<span data-ttu-id="f912e-1119">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f912e-1119">Asynchronously saves an item.</span></span>

<span data-ttu-id="f912e-1120">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1120">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f912e-1121">Dans Outlook sur le Web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f912e-1121">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f912e-1122">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="f912e-1122">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-1123">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="f912e-1123">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f912e-1124">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="f912e-1124">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f912e-p175">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="f912e-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f912e-1128">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="f912e-1128">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f912e-1129">Outlook sur Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="f912e-1129">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f912e-1130">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f912e-1130">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f912e-1131">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="f912e-1131">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f912e-1132">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f912e-1132">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-1133">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-1133">Parameters</span></span>

|<span data-ttu-id="f912e-1134">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-1134">Name</span></span>| <span data-ttu-id="f912e-1135">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-1135">Type</span></span>| <span data-ttu-id="f912e-1136">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-1136">Attributes</span></span>| <span data-ttu-id="f912e-1137">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-1137">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f912e-1138">Object</span><span class="sxs-lookup"><span data-stu-id="f912e-1138">Object</span></span>| <span data-ttu-id="f912e-1139">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1139">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1140">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-1140">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f912e-1141">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1141">Object</span></span>| <span data-ttu-id="f912e-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1143">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1143">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f912e-1144">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-1144">function</span></span>||<span data-ttu-id="f912e-1145">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-1145">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f912e-1146">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f912e-1146">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f912e-1147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1147">Requirements</span></span>

|<span data-ttu-id="f912e-1148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1148">Requirement</span></span>| <span data-ttu-id="f912e-1149">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1149">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1151">1.3</span><span class="sxs-lookup"><span data-stu-id="f912e-1151">1.3</span></span>|
|[<span data-ttu-id="f912e-1152">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1153">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1153">ReadWriteItem</span></span>|
|[<span data-ttu-id="f912e-1154">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1155">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-1155">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f912e-1156">範例</span><span class="sxs-lookup"><span data-stu-id="f912e-1156">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f912e-p177">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f912e-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f912e-1159">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f912e-1159">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f912e-1160">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f912e-1160">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f912e-p178">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="f912e-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f912e-1164">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f912e-1164">Parameters</span></span>

|<span data-ttu-id="f912e-1165">Nom</span><span class="sxs-lookup"><span data-stu-id="f912e-1165">Name</span></span>| <span data-ttu-id="f912e-1166">Type</span><span class="sxs-lookup"><span data-stu-id="f912e-1166">Type</span></span>| <span data-ttu-id="f912e-1167">Attributs</span><span class="sxs-lookup"><span data-stu-id="f912e-1167">Attributes</span></span>| <span data-ttu-id="f912e-1168">Description</span><span class="sxs-lookup"><span data-stu-id="f912e-1168">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f912e-1169">String</span><span class="sxs-lookup"><span data-stu-id="f912e-1169">String</span></span>||<span data-ttu-id="f912e-p179">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="f912e-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f912e-1173">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1173">Object</span></span>| <span data-ttu-id="f912e-1174">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1174">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1175">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f912e-1175">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f912e-1176">Objet</span><span class="sxs-lookup"><span data-stu-id="f912e-1176">Object</span></span>| <span data-ttu-id="f912e-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1178">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f912e-1178">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f912e-1179">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f912e-1179">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f912e-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f912e-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="f912e-1181">Si `text`, le style actuel est appliqué dans Outlook sur le Web et les clients de bureau.</span><span class="sxs-lookup"><span data-stu-id="f912e-1181">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f912e-1182">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f912e-1182">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f912e-1183">Si `html` et que le champ prend en charge le format html (l’objet ne l’est pas), le style actuel est appliqué dans Outlook sur le Web et le style par défaut est appliqué dans les clients de bureau Outlook.</span><span class="sxs-lookup"><span data-stu-id="f912e-1183">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f912e-1184">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="f912e-1184">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f912e-1185">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="f912e-1185">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f912e-1186">fonction</span><span class="sxs-lookup"><span data-stu-id="f912e-1186">function</span></span>||<span data-ttu-id="f912e-1187">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f912e-1187">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f912e-1188">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f912e-1188">Requirements</span></span>

|<span data-ttu-id="f912e-1189">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f912e-1189">Requirement</span></span>| <span data-ttu-id="f912e-1190">Valeur</span><span class="sxs-lookup"><span data-stu-id="f912e-1190">Value</span></span>|
|---|---|
|[<span data-ttu-id="f912e-1191">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f912e-1191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f912e-1192">1.2</span><span class="sxs-lookup"><span data-stu-id="f912e-1192">1.2</span></span>|
|[<span data-ttu-id="f912e-1193">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f912e-1193">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f912e-1194">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f912e-1194">ReadWriteItem</span></span>|
|[<span data-ttu-id="f912e-1195">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f912e-1195">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f912e-1196">Composition</span><span class="sxs-lookup"><span data-stu-id="f912e-1196">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f912e-1197">Exemple</span><span class="sxs-lookup"><span data-stu-id="f912e-1197">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
