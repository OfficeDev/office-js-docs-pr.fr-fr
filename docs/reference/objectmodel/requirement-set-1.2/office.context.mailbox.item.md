---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,2
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: c765b0901c15adb7c3651ac279f224de05002023
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167346"
---
# <a name="item"></a><span data-ttu-id="0cc24-102">élément</span><span class="sxs-lookup"><span data-stu-id="0cc24-102">item</span></span>

### <span data-ttu-id="0cc24-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="0cc24-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="0cc24-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="0cc24-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-107">Requirements</span></span>

|<span data-ttu-id="0cc24-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-108">Requirement</span></span>| <span data-ttu-id="0cc24-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-111">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-111">1.0</span></span>|
|[<span data-ttu-id="0cc24-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0cc24-113">Restricted</span></span>|
|[<span data-ttu-id="0cc24-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0cc24-116">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="0cc24-116">Members and methods</span></span>

| <span data-ttu-id="0cc24-117">Membre	</span><span class="sxs-lookup"><span data-stu-id="0cc24-117">Member</span></span> | <span data-ttu-id="0cc24-118">Type	</span><span class="sxs-lookup"><span data-stu-id="0cc24-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0cc24-119">attachments</span><span class="sxs-lookup"><span data-stu-id="0cc24-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0cc24-120">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-120">Member</span></span> |
| [<span data-ttu-id="0cc24-121">bcc</span><span class="sxs-lookup"><span data-stu-id="0cc24-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0cc24-122">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-122">Member</span></span> |
| [<span data-ttu-id="0cc24-123">body</span><span class="sxs-lookup"><span data-stu-id="0cc24-123">body</span></span>](#body-body) | <span data-ttu-id="0cc24-124">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-124">Member</span></span> |
| [<span data-ttu-id="0cc24-125">cc</span><span class="sxs-lookup"><span data-stu-id="0cc24-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0cc24-126">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-126">Member</span></span> |
| [<span data-ttu-id="0cc24-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="0cc24-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0cc24-128">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-128">Member</span></span> |
| [<span data-ttu-id="0cc24-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0cc24-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0cc24-130">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-130">Member</span></span> |
| [<span data-ttu-id="0cc24-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0cc24-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0cc24-132">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-132">Member</span></span> |
| [<span data-ttu-id="0cc24-133">end</span><span class="sxs-lookup"><span data-stu-id="0cc24-133">end</span></span>](#end-datetime) | <span data-ttu-id="0cc24-134">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-134">Member</span></span> |
| [<span data-ttu-id="0cc24-135">from</span><span class="sxs-lookup"><span data-stu-id="0cc24-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="0cc24-136">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-136">Member</span></span> |
| [<span data-ttu-id="0cc24-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0cc24-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0cc24-138">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-138">Member</span></span> |
| [<span data-ttu-id="0cc24-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="0cc24-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0cc24-140">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-140">Member</span></span> |
| [<span data-ttu-id="0cc24-141">itemId</span><span class="sxs-lookup"><span data-stu-id="0cc24-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0cc24-142">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-142">Member</span></span> |
| [<span data-ttu-id="0cc24-143">itemType</span><span class="sxs-lookup"><span data-stu-id="0cc24-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0cc24-144">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-144">Member</span></span> |
| [<span data-ttu-id="0cc24-145">location</span><span class="sxs-lookup"><span data-stu-id="0cc24-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="0cc24-146">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-146">Member</span></span> |
| [<span data-ttu-id="0cc24-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0cc24-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0cc24-148">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-148">Member</span></span> |
| [<span data-ttu-id="0cc24-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0cc24-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0cc24-150">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-150">Member</span></span> |
| [<span data-ttu-id="0cc24-151">organizer</span><span class="sxs-lookup"><span data-stu-id="0cc24-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="0cc24-152">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-152">Member</span></span> |
| [<span data-ttu-id="0cc24-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0cc24-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0cc24-154">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-154">Member</span></span> |
| [<span data-ttu-id="0cc24-155">sender</span><span class="sxs-lookup"><span data-stu-id="0cc24-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0cc24-156">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-156">Member</span></span> |
| [<span data-ttu-id="0cc24-157">start</span><span class="sxs-lookup"><span data-stu-id="0cc24-157">start</span></span>](#start-datetime) | <span data-ttu-id="0cc24-158">Member</span><span class="sxs-lookup"><span data-stu-id="0cc24-158">Member</span></span> |
| [<span data-ttu-id="0cc24-159">subject</span><span class="sxs-lookup"><span data-stu-id="0cc24-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0cc24-160">Membre</span><span class="sxs-lookup"><span data-stu-id="0cc24-160">Member</span></span> |
| [<span data-ttu-id="0cc24-161">to</span><span class="sxs-lookup"><span data-stu-id="0cc24-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0cc24-162">Membre</span><span class="sxs-lookup"><span data-stu-id="0cc24-162">Member</span></span> |
| [<span data-ttu-id="0cc24-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0cc24-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-164">Method</span></span> |
| [<span data-ttu-id="0cc24-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0cc24-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-166">Method</span></span> |
| [<span data-ttu-id="0cc24-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0cc24-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0cc24-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-168">Method</span></span> |
| [<span data-ttu-id="0cc24-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0cc24-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0cc24-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-170">Method</span></span> |
| [<span data-ttu-id="0cc24-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="0cc24-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0cc24-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-172">Method</span></span> |
| [<span data-ttu-id="0cc24-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0cc24-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0cc24-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-174">Method</span></span> |
| [<span data-ttu-id="0cc24-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0cc24-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0cc24-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-176">Method</span></span> |
| [<span data-ttu-id="0cc24-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0cc24-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0cc24-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-178">Method</span></span> |
| [<span data-ttu-id="0cc24-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0cc24-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0cc24-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-180">Method</span></span> |
| [<span data-ttu-id="0cc24-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0cc24-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-182">Method</span></span> |
| [<span data-ttu-id="0cc24-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0cc24-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-184">Method</span></span> |
| [<span data-ttu-id="0cc24-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0cc24-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-186">Method</span></span> |
| [<span data-ttu-id="0cc24-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0cc24-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0cc24-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="0cc24-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0cc24-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-189">Example</span></span>

<span data-ttu-id="0cc24-190">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="0cc24-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0cc24-191">Membres</span><span class="sxs-lookup"><span data-stu-id="0cc24-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="0cc24-192">pièces jointes : tableau. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="0cc24-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="0cc24-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-195">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="0cc24-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0cc24-196">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="0cc24-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-197">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-197">Type</span></span>

*   <span data-ttu-id="0cc24-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="0cc24-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-199">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-199">Requirements</span></span>

|<span data-ttu-id="0cc24-200">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-200">Requirement</span></span>| <span data-ttu-id="0cc24-201">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-202">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-203">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-203">1.0</span></span>|
|[<span data-ttu-id="0cc24-204">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-205">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-206">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-207">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-208">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-208">Example</span></span>

<span data-ttu-id="0cc24-209">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0cc24-210">CCI : [destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-211">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0cc24-212">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-213">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-213">Type</span></span>

*   [<span data-ttu-id="0cc24-214">Destinataires</span><span class="sxs-lookup"><span data-stu-id="0cc24-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-215">Requirements</span></span>

|<span data-ttu-id="0cc24-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-216">Requirement</span></span>| <span data-ttu-id="0cc24-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-219">1.1</span><span class="sxs-lookup"><span data-stu-id="0cc24-219">1.1</span></span>|
|[<span data-ttu-id="0cc24-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-221">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-223">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="0cc24-225">Body : [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-226">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-227">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-227">Type</span></span>

*   [<span data-ttu-id="0cc24-228">Body</span><span class="sxs-lookup"><span data-stu-id="0cc24-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-229">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-229">Requirements</span></span>

|<span data-ttu-id="0cc24-230">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-230">Requirement</span></span>| <span data-ttu-id="0cc24-231">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-232">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-233">1.1</span><span class="sxs-lookup"><span data-stu-id="0cc24-233">1.1</span></span>|
|[<span data-ttu-id="0cc24-234">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-235">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-238">Example</span></span>

<span data-ttu-id="0cc24-239">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="0cc24-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0cc24-240">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0cc24-241">CC : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-242">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0cc24-243">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-244">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-244">Read mode</span></span>

<span data-ttu-id="0cc24-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-247">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-247">Compose mode</span></span>

<span data-ttu-id="0cc24-248">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-249">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-249">Type</span></span>

*   <span data-ttu-id="0cc24-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-251">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-251">Requirements</span></span>

|<span data-ttu-id="0cc24-252">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-252">Requirement</span></span>| <span data-ttu-id="0cc24-253">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-254">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-255">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-255">1.0</span></span>|
|[<span data-ttu-id="0cc24-256">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-257">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-258">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-259">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-259">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0cc24-260">(Nullable) conversationId : chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="0cc24-261">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="0cc24-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0cc24-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0cc24-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-266">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-266">Type</span></span>

*   <span data-ttu-id="0cc24-267">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-268">Requirements</span></span>

|<span data-ttu-id="0cc24-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-269">Requirement</span></span>| <span data-ttu-id="0cc24-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-272">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-272">1.0</span></span>|
|[<span data-ttu-id="0cc24-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-274">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-277">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0cc24-278">dateTimeCreated : date</span><span class="sxs-lookup"><span data-stu-id="0cc24-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="0cc24-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-281">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-281">Type</span></span>

*   <span data-ttu-id="0cc24-282">Date</span><span class="sxs-lookup"><span data-stu-id="0cc24-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-283">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-283">Requirements</span></span>

|<span data-ttu-id="0cc24-284">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-284">Requirement</span></span>| <span data-ttu-id="0cc24-285">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-286">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-287">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-287">1.0</span></span>|
|[<span data-ttu-id="0cc24-288">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-289">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-290">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-291">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-292">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-292">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0cc24-293">dateTimeModified : date</span><span class="sxs-lookup"><span data-stu-id="0cc24-293">dateTimeModified: Date</span></span>

<span data-ttu-id="0cc24-294">Obtient la date et l’heure de la dernière modification d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="0cc24-295">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-296">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-297">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-297">Type</span></span>

*   <span data-ttu-id="0cc24-298">Date</span><span class="sxs-lookup"><span data-stu-id="0cc24-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-299">Requirements</span></span>

|<span data-ttu-id="0cc24-300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-300">Requirement</span></span>| <span data-ttu-id="0cc24-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-303">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-303">1.0</span></span>|
|[<span data-ttu-id="0cc24-304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-305">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-307">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-308">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="0cc24-309">fin : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-310">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0cc24-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-313">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-313">Read mode</span></span>

<span data-ttu-id="0cc24-314">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-314">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-315">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-315">Compose mode</span></span>

<span data-ttu-id="0cc24-316">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0cc24-317">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0cc24-318">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0cc24-319">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-319">Type</span></span>

*   <span data-ttu-id="0cc24-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-321">Requirements</span></span>

|<span data-ttu-id="0cc24-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-322">Requirement</span></span>| <span data-ttu-id="0cc24-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-325">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-325">1.0</span></span>|
|[<span data-ttu-id="0cc24-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-327">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-329">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-329">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0cc24-330">de : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="0cc24-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-335">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-336">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-336">Type</span></span>

*   [<span data-ttu-id="0cc24-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0cc24-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-338">Requirements</span></span>

|<span data-ttu-id="0cc24-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-339">Requirement</span></span>| <span data-ttu-id="0cc24-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-342">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-342">1.0</span></span>|
|[<span data-ttu-id="0cc24-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-344">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-346">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="0cc24-348">internetMessageId : chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-348">internetMessageId: String</span></span>

<span data-ttu-id="0cc24-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-351">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-351">Type</span></span>

*   <span data-ttu-id="0cc24-352">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-353">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-353">Requirements</span></span>

|<span data-ttu-id="0cc24-354">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-354">Requirement</span></span>| <span data-ttu-id="0cc24-355">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-356">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-357">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-357">1.0</span></span>|
|[<span data-ttu-id="0cc24-358">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-359">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-360">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-361">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-362">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-362">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0cc24-363">itemClass : chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-363">itemClass: String</span></span>

<span data-ttu-id="0cc24-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0cc24-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="0cc24-368">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-368">Type</span></span> | <span data-ttu-id="0cc24-369">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-369">Description</span></span> | <span data-ttu-id="0cc24-370">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="0cc24-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="0cc24-371">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="0cc24-371">Appointment items</span></span> | <span data-ttu-id="0cc24-372">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="0cc24-373">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="0cc24-373">Message items</span></span> | <span data-ttu-id="0cc24-374">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="0cc24-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="0cc24-375">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-376">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-376">Type</span></span>

*   <span data-ttu-id="0cc24-377">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-378">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-378">Requirements</span></span>

|<span data-ttu-id="0cc24-379">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-379">Requirement</span></span>| <span data-ttu-id="0cc24-380">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-381">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-382">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-382">1.0</span></span>|
|[<span data-ttu-id="0cc24-383">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-384">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-385">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-386">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-387">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-387">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0cc24-388">(Nullable) itemId : String</span><span class="sxs-lookup"><span data-stu-id="0cc24-388">(nullable) itemId: String</span></span>

<span data-ttu-id="0cc24-389">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="0cc24-390">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-391">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="0cc24-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0cc24-392">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0cc24-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0cc24-393">Avant d’effectuer des appels d’API REST à l’aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l’aide de, qui est disponible à partir de l’ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="0cc24-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="0cc24-394">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="0cc24-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-395">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-395">Type</span></span>

*   <span data-ttu-id="0cc24-396">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-397">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-397">Requirements</span></span>

|<span data-ttu-id="0cc24-398">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-398">Requirement</span></span>| <span data-ttu-id="0cc24-399">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-400">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-401">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-401">1.0</span></span>|
|[<span data-ttu-id="0cc24-402">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-403">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-404">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-405">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-406">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-406">Example</span></span>

<span data-ttu-id="0cc24-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="0cc24-409">itemType : [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-410">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="0cc24-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0cc24-411">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-412">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-412">Type</span></span>

*   [<span data-ttu-id="0cc24-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0cc24-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-414">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-414">Requirements</span></span>

|<span data-ttu-id="0cc24-415">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-415">Requirement</span></span>| <span data-ttu-id="0cc24-416">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-417">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-418">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-418">1.0</span></span>|
|[<span data-ttu-id="0cc24-419">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-420">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-421">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-422">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-423">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-423">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="0cc24-424">Location : String | [Emplacement](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-425">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-426">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-426">Read mode</span></span>

<span data-ttu-id="0cc24-427">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-428">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-428">Compose mode</span></span>

<span data-ttu-id="0cc24-429">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-430">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-430">Type</span></span>

*   <span data-ttu-id="0cc24-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-432">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-432">Requirements</span></span>

|<span data-ttu-id="0cc24-433">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-433">Requirement</span></span>| <span data-ttu-id="0cc24-434">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-435">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-436">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-436">1.0</span></span>|
|[<span data-ttu-id="0cc24-437">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-438">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-439">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-440">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-440">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0cc24-441">normalizedSubject : chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-441">normalizedSubject: String</span></span>

<span data-ttu-id="0cc24-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0cc24-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="0cc24-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-446">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-446">Type</span></span>

*   <span data-ttu-id="0cc24-447">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-448">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-448">Requirements</span></span>

|<span data-ttu-id="0cc24-449">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-449">Requirement</span></span>| <span data-ttu-id="0cc24-450">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-451">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-452">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-452">1.0</span></span>|
|[<span data-ttu-id="0cc24-453">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-454">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-455">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-456">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-457">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-457">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0cc24-458">optionalAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="0cc24-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-459">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0cc24-460">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-461">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-461">Read mode</span></span>

<span data-ttu-id="0cc24-462">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0cc24-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-463">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-463">Compose mode</span></span>

<span data-ttu-id="0cc24-464">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="0cc24-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-465">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-465">Type</span></span>

*   <span data-ttu-id="0cc24-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-467">Requirements</span></span>

|<span data-ttu-id="0cc24-468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-468">Requirement</span></span>| <span data-ttu-id="0cc24-469">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-471">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-471">1.0</span></span>|
|[<span data-ttu-id="0cc24-472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-473">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-475">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-475">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0cc24-476">Organisateur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-479">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-479">Type</span></span>

*   [<span data-ttu-id="0cc24-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0cc24-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-481">Requirements</span></span>

|<span data-ttu-id="0cc24-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-482">Requirement</span></span>| <span data-ttu-id="0cc24-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-485">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-485">1.0</span></span>|
|[<span data-ttu-id="0cc24-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-487">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-489">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-490">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-490">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0cc24-491">requiredAttendees : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|des[destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) de tableau. <</span><span class="sxs-lookup"><span data-stu-id="0cc24-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-492">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0cc24-493">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-494">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-494">Read mode</span></span>

<span data-ttu-id="0cc24-495">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="0cc24-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-496">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-496">Compose mode</span></span>

<span data-ttu-id="0cc24-497">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="0cc24-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-498">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-498">Type</span></span>

*   <span data-ttu-id="0cc24-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-500">Requirements</span></span>

|<span data-ttu-id="0cc24-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-501">Requirement</span></span>| <span data-ttu-id="0cc24-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-504">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-504">1.0</span></span>|
|[<span data-ttu-id="0cc24-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-506">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-508">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-508">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0cc24-509">expéditeur : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0cc24-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-514">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0cc24-515">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-515">Type</span></span>

*   [<span data-ttu-id="0cc24-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0cc24-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0cc24-517">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-517">Requirements</span></span>

|<span data-ttu-id="0cc24-518">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-518">Requirement</span></span>| <span data-ttu-id="0cc24-519">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-520">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-521">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-521">1.0</span></span>|
|[<span data-ttu-id="0cc24-522">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-523">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-524">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-525">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-526">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-526">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="0cc24-527">début : date | [Fois](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-528">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0cc24-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-531">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-531">Read mode</span></span>

<span data-ttu-id="0cc24-532">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-532">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-533">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-533">Compose mode</span></span>

<span data-ttu-id="0cc24-534">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0cc24-535">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="0cc24-536">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0cc24-537">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-537">Type</span></span>

*   <span data-ttu-id="0cc24-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-539">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-539">Requirements</span></span>

|<span data-ttu-id="0cc24-540">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-540">Requirement</span></span>| <span data-ttu-id="0cc24-541">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-542">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-543">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-543">1.0</span></span>|
|[<span data-ttu-id="0cc24-544">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-545">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-546">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-547">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-547">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="0cc24-548">Subject : String | [Objet](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-549">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0cc24-550">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="0cc24-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-551">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-551">Read mode</span></span>

<span data-ttu-id="0cc24-p130">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-554">Compose mode</span></span>

<span data-ttu-id="0cc24-555">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="0cc24-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-556">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-556">Type</span></span>

*   <span data-ttu-id="0cc24-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-558">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-558">Requirements</span></span>

|<span data-ttu-id="0cc24-559">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-559">Requirement</span></span>| <span data-ttu-id="0cc24-560">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-561">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-562">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-562">1.0</span></span>|
|[<span data-ttu-id="0cc24-563">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-564">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-565">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-566">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-566">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0cc24-567">to : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0cc24-568">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0cc24-569">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0cc24-570">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-570">Read mode</span></span>

<span data-ttu-id="0cc24-p132">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0cc24-573">Mode composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-573">Compose mode</span></span>

<span data-ttu-id="0cc24-574">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0cc24-575">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-575">Type</span></span>

*   <span data-ttu-id="0cc24-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-577">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-577">Requirements</span></span>

|<span data-ttu-id="0cc24-578">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-578">Requirement</span></span>| <span data-ttu-id="0cc24-579">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-580">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-581">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-581">1.0</span></span>|
|[<span data-ttu-id="0cc24-582">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-583">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-584">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-585">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0cc24-586">Méthodes</span><span class="sxs-lookup"><span data-stu-id="0cc24-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0cc24-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0cc24-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0cc24-588">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="0cc24-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0cc24-589">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="0cc24-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0cc24-590">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0cc24-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-591">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-591">Parameters</span></span>

|<span data-ttu-id="0cc24-592">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-592">Name</span></span>| <span data-ttu-id="0cc24-593">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-593">Type</span></span>| <span data-ttu-id="0cc24-594">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-594">Attributes</span></span>| <span data-ttu-id="0cc24-595">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="0cc24-596">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-596">String</span></span>||<span data-ttu-id="0cc24-p133">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0cc24-599">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-599">String</span></span>||<span data-ttu-id="0cc24-p134">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0cc24-602">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-602">Object</span></span>| <span data-ttu-id="0cc24-603">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-603">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-604">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0cc24-605">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-605">Object</span></span>| <span data-ttu-id="0cc24-606">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-606">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-607">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0cc24-608">fonction</span><span class="sxs-lookup"><span data-stu-id="0cc24-608">function</span></span>| <span data-ttu-id="0cc24-609">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-609">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-610">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0cc24-611">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0cc24-612">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0cc24-613">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0cc24-613">Errors</span></span>

| <span data-ttu-id="0cc24-614">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0cc24-614">Error code</span></span> | <span data-ttu-id="0cc24-615">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="0cc24-616">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="0cc24-617">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0cc24-618">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-619">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-619">Requirements</span></span>

|<span data-ttu-id="0cc24-620">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-620">Requirement</span></span>| <span data-ttu-id="0cc24-621">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-622">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-623">1.1</span><span class="sxs-lookup"><span data-stu-id="0cc24-623">1.1</span></span>|
|[<span data-ttu-id="0cc24-624">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="0cc24-626">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-627">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-628">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-628">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0cc24-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0cc24-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0cc24-630">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0cc24-p135">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0cc24-634">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="0cc24-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0cc24-635">Si votre complément Office est en cours d’exécution dans Outlook sur le Web, `addItemAttachmentAsync` la méthode peut joindre des éléments à des éléments autres que l’élément que vous modifiez ; Toutefois, cette option n’est pas prise en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-636">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-636">Parameters</span></span>

|<span data-ttu-id="0cc24-637">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-637">Name</span></span>| <span data-ttu-id="0cc24-638">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-638">Type</span></span>| <span data-ttu-id="0cc24-639">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-639">Attributes</span></span>| <span data-ttu-id="0cc24-640">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="0cc24-641">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-641">String</span></span>||<span data-ttu-id="0cc24-p136">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0cc24-644">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-644">String</span></span>||<span data-ttu-id="0cc24-645">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="0cc24-645">The subject of the item to be attached.</span></span> <span data-ttu-id="0cc24-646">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0cc24-647">Object</span><span class="sxs-lookup"><span data-stu-id="0cc24-647">Object</span></span>| <span data-ttu-id="0cc24-648">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-648">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-649">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0cc24-650">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-650">Object</span></span>| <span data-ttu-id="0cc24-651">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-651">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-652">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0cc24-653">fonction</span><span class="sxs-lookup"><span data-stu-id="0cc24-653">function</span></span>| <span data-ttu-id="0cc24-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-654">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-655">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0cc24-656">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0cc24-657">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0cc24-658">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0cc24-658">Errors</span></span>

| <span data-ttu-id="0cc24-659">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0cc24-659">Error code</span></span> | <span data-ttu-id="0cc24-660">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0cc24-661">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-662">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-662">Requirements</span></span>

|<span data-ttu-id="0cc24-663">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-663">Requirement</span></span>| <span data-ttu-id="0cc24-664">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-665">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-666">1.1</span><span class="sxs-lookup"><span data-stu-id="0cc24-666">1.1</span></span>|
|[<span data-ttu-id="0cc24-667">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="0cc24-669">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-670">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-671">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-671">Example</span></span>

<span data-ttu-id="0cc24-672">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0cc24-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0cc24-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0cc24-674">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0cc24-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-675">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0cc24-676">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="0cc24-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0cc24-677">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0cc24-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0cc24-678">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="0cc24-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0cc24-679">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="0cc24-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0cc24-680">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0cc24-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-681">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-681">Parameters</span></span>

|<span data-ttu-id="0cc24-682">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-682">Name</span></span>| <span data-ttu-id="0cc24-683">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-683">Type</span></span>| <span data-ttu-id="0cc24-684">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0cc24-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0cc24-685">String &#124; Object</span></span>| |<span data-ttu-id="0cc24-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0cc24-688">**OU**</span><span class="sxs-lookup"><span data-stu-id="0cc24-688">**OR**</span></span><br/><span data-ttu-id="0cc24-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0cc24-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0cc24-691">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-691">String</span></span> | <span data-ttu-id="0cc24-692">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-692">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0cc24-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0cc24-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-696">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-697">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0cc24-698">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-698">String</span></span> | | <span data-ttu-id="0cc24-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0cc24-701">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-701">String</span></span> | | <span data-ttu-id="0cc24-702">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0cc24-703">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-703">String</span></span> | | <span data-ttu-id="0cc24-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0cc24-706">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-706">String</span></span> | | <span data-ttu-id="0cc24-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0cc24-710">function</span><span class="sxs-lookup"><span data-stu-id="0cc24-710">function</span></span> | <span data-ttu-id="0cc24-711">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-711">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-712">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-713">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-713">Requirements</span></span>

|<span data-ttu-id="0cc24-714">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-714">Requirement</span></span>| <span data-ttu-id="0cc24-715">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-716">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-717">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-717">1.0</span></span>|
|[<span data-ttu-id="0cc24-718">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-719">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-720">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-721">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0cc24-722">Exemples</span><span class="sxs-lookup"><span data-stu-id="0cc24-722">Examples</span></span>

<span data-ttu-id="0cc24-723">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0cc24-724">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0cc24-724">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0cc24-725">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0cc24-725">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0cc24-726">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0cc24-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0cc24-727">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0cc24-728">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0cc24-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0cc24-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0cc24-730">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0cc24-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-731">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0cc24-732">Dans Outlook sur le Web, le formulaire de réponse s’affiche sous la forme d’un formulaire indépendant dans un affichage à 3 colonnes et sous forme de formulaire contextuel en affichage 2 ou 1 colonne.</span><span class="sxs-lookup"><span data-stu-id="0cc24-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0cc24-733">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="0cc24-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0cc24-734">Lorsque des pièces jointes sont `formData.attachments` spécifiées dans le paramètre, Outlook sur le Web et les clients de bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="0cc24-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0cc24-735">Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire.</span><span class="sxs-lookup"><span data-stu-id="0cc24-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0cc24-736">Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="0cc24-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-737">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-737">Parameters</span></span>

|<span data-ttu-id="0cc24-738">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-738">Name</span></span>| <span data-ttu-id="0cc24-739">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-739">Type</span></span>| <span data-ttu-id="0cc24-740">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0cc24-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0cc24-741">String &#124; Object</span></span>| | <span data-ttu-id="0cc24-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0cc24-744">**OU**</span><span class="sxs-lookup"><span data-stu-id="0cc24-744">**OR**</span></span><br/><span data-ttu-id="0cc24-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="0cc24-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0cc24-747">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-747">String</span></span> | <span data-ttu-id="0cc24-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-748">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0cc24-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0cc24-752">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-752">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-753">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0cc24-754">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-754">String</span></span> | | <span data-ttu-id="0cc24-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0cc24-757">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-757">String</span></span> | | <span data-ttu-id="0cc24-758">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0cc24-759">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-759">String</span></span> | | <span data-ttu-id="0cc24-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0cc24-762">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-762">String</span></span> | | <span data-ttu-id="0cc24-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0cc24-766">function</span><span class="sxs-lookup"><span data-stu-id="0cc24-766">function</span></span> | <span data-ttu-id="0cc24-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-767">&lt;optional&gt;</span></span> | <span data-ttu-id="0cc24-768">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-769">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-769">Requirements</span></span>

|<span data-ttu-id="0cc24-770">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-770">Requirement</span></span>| <span data-ttu-id="0cc24-771">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-772">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-773">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-773">1.0</span></span>|
|[<span data-ttu-id="0cc24-774">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-775">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-776">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-777">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0cc24-778">Exemples</span><span class="sxs-lookup"><span data-stu-id="0cc24-778">Examples</span></span>

<span data-ttu-id="0cc24-779">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0cc24-780">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="0cc24-780">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0cc24-781">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="0cc24-781">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0cc24-782">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="0cc24-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0cc24-783">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0cc24-784">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="0cc24-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="0cc24-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="0cc24-786">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0cc24-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-787">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-788">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-788">Requirements</span></span>

|<span data-ttu-id="0cc24-789">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-789">Requirement</span></span>| <span data-ttu-id="0cc24-790">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-791">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-792">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-792">1.0</span></span>|
|[<span data-ttu-id="0cc24-793">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-794">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-795">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-796">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-797">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-797">Returns:</span></span>

<span data-ttu-id="0cc24-798">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0cc24-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="0cc24-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-799">Example</span></span>

<span data-ttu-id="0cc24-800">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-800">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="0cc24-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="0cc24-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="0cc24-802">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0cc24-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-803">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-804">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-804">Parameters</span></span>

|<span data-ttu-id="0cc24-805">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-805">Name</span></span>| <span data-ttu-id="0cc24-806">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-806">Type</span></span>| <span data-ttu-id="0cc24-807">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="0cc24-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0cc24-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="0cc24-809">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="0cc24-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0cc24-810">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-810">Requirements</span></span>

|<span data-ttu-id="0cc24-811">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-811">Requirement</span></span>| <span data-ttu-id="0cc24-812">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-813">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-814">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-814">1.0</span></span>|
|[<span data-ttu-id="0cc24-815">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-816">Restreinte</span><span class="sxs-lookup"><span data-stu-id="0cc24-816">Restricted</span></span>|
|[<span data-ttu-id="0cc24-817">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-818">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-819">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-819">Returns:</span></span>

<span data-ttu-id="0cc24-820">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="0cc24-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0cc24-821">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0cc24-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0cc24-822">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0cc24-823">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="0cc24-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="0cc24-824">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="0cc24-824">Value of `entityType`</span></span> | <span data-ttu-id="0cc24-825">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="0cc24-825">Type of objects in returned array</span></span> | <span data-ttu-id="0cc24-826">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="0cc24-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="0cc24-827">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-827">String</span></span> | <span data-ttu-id="0cc24-828">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0cc24-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="0cc24-829">Contact</span><span class="sxs-lookup"><span data-stu-id="0cc24-829">Contact</span></span> | <span data-ttu-id="0cc24-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0cc24-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="0cc24-831">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-831">String</span></span> | <span data-ttu-id="0cc24-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0cc24-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="0cc24-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0cc24-833">MeetingSuggestion</span></span> | <span data-ttu-id="0cc24-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0cc24-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="0cc24-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0cc24-835">PhoneNumber</span></span> | <span data-ttu-id="0cc24-836">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0cc24-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="0cc24-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0cc24-837">TaskSuggestion</span></span> | <span data-ttu-id="0cc24-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0cc24-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="0cc24-839">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-839">String</span></span> | <span data-ttu-id="0cc24-840">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0cc24-840">**Restricted**</span></span> |

<span data-ttu-id="0cc24-841">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="0cc24-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="0cc24-842">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-842">Example</span></span>

<span data-ttu-id="0cc24-843">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="0cc24-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="0cc24-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="0cc24-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="0cc24-845">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-846">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0cc24-847">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0cc24-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-848">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-848">Parameters</span></span>

|<span data-ttu-id="0cc24-849">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-849">Name</span></span>| <span data-ttu-id="0cc24-850">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-850">Type</span></span>| <span data-ttu-id="0cc24-851">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0cc24-852">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-852">String</span></span>|<span data-ttu-id="0cc24-853">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0cc24-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0cc24-854">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-854">Requirements</span></span>

|<span data-ttu-id="0cc24-855">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-855">Requirement</span></span>| <span data-ttu-id="0cc24-856">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-857">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-858">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-858">1.0</span></span>|
|[<span data-ttu-id="0cc24-859">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-860">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-861">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-862">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-863">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-863">Returns:</span></span>

<span data-ttu-id="0cc24-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0cc24-866">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="0cc24-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0cc24-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0cc24-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0cc24-868">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-869">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0cc24-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0cc24-873">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="0cc24-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0cc24-874">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="0cc24-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cc24-877">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-877">Requirements</span></span>

|<span data-ttu-id="0cc24-878">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-878">Requirement</span></span>| <span data-ttu-id="0cc24-879">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-880">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-881">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-881">1.0</span></span>|
|[<span data-ttu-id="0cc24-882">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-883">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-884">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-885">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-886">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-886">Returns:</span></span>

<span data-ttu-id="0cc24-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="0cc24-889">Type : objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-889">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="0cc24-890">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-890">Example</span></span>

<span data-ttu-id="0cc24-891">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="0cc24-891">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0cc24-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0cc24-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0cc24-893">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-893">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc24-894">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="0cc24-894">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0cc24-895">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="0cc24-895">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0cc24-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-898">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-898">Parameters</span></span>

|<span data-ttu-id="0cc24-899">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-899">Name</span></span>| <span data-ttu-id="0cc24-900">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-900">Type</span></span>| <span data-ttu-id="0cc24-901">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-901">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0cc24-902">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0cc24-902">String</span></span>|<span data-ttu-id="0cc24-903">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="0cc24-903">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0cc24-904">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-904">Requirements</span></span>

|<span data-ttu-id="0cc24-905">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-905">Requirement</span></span>| <span data-ttu-id="0cc24-906">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-906">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-907">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-907">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-908">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-908">1.0</span></span>|
|[<span data-ttu-id="0cc24-909">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-909">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-910">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-910">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-911">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-911">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-912">Lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-912">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-913">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-913">Returns:</span></span>

<span data-ttu-id="0cc24-914">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-914">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="0cc24-915">Type : Array. < String ></span><span class="sxs-lookup"><span data-stu-id="0cc24-915">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="0cc24-916">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-916">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0cc24-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0cc24-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0cc24-918">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="0cc24-918">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0cc24-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-921">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-921">Parameters</span></span>

|<span data-ttu-id="0cc24-922">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-922">Name</span></span>| <span data-ttu-id="0cc24-923">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-923">Type</span></span>| <span data-ttu-id="0cc24-924">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-924">Attributes</span></span>| <span data-ttu-id="0cc24-925">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-925">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="0cc24-926">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0cc24-926">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0cc24-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="0cc24-930">Object</span><span class="sxs-lookup"><span data-stu-id="0cc24-930">Object</span></span>| <span data-ttu-id="0cc24-931">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-931">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-932">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-932">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0cc24-933">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-933">Object</span></span>| <span data-ttu-id="0cc24-934">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-934">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-935">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-935">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0cc24-936">fonction</span><span class="sxs-lookup"><span data-stu-id="0cc24-936">function</span></span>||<span data-ttu-id="0cc24-937">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0cc24-938">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-938">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0cc24-939">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-939">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0cc24-940">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-940">Requirements</span></span>

|<span data-ttu-id="0cc24-941">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-941">Requirement</span></span>| <span data-ttu-id="0cc24-942">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-943">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-944">1.2</span><span class="sxs-lookup"><span data-stu-id="0cc24-944">1.2</span></span>|
|[<span data-ttu-id="0cc24-945">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-946">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-947">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-948">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-948">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0cc24-949">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="0cc24-949">Returns:</span></span>

<span data-ttu-id="0cc24-950">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-950">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="0cc24-951">Type : String</span><span class="sxs-lookup"><span data-stu-id="0cc24-951">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0cc24-952">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-952">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0cc24-953">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0cc24-953">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0cc24-954">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0cc24-954">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0cc24-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-958">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-958">Parameters</span></span>

|<span data-ttu-id="0cc24-959">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-959">Name</span></span>| <span data-ttu-id="0cc24-960">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-960">Type</span></span>| <span data-ttu-id="0cc24-961">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-961">Attributes</span></span>| <span data-ttu-id="0cc24-962">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-962">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0cc24-963">function</span><span class="sxs-lookup"><span data-stu-id="0cc24-963">function</span></span>||<span data-ttu-id="0cc24-964">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-964">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0cc24-965">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0cc24-965">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0cc24-966">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0cc24-966">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="0cc24-967">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-967">Object</span></span>| <span data-ttu-id="0cc24-968">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-968">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-969">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-969">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0cc24-970">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-970">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0cc24-971">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-971">Requirements</span></span>

|<span data-ttu-id="0cc24-972">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-972">Requirement</span></span>| <span data-ttu-id="0cc24-973">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-973">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-974">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-974">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-975">1.0</span><span class="sxs-lookup"><span data-stu-id="0cc24-975">1.0</span></span>|
|[<span data-ttu-id="0cc24-976">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-976">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-977">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-977">ReadItem</span></span>|
|[<span data-ttu-id="0cc24-978">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-978">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-979">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0cc24-979">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-980">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-980">Example</span></span>

<span data-ttu-id="0cc24-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0cc24-984">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0cc24-984">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0cc24-985">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="0cc24-985">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0cc24-986">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="0cc24-986">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0cc24-987">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="0cc24-987">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0cc24-988">Dans Outlook sur le Web et les appareils mobiles, l’identificateur de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="0cc24-988">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0cc24-989">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-989">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-990">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-990">Parameters</span></span>

|<span data-ttu-id="0cc24-991">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-991">Name</span></span>| <span data-ttu-id="0cc24-992">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-992">Type</span></span>| <span data-ttu-id="0cc24-993">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-993">Attributes</span></span>| <span data-ttu-id="0cc24-994">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-994">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="0cc24-995">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-995">String</span></span>||<span data-ttu-id="0cc24-996">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="0cc24-996">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="0cc24-997">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-997">Object</span></span>| <span data-ttu-id="0cc24-998">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-998">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-999">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-999">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0cc24-1000">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-1000">Object</span></span>| <span data-ttu-id="0cc24-1001">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-1002">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1002">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0cc24-1003">fonction</span><span class="sxs-lookup"><span data-stu-id="0cc24-1003">function</span></span>| <span data-ttu-id="0cc24-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-1005">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-1005">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0cc24-1006">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1006">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0cc24-1007">Erreurs</span><span class="sxs-lookup"><span data-stu-id="0cc24-1007">Errors</span></span>

| <span data-ttu-id="0cc24-1008">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="0cc24-1008">Error code</span></span> | <span data-ttu-id="0cc24-1009">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-1009">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="0cc24-1010">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1010">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-1011">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-1011">Requirements</span></span>

|<span data-ttu-id="0cc24-1012">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-1012">Requirement</span></span>| <span data-ttu-id="0cc24-1013">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-1014">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-1015">1.1</span><span class="sxs-lookup"><span data-stu-id="0cc24-1015">1.1</span></span>|
|[<span data-ttu-id="0cc24-1016">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="0cc24-1018">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-1019">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-1019">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-1020">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-1020">Example</span></span>

<span data-ttu-id="0cc24-1021">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="0cc24-1021">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0cc24-1022">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0cc24-1022">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0cc24-1023">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1023">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0cc24-p166">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0cc24-1027">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0cc24-1027">Parameters</span></span>

|<span data-ttu-id="0cc24-1028">Nom</span><span class="sxs-lookup"><span data-stu-id="0cc24-1028">Name</span></span>| <span data-ttu-id="0cc24-1029">Type</span><span class="sxs-lookup"><span data-stu-id="0cc24-1029">Type</span></span>| <span data-ttu-id="0cc24-1030">Attributs</span><span class="sxs-lookup"><span data-stu-id="0cc24-1030">Attributes</span></span>| <span data-ttu-id="0cc24-1031">Description</span><span class="sxs-lookup"><span data-stu-id="0cc24-1031">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0cc24-1032">String</span><span class="sxs-lookup"><span data-stu-id="0cc24-1032">String</span></span>||<span data-ttu-id="0cc24-p167">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="0cc24-1036">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-1036">Object</span></span>| <span data-ttu-id="0cc24-1037">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-1038">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1038">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0cc24-1039">Objet</span><span class="sxs-lookup"><span data-stu-id="0cc24-1039">Object</span></span>| <span data-ttu-id="0cc24-1040">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-1041">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1041">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0cc24-1042">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0cc24-1042">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0cc24-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0cc24-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="0cc24-1044">Si `text`, le style actuel est appliqué dans Outlook sur le Web et les clients de bureau.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1044">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="0cc24-1045">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1045">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0cc24-1046">Si `html` et que le champ prend en charge le format html (l’objet ne l’est pas), le style actuel est appliqué dans Outlook sur le Web et le style par défaut est appliqué dans les clients de bureau Outlook.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1046">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="0cc24-1047">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1047">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0cc24-1048">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="0cc24-1048">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="0cc24-1049">fonction</span><span class="sxs-lookup"><span data-stu-id="0cc24-1049">function</span></span>||<span data-ttu-id="0cc24-1050">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0cc24-1050">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0cc24-1051">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0cc24-1051">Requirements</span></span>

|<span data-ttu-id="0cc24-1052">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0cc24-1052">Requirement</span></span>| <span data-ttu-id="0cc24-1053">Valeur</span><span class="sxs-lookup"><span data-stu-id="0cc24-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cc24-1054">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0cc24-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cc24-1055">1.2</span><span class="sxs-lookup"><span data-stu-id="0cc24-1055">1.2</span></span>|
|[<span data-ttu-id="0cc24-1056">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="0cc24-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cc24-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0cc24-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="0cc24-1058">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0cc24-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cc24-1059">Composition</span><span class="sxs-lookup"><span data-stu-id="0cc24-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0cc24-1060">Exemple</span><span class="sxs-lookup"><span data-stu-id="0cc24-1060">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
