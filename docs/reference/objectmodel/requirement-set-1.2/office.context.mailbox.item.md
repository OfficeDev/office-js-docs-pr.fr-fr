---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,2
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 97fa271f500e89c6ce69d82b95a0818f6d5bc7d4
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001606"
---
# <a name="item"></a><span data-ttu-id="04352-102">élément</span><span class="sxs-lookup"><span data-stu-id="04352-102">item</span></span>

### <span data-ttu-id="04352-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="04352-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="04352-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="04352-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-107">Requirements</span></span>

|<span data-ttu-id="04352-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-108">Requirement</span></span>| <span data-ttu-id="04352-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-111">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-111">1.0</span></span>|
|[<span data-ttu-id="04352-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="04352-113">Restricted</span></span>|
|[<span data-ttu-id="04352-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="04352-116">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="04352-116">Members and methods</span></span>

| <span data-ttu-id="04352-117">Membre	</span><span class="sxs-lookup"><span data-stu-id="04352-117">Member</span></span> | <span data-ttu-id="04352-118">Type	</span><span class="sxs-lookup"><span data-stu-id="04352-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="04352-119">attachments</span><span class="sxs-lookup"><span data-stu-id="04352-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="04352-120">Member</span><span class="sxs-lookup"><span data-stu-id="04352-120">Member</span></span> |
| [<span data-ttu-id="04352-121">bcc</span><span class="sxs-lookup"><span data-stu-id="04352-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="04352-122">Member</span><span class="sxs-lookup"><span data-stu-id="04352-122">Member</span></span> |
| [<span data-ttu-id="04352-123">body</span><span class="sxs-lookup"><span data-stu-id="04352-123">body</span></span>](#body-body) | <span data-ttu-id="04352-124">Member</span><span class="sxs-lookup"><span data-stu-id="04352-124">Member</span></span> |
| [<span data-ttu-id="04352-125">cc</span><span class="sxs-lookup"><span data-stu-id="04352-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="04352-126">Member</span><span class="sxs-lookup"><span data-stu-id="04352-126">Member</span></span> |
| [<span data-ttu-id="04352-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="04352-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="04352-128">Member</span><span class="sxs-lookup"><span data-stu-id="04352-128">Member</span></span> |
| [<span data-ttu-id="04352-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="04352-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="04352-130">Member</span><span class="sxs-lookup"><span data-stu-id="04352-130">Member</span></span> |
| [<span data-ttu-id="04352-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="04352-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="04352-132">Member</span><span class="sxs-lookup"><span data-stu-id="04352-132">Member</span></span> |
| [<span data-ttu-id="04352-133">end</span><span class="sxs-lookup"><span data-stu-id="04352-133">end</span></span>](#end-datetime) | <span data-ttu-id="04352-134">Member</span><span class="sxs-lookup"><span data-stu-id="04352-134">Member</span></span> |
| [<span data-ttu-id="04352-135">from</span><span class="sxs-lookup"><span data-stu-id="04352-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="04352-136">Member</span><span class="sxs-lookup"><span data-stu-id="04352-136">Member</span></span> |
| [<span data-ttu-id="04352-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="04352-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="04352-138">Member</span><span class="sxs-lookup"><span data-stu-id="04352-138">Member</span></span> |
| [<span data-ttu-id="04352-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="04352-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="04352-140">Member</span><span class="sxs-lookup"><span data-stu-id="04352-140">Member</span></span> |
| [<span data-ttu-id="04352-141">itemId</span><span class="sxs-lookup"><span data-stu-id="04352-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="04352-142">Member</span><span class="sxs-lookup"><span data-stu-id="04352-142">Member</span></span> |
| [<span data-ttu-id="04352-143">itemType</span><span class="sxs-lookup"><span data-stu-id="04352-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="04352-144">Member</span><span class="sxs-lookup"><span data-stu-id="04352-144">Member</span></span> |
| [<span data-ttu-id="04352-145">location</span><span class="sxs-lookup"><span data-stu-id="04352-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="04352-146">Member</span><span class="sxs-lookup"><span data-stu-id="04352-146">Member</span></span> |
| [<span data-ttu-id="04352-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="04352-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="04352-148">Member</span><span class="sxs-lookup"><span data-stu-id="04352-148">Member</span></span> |
| [<span data-ttu-id="04352-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="04352-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="04352-150">Member</span><span class="sxs-lookup"><span data-stu-id="04352-150">Member</span></span> |
| [<span data-ttu-id="04352-151">organizer</span><span class="sxs-lookup"><span data-stu-id="04352-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="04352-152">Member</span><span class="sxs-lookup"><span data-stu-id="04352-152">Member</span></span> |
| [<span data-ttu-id="04352-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="04352-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="04352-154">Member</span><span class="sxs-lookup"><span data-stu-id="04352-154">Member</span></span> |
| [<span data-ttu-id="04352-155">sender</span><span class="sxs-lookup"><span data-stu-id="04352-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="04352-156">Member</span><span class="sxs-lookup"><span data-stu-id="04352-156">Member</span></span> |
| [<span data-ttu-id="04352-157">start</span><span class="sxs-lookup"><span data-stu-id="04352-157">start</span></span>](#start-datetime) | <span data-ttu-id="04352-158">Member</span><span class="sxs-lookup"><span data-stu-id="04352-158">Member</span></span> |
| [<span data-ttu-id="04352-159">subject</span><span class="sxs-lookup"><span data-stu-id="04352-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="04352-160">Membre</span><span class="sxs-lookup"><span data-stu-id="04352-160">Member</span></span> |
| [<span data-ttu-id="04352-161">to</span><span class="sxs-lookup"><span data-stu-id="04352-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="04352-162">Membre</span><span class="sxs-lookup"><span data-stu-id="04352-162">Member</span></span> |
| [<span data-ttu-id="04352-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="04352-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="04352-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-164">Method</span></span> |
| [<span data-ttu-id="04352-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="04352-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="04352-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-166">Method</span></span> |
| [<span data-ttu-id="04352-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="04352-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="04352-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-168">Method</span></span> |
| [<span data-ttu-id="04352-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="04352-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="04352-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-170">Method</span></span> |
| [<span data-ttu-id="04352-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="04352-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="04352-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-172">Method</span></span> |
| [<span data-ttu-id="04352-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="04352-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="04352-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-174">Method</span></span> |
| [<span data-ttu-id="04352-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="04352-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="04352-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-176">Method</span></span> |
| [<span data-ttu-id="04352-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="04352-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="04352-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-178">Method</span></span> |
| [<span data-ttu-id="04352-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="04352-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="04352-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-180">Method</span></span> |
| [<span data-ttu-id="04352-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="04352-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="04352-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-182">Method</span></span> |
| [<span data-ttu-id="04352-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="04352-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="04352-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-184">Method</span></span> |
| [<span data-ttu-id="04352-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="04352-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="04352-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-186">Method</span></span> |
| [<span data-ttu-id="04352-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="04352-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="04352-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="04352-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="04352-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-189">Example</span></span>

<span data-ttu-id="04352-190">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="04352-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="04352-191">Members</span><span class="sxs-lookup"><span data-stu-id="04352-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="04352-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="04352-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="04352-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-195">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="04352-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="04352-196">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="04352-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="04352-197">Type</span><span class="sxs-lookup"><span data-stu-id="04352-197">Type</span></span>

*   <span data-ttu-id="04352-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="04352-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-199">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-199">Requirements</span></span>

|<span data-ttu-id="04352-200">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-200">Requirement</span></span>| <span data-ttu-id="04352-201">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-202">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-203">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-203">1.0</span></span>|
|[<span data-ttu-id="04352-204">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-205">ReadItem</span></span>|
|[<span data-ttu-id="04352-206">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-207">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-208">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-208">Example</span></span>

<span data-ttu-id="04352-209">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="04352-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-211">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="04352-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="04352-212">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-212">Compose mode only.</span></span>

<span data-ttu-id="04352-213">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-214">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="04352-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="04352-215">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="04352-216">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="04352-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-217">Type</span><span class="sxs-lookup"><span data-stu-id="04352-217">Type</span></span>

*   [<span data-ttu-id="04352-218">Destinataires</span><span class="sxs-lookup"><span data-stu-id="04352-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-219">Requirements</span></span>

|<span data-ttu-id="04352-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-220">Requirement</span></span>| <span data-ttu-id="04352-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-223">1.1</span><span class="sxs-lookup"><span data-stu-id="04352-223">1.1</span></span>|
|[<span data-ttu-id="04352-224">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-225">ReadItem</span></span>|
|[<span data-ttu-id="04352-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-227">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-228">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="04352-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-230">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="04352-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-231">Type</span><span class="sxs-lookup"><span data-stu-id="04352-231">Type</span></span>

*   [<span data-ttu-id="04352-232">Body</span><span class="sxs-lookup"><span data-stu-id="04352-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-233">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-233">Requirements</span></span>

|<span data-ttu-id="04352-234">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-234">Requirement</span></span>| <span data-ttu-id="04352-235">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-236">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-237">1.1</span><span class="sxs-lookup"><span data-stu-id="04352-237">1.1</span></span>|
|[<span data-ttu-id="04352-238">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-239">ReadItem</span></span>|
|[<span data-ttu-id="04352-240">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-241">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-242">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-242">Example</span></span>

<span data-ttu-id="04352-243">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="04352-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="04352-244">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="04352-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-246">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="04352-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="04352-247">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-248">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-248">Read mode</span></span>

<span data-ttu-id="04352-249">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="04352-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="04352-250">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-251">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-252">Compose mode</span></span>

<span data-ttu-id="04352-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="04352-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="04352-254">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-255">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="04352-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="04352-256">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="04352-257">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="04352-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04352-258">Type</span><span class="sxs-lookup"><span data-stu-id="04352-258">Type</span></span>

*   <span data-ttu-id="04352-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-260">Requirements</span></span>

|<span data-ttu-id="04352-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-261">Requirement</span></span>| <span data-ttu-id="04352-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-264">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-264">1.0</span></span>|
|[<span data-ttu-id="04352-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-266">ReadItem</span></span>|
|[<span data-ttu-id="04352-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="04352-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="04352-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="04352-270">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="04352-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="04352-p110">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="04352-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="04352-p111">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="04352-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-275">Type</span><span class="sxs-lookup"><span data-stu-id="04352-275">Type</span></span>

*   <span data-ttu-id="04352-276">String</span><span class="sxs-lookup"><span data-stu-id="04352-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-277">Requirements</span></span>

|<span data-ttu-id="04352-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-278">Requirement</span></span>| <span data-ttu-id="04352-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-281">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-281">1.0</span></span>|
|[<span data-ttu-id="04352-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-283">ReadItem</span></span>|
|[<span data-ttu-id="04352-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="04352-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="04352-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="04352-p112">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-290">Type</span><span class="sxs-lookup"><span data-stu-id="04352-290">Type</span></span>

*   <span data-ttu-id="04352-291">Date</span><span class="sxs-lookup"><span data-stu-id="04352-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-292">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-292">Requirements</span></span>

|<span data-ttu-id="04352-293">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-293">Requirement</span></span>| <span data-ttu-id="04352-294">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-295">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-296">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-296">1.0</span></span>|
|[<span data-ttu-id="04352-297">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-298">ReadItem</span></span>|
|[<span data-ttu-id="04352-299">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-300">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-301">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="04352-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="04352-302">dateTimeModified: Date</span></span>

<span data-ttu-id="04352-p113">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-305">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-306">Type</span><span class="sxs-lookup"><span data-stu-id="04352-306">Type</span></span>

*   <span data-ttu-id="04352-307">Date</span><span class="sxs-lookup"><span data-stu-id="04352-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-308">Requirements</span></span>

|<span data-ttu-id="04352-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-309">Requirement</span></span>| <span data-ttu-id="04352-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-312">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-312">1.0</span></span>|
|[<span data-ttu-id="04352-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-314">ReadItem</span></span>|
|[<span data-ttu-id="04352-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-316">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="04352-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-319">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="04352-p114">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="04352-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-322">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-322">Read mode</span></span>

<span data-ttu-id="04352-323">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="04352-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-324">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-324">Compose mode</span></span>

<span data-ttu-id="04352-325">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="04352-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="04352-326">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="04352-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="04352-327">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="04352-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="04352-328">Type</span><span class="sxs-lookup"><span data-stu-id="04352-328">Type</span></span>

*   <span data-ttu-id="04352-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-330">Requirements</span></span>

|<span data-ttu-id="04352-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-331">Requirement</span></span>| <span data-ttu-id="04352-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-334">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-334">1.0</span></span>|
|[<span data-ttu-id="04352-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-336">ReadItem</span></span>|
|[<span data-ttu-id="04352-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-338">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="04352-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-p115">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="04352-p116">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="04352-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-344">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="04352-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-345">Type</span><span class="sxs-lookup"><span data-stu-id="04352-345">Type</span></span>

*   [<span data-ttu-id="04352-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04352-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-347">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-347">Requirements</span></span>

|<span data-ttu-id="04352-348">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-348">Requirement</span></span>| <span data-ttu-id="04352-349">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-350">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-351">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-351">1.0</span></span>|
|[<span data-ttu-id="04352-352">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-353">ReadItem</span></span>|
|[<span data-ttu-id="04352-354">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-355">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-356">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="04352-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="04352-357">internetMessageId: String</span></span>

<span data-ttu-id="04352-p117">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-360">Type</span><span class="sxs-lookup"><span data-stu-id="04352-360">Type</span></span>

*   <span data-ttu-id="04352-361">String</span><span class="sxs-lookup"><span data-stu-id="04352-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-362">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-362">Requirements</span></span>

|<span data-ttu-id="04352-363">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-363">Requirement</span></span>| <span data-ttu-id="04352-364">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-365">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-366">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-366">1.0</span></span>|
|[<span data-ttu-id="04352-367">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-368">ReadItem</span></span>|
|[<span data-ttu-id="04352-369">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-370">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-371">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="04352-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="04352-372">itemClass: String</span></span>

<span data-ttu-id="04352-p118">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="04352-p119">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="04352-377">Type</span><span class="sxs-lookup"><span data-stu-id="04352-377">Type</span></span> | <span data-ttu-id="04352-378">Description</span><span class="sxs-lookup"><span data-stu-id="04352-378">Description</span></span> | <span data-ttu-id="04352-379">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="04352-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="04352-380">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="04352-380">Appointment items</span></span> | <span data-ttu-id="04352-381">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="04352-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="04352-382">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="04352-382">Message items</span></span> | <span data-ttu-id="04352-383">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="04352-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="04352-384">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="04352-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-385">Type</span><span class="sxs-lookup"><span data-stu-id="04352-385">Type</span></span>

*   <span data-ttu-id="04352-386">String</span><span class="sxs-lookup"><span data-stu-id="04352-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-387">Requirements</span></span>

|<span data-ttu-id="04352-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-388">Requirement</span></span>| <span data-ttu-id="04352-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-391">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-391">1.0</span></span>|
|[<span data-ttu-id="04352-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-393">ReadItem</span></span>|
|[<span data-ttu-id="04352-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-395">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-396">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="04352-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="04352-397">(nullable) itemId: String</span></span>

<span data-ttu-id="04352-398">Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="04352-398">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="04352-399">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-399">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-400">L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="04352-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="04352-401">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="04352-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="04352-402">Avant d’effectuer des appels d’API REST à l’aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l’aide de, qui est disponible à partir de l’ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="04352-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="04352-403">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="04352-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="04352-404">Type</span><span class="sxs-lookup"><span data-stu-id="04352-404">Type</span></span>

*   <span data-ttu-id="04352-405">String</span><span class="sxs-lookup"><span data-stu-id="04352-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-406">Requirements</span></span>

|<span data-ttu-id="04352-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-407">Requirement</span></span>| <span data-ttu-id="04352-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-410">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-410">1.0</span></span>|
|[<span data-ttu-id="04352-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-412">ReadItem</span></span>|
|[<span data-ttu-id="04352-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-415">Example</span></span>

<span data-ttu-id="04352-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="04352-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="04352-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-419">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="04352-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="04352-420">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-421">Type</span><span class="sxs-lookup"><span data-stu-id="04352-421">Type</span></span>

*   [<span data-ttu-id="04352-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="04352-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-423">Requirements</span></span>

|<span data-ttu-id="04352-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-424">Requirement</span></span>| <span data-ttu-id="04352-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-427">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-427">1.0</span></span>|
|[<span data-ttu-id="04352-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-429">ReadItem</span></span>|
|[<span data-ttu-id="04352-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-431">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="04352-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-434">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-435">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-435">Read mode</span></span>

<span data-ttu-id="04352-436">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-437">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-437">Compose mode</span></span>

<span data-ttu-id="04352-438">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04352-439">Type</span><span class="sxs-lookup"><span data-stu-id="04352-439">Type</span></span>

*   <span data-ttu-id="04352-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-441">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-441">Requirements</span></span>

|<span data-ttu-id="04352-442">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-442">Requirement</span></span>| <span data-ttu-id="04352-443">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-444">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-445">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-445">1.0</span></span>|
|[<span data-ttu-id="04352-446">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-447">ReadItem</span></span>|
|[<span data-ttu-id="04352-448">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-449">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="04352-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="04352-450">normalizedSubject: String</span></span>

<span data-ttu-id="04352-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="04352-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="04352-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-455">Type</span><span class="sxs-lookup"><span data-stu-id="04352-455">Type</span></span>

*   <span data-ttu-id="04352-456">String</span><span class="sxs-lookup"><span data-stu-id="04352-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-457">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-457">Requirements</span></span>

|<span data-ttu-id="04352-458">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-458">Requirement</span></span>| <span data-ttu-id="04352-459">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-460">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-461">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-461">1.0</span></span>|
|[<span data-ttu-id="04352-462">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-463">ReadItem</span></span>|
|[<span data-ttu-id="04352-464">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-465">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-466">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="04352-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-468">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="04352-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="04352-469">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-470">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-470">Read mode</span></span>

<span data-ttu-id="04352-471">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="04352-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="04352-472">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-473">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-474">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-474">Compose mode</span></span>

<span data-ttu-id="04352-475">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="04352-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="04352-476">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-477">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="04352-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="04352-478">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="04352-479">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="04352-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04352-480">Type</span><span class="sxs-lookup"><span data-stu-id="04352-480">Type</span></span>

*   <span data-ttu-id="04352-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-482">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-482">Requirements</span></span>

|<span data-ttu-id="04352-483">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-483">Requirement</span></span>| <span data-ttu-id="04352-484">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-485">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-486">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-486">1.0</span></span>|
|[<span data-ttu-id="04352-487">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-488">ReadItem</span></span>|
|[<span data-ttu-id="04352-489">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-490">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="04352-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-p128">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-494">Type</span><span class="sxs-lookup"><span data-stu-id="04352-494">Type</span></span>

*   [<span data-ttu-id="04352-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04352-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-496">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-496">Requirements</span></span>

|<span data-ttu-id="04352-497">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-497">Requirement</span></span>| <span data-ttu-id="04352-498">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-499">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-500">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-500">1.0</span></span>|
|[<span data-ttu-id="04352-501">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-502">ReadItem</span></span>|
|[<span data-ttu-id="04352-503">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-504">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-505">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="04352-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-507">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="04352-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="04352-508">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-509">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-509">Read mode</span></span>

<span data-ttu-id="04352-510">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="04352-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="04352-511">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-512">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-513">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-513">Compose mode</span></span>

<span data-ttu-id="04352-514">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="04352-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="04352-515">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-516">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="04352-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="04352-517">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="04352-518">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="04352-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="04352-519">Type</span><span class="sxs-lookup"><span data-stu-id="04352-519">Type</span></span>

*   <span data-ttu-id="04352-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-521">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-521">Requirements</span></span>

|<span data-ttu-id="04352-522">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-522">Requirement</span></span>| <span data-ttu-id="04352-523">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-524">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-525">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-525">1.0</span></span>|
|[<span data-ttu-id="04352-526">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-527">ReadItem</span></span>|
|[<span data-ttu-id="04352-528">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-529">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="04352-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-p132">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="04352-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="04352-p133">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="04352-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-535">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="04352-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="04352-536">Type</span><span class="sxs-lookup"><span data-stu-id="04352-536">Type</span></span>

*   [<span data-ttu-id="04352-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04352-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="04352-538">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-538">Requirements</span></span>

|<span data-ttu-id="04352-539">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-539">Requirement</span></span>| <span data-ttu-id="04352-540">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-541">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-542">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-542">1.0</span></span>|
|[<span data-ttu-id="04352-543">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-544">ReadItem</span></span>|
|[<span data-ttu-id="04352-545">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-546">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-547">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="04352-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-549">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="04352-p134">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="04352-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-552">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-552">Read mode</span></span>

<span data-ttu-id="04352-553">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="04352-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-554">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-554">Compose mode</span></span>

<span data-ttu-id="04352-555">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="04352-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="04352-556">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="04352-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="04352-557">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="04352-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="04352-558">Type</span><span class="sxs-lookup"><span data-stu-id="04352-558">Type</span></span>

*   <span data-ttu-id="04352-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-560">Requirements</span></span>

|<span data-ttu-id="04352-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-561">Requirement</span></span>| <span data-ttu-id="04352-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-564">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-564">1.0</span></span>|
|[<span data-ttu-id="04352-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-566">ReadItem</span></span>|
|[<span data-ttu-id="04352-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="04352-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-570">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="04352-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="04352-571">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="04352-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-572">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-572">Read mode</span></span>

<span data-ttu-id="04352-p136">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="04352-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-575">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-575">Compose mode</span></span>

<span data-ttu-id="04352-576">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="04352-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="04352-577">Type</span><span class="sxs-lookup"><span data-stu-id="04352-577">Type</span></span>

*   <span data-ttu-id="04352-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-579">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-579">Requirements</span></span>

|<span data-ttu-id="04352-580">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-580">Requirement</span></span>| <span data-ttu-id="04352-581">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-582">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-583">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-583">1.0</span></span>|
|[<span data-ttu-id="04352-584">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-585">ReadItem</span></span>|
|[<span data-ttu-id="04352-586">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-587">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="04352-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="04352-589">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="04352-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="04352-590">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04352-591">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="04352-591">Read mode</span></span>

<span data-ttu-id="04352-592">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="04352-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="04352-593">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-594">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="04352-595">Mode composition</span><span class="sxs-lookup"><span data-stu-id="04352-595">Compose mode</span></span>

<span data-ttu-id="04352-596">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="04352-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="04352-597">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="04352-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="04352-598">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="04352-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="04352-599">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="04352-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="04352-600">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="04352-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04352-601">Type</span><span class="sxs-lookup"><span data-stu-id="04352-601">Type</span></span>

*   <span data-ttu-id="04352-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-603">Requirements</span></span>

|<span data-ttu-id="04352-604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-604">Requirement</span></span>| <span data-ttu-id="04352-605">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-607">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-607">1.0</span></span>|
|[<span data-ttu-id="04352-608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-609">ReadItem</span></span>|
|[<span data-ttu-id="04352-610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-611">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="04352-612">Méthodes</span><span class="sxs-lookup"><span data-stu-id="04352-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="04352-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04352-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="04352-614">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="04352-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="04352-615">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="04352-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="04352-616">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="04352-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-617">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-617">Parameters</span></span>

|<span data-ttu-id="04352-618">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-618">Name</span></span>| <span data-ttu-id="04352-619">Type</span><span class="sxs-lookup"><span data-stu-id="04352-619">Type</span></span>| <span data-ttu-id="04352-620">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-620">Attributes</span></span>| <span data-ttu-id="04352-621">Description</span><span class="sxs-lookup"><span data-stu-id="04352-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="04352-622">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-622">String</span></span>||<span data-ttu-id="04352-p140">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="04352-625">String</span><span class="sxs-lookup"><span data-stu-id="04352-625">String</span></span>||<span data-ttu-id="04352-p141">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="04352-628">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-628">Object</span></span>| <span data-ttu-id="04352-629">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-629">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-630">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="04352-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04352-631">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-631">Object</span></span>| <span data-ttu-id="04352-632">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-632">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-633">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04352-634">fonction</span><span class="sxs-lookup"><span data-stu-id="04352-634">function</span></span>| <span data-ttu-id="04352-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-635">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-636">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04352-637">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04352-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="04352-638">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="04352-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04352-639">Erreurs</span><span class="sxs-lookup"><span data-stu-id="04352-639">Errors</span></span>

| <span data-ttu-id="04352-640">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="04352-640">Error code</span></span> | <span data-ttu-id="04352-641">Description</span><span class="sxs-lookup"><span data-stu-id="04352-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="04352-642">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="04352-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="04352-643">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="04352-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="04352-644">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="04352-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-645">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-645">Requirements</span></span>

|<span data-ttu-id="04352-646">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-646">Requirement</span></span>| <span data-ttu-id="04352-647">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-648">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-649">1.1</span><span class="sxs-lookup"><span data-stu-id="04352-649">1.1</span></span>|
|[<span data-ttu-id="04352-650">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04352-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="04352-652">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-653">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-654">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="04352-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04352-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="04352-656">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="04352-p142">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="04352-660">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="04352-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="04352-661">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="04352-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-662">Parameters</span><span class="sxs-lookup"><span data-stu-id="04352-662">Parameters</span></span>

|<span data-ttu-id="04352-663">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-663">Name</span></span>| <span data-ttu-id="04352-664">Type</span><span class="sxs-lookup"><span data-stu-id="04352-664">Type</span></span>| <span data-ttu-id="04352-665">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-665">Attributes</span></span>| <span data-ttu-id="04352-666">Description</span><span class="sxs-lookup"><span data-stu-id="04352-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="04352-667">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-667">String</span></span>||<span data-ttu-id="04352-p143">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="04352-670">String</span><span class="sxs-lookup"><span data-stu-id="04352-670">String</span></span>||<span data-ttu-id="04352-671">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="04352-671">The subject of the item to be attached.</span></span> <span data-ttu-id="04352-672">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="04352-673">Object</span><span class="sxs-lookup"><span data-stu-id="04352-673">Object</span></span>| <span data-ttu-id="04352-674">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-674">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-675">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="04352-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04352-676">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-676">Object</span></span>| <span data-ttu-id="04352-677">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-677">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-678">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04352-679">fonction</span><span class="sxs-lookup"><span data-stu-id="04352-679">function</span></span>| <span data-ttu-id="04352-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-680">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-681">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04352-682">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04352-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="04352-683">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="04352-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04352-684">Erreurs</span><span class="sxs-lookup"><span data-stu-id="04352-684">Errors</span></span>

| <span data-ttu-id="04352-685">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="04352-685">Error code</span></span> | <span data-ttu-id="04352-686">Description</span><span class="sxs-lookup"><span data-stu-id="04352-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="04352-687">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="04352-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-688">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-688">Requirements</span></span>

|<span data-ttu-id="04352-689">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-689">Requirement</span></span>| <span data-ttu-id="04352-690">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-691">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-692">1.1</span><span class="sxs-lookup"><span data-stu-id="04352-692">1.1</span></span>|
|[<span data-ttu-id="04352-693">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04352-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="04352-695">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-696">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-697">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-697">Example</span></span>

<span data-ttu-id="04352-698">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="04352-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="04352-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="04352-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="04352-700">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="04352-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-701">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04352-702">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="04352-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="04352-703">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="04352-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="04352-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="04352-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-707">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-707">Parameters</span></span>

|<span data-ttu-id="04352-708">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-708">Name</span></span>| <span data-ttu-id="04352-709">Type</span><span class="sxs-lookup"><span data-stu-id="04352-709">Type</span></span>| <span data-ttu-id="04352-710">Description</span><span class="sxs-lookup"><span data-stu-id="04352-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="04352-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="04352-711">String &#124; Object</span></span>| |<span data-ttu-id="04352-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="04352-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="04352-714">**OU**</span><span class="sxs-lookup"><span data-stu-id="04352-714">**OR**</span></span><br/><span data-ttu-id="04352-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="04352-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="04352-717">String</span><span class="sxs-lookup"><span data-stu-id="04352-717">String</span></span> | <span data-ttu-id="04352-718">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-718">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="04352-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="04352-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="04352-722">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-722">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-723">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="04352-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="04352-724">String</span><span class="sxs-lookup"><span data-stu-id="04352-724">String</span></span> | | <span data-ttu-id="04352-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="04352-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="04352-727">String</span><span class="sxs-lookup"><span data-stu-id="04352-727">String</span></span> | | <span data-ttu-id="04352-728">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="04352-729">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-729">String</span></span> | | <span data-ttu-id="04352-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="04352-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="04352-732">String</span><span class="sxs-lookup"><span data-stu-id="04352-732">String</span></span> | | <span data-ttu-id="04352-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="04352-736">function</span><span class="sxs-lookup"><span data-stu-id="04352-736">function</span></span> | <span data-ttu-id="04352-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-737">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-738">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-739">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-739">Requirements</span></span>

|<span data-ttu-id="04352-740">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-740">Requirement</span></span>| <span data-ttu-id="04352-741">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-742">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-743">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-743">1.0</span></span>|
|[<span data-ttu-id="04352-744">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-745">ReadItem</span></span>|
|[<span data-ttu-id="04352-746">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-747">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="04352-748">Exemples</span><span class="sxs-lookup"><span data-stu-id="04352-748">Examples</span></span>

<span data-ttu-id="04352-749">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="04352-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="04352-750">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="04352-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="04352-751">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="04352-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="04352-752">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="04352-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="04352-753">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="04352-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="04352-754">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="04352-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="04352-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="04352-756">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="04352-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-757">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04352-758">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="04352-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="04352-759">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="04352-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="04352-p152">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="04352-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-763">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-763">Parameters</span></span>

|<span data-ttu-id="04352-764">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-764">Name</span></span>| <span data-ttu-id="04352-765">Type</span><span class="sxs-lookup"><span data-stu-id="04352-765">Type</span></span>| <span data-ttu-id="04352-766">Description</span><span class="sxs-lookup"><span data-stu-id="04352-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="04352-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="04352-767">String &#124; Object</span></span>| | <span data-ttu-id="04352-p153">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="04352-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="04352-770">**OU**</span><span class="sxs-lookup"><span data-stu-id="04352-770">**OR**</span></span><br/><span data-ttu-id="04352-p154">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="04352-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="04352-773">String</span><span class="sxs-lookup"><span data-stu-id="04352-773">String</span></span> | <span data-ttu-id="04352-774">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-774">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-p155">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="04352-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="04352-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="04352-778">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-778">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-779">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="04352-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="04352-780">String</span><span class="sxs-lookup"><span data-stu-id="04352-780">String</span></span> | | <span data-ttu-id="04352-p156">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="04352-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="04352-783">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-783">String</span></span> | | <span data-ttu-id="04352-784">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="04352-785">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-785">String</span></span> | | <span data-ttu-id="04352-p157">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="04352-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="04352-788">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-788">String</span></span> | | <span data-ttu-id="04352-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="04352-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="04352-792">function</span><span class="sxs-lookup"><span data-stu-id="04352-792">function</span></span> | <span data-ttu-id="04352-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-793">&lt;optional&gt;</span></span> | <span data-ttu-id="04352-794">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-795">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-795">Requirements</span></span>

|<span data-ttu-id="04352-796">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-796">Requirement</span></span>| <span data-ttu-id="04352-797">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-798">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-799">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-799">1.0</span></span>|
|[<span data-ttu-id="04352-800">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-801">ReadItem</span></span>|
|[<span data-ttu-id="04352-802">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-803">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="04352-804">Exemples</span><span class="sxs-lookup"><span data-stu-id="04352-804">Examples</span></span>

<span data-ttu-id="04352-805">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="04352-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="04352-806">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="04352-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="04352-807">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="04352-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="04352-808">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="04352-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="04352-809">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="04352-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="04352-810">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="04352-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="04352-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="04352-812">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="04352-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-813">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-814">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-814">Requirements</span></span>

|<span data-ttu-id="04352-815">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-815">Requirement</span></span>| <span data-ttu-id="04352-816">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-817">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-818">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-818">1.0</span></span>|
|[<span data-ttu-id="04352-819">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-820">ReadItem</span></span>|
|[<span data-ttu-id="04352-821">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-822">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-823">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-823">Returns:</span></span>

<span data-ttu-id="04352-824">Type : [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="04352-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="04352-825">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-825">Example</span></span>

<span data-ttu-id="04352-826">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="04352-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="04352-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="04352-828">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="04352-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-829">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-830">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-830">Parameters</span></span>

|<span data-ttu-id="04352-831">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-831">Name</span></span>| <span data-ttu-id="04352-832">Type</span><span class="sxs-lookup"><span data-stu-id="04352-832">Type</span></span>| <span data-ttu-id="04352-833">Description</span><span class="sxs-lookup"><span data-stu-id="04352-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="04352-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="04352-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="04352-835">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="04352-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04352-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-836">Requirements</span></span>

|<span data-ttu-id="04352-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-837">Requirement</span></span>| <span data-ttu-id="04352-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-840">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-840">1.0</span></span>|
|[<span data-ttu-id="04352-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-842">Restreinte</span><span class="sxs-lookup"><span data-stu-id="04352-842">Restricted</span></span>|
|[<span data-ttu-id="04352-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-845">Returns:</span></span>

<span data-ttu-id="04352-846">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="04352-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="04352-847">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="04352-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="04352-848">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="04352-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="04352-849">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="04352-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="04352-850">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="04352-850">Value of `entityType`</span></span> | <span data-ttu-id="04352-851">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="04352-851">Type of objects in returned array</span></span> | <span data-ttu-id="04352-852">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="04352-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="04352-853">String</span><span class="sxs-lookup"><span data-stu-id="04352-853">String</span></span> | <span data-ttu-id="04352-854">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04352-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="04352-855">Contact</span><span class="sxs-lookup"><span data-stu-id="04352-855">Contact</span></span> | <span data-ttu-id="04352-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04352-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="04352-857">String</span><span class="sxs-lookup"><span data-stu-id="04352-857">String</span></span> | <span data-ttu-id="04352-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04352-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="04352-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="04352-859">MeetingSuggestion</span></span> | <span data-ttu-id="04352-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04352-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="04352-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="04352-861">PhoneNumber</span></span> | <span data-ttu-id="04352-862">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04352-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="04352-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="04352-863">TaskSuggestion</span></span> | <span data-ttu-id="04352-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04352-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="04352-865">String</span><span class="sxs-lookup"><span data-stu-id="04352-865">String</span></span> | <span data-ttu-id="04352-866">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04352-866">**Restricted**</span></span> |

<span data-ttu-id="04352-867">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="04352-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="04352-868">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-868">Example</span></span>

<span data-ttu-id="04352-869">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="04352-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="04352-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="04352-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="04352-871">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="04352-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-872">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04352-873">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="04352-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-874">Parameters</span><span class="sxs-lookup"><span data-stu-id="04352-874">Parameters</span></span>

|<span data-ttu-id="04352-875">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-875">Name</span></span>| <span data-ttu-id="04352-876">Type</span><span class="sxs-lookup"><span data-stu-id="04352-876">Type</span></span>| <span data-ttu-id="04352-877">Description</span><span class="sxs-lookup"><span data-stu-id="04352-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="04352-878">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-878">String</span></span>|<span data-ttu-id="04352-879">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="04352-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04352-880">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-880">Requirements</span></span>

|<span data-ttu-id="04352-881">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-881">Requirement</span></span>| <span data-ttu-id="04352-882">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-883">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-884">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-884">1.0</span></span>|
|[<span data-ttu-id="04352-885">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-886">ReadItem</span></span>|
|[<span data-ttu-id="04352-887">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-888">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-889">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-889">Returns:</span></span>

<span data-ttu-id="04352-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="04352-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="04352-892">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="04352-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="04352-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="04352-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="04352-894">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="04352-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-895">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04352-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="04352-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="04352-899">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="04352-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="04352-900">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="04352-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="04352-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="04352-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04352-903">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-903">Requirements</span></span>

|<span data-ttu-id="04352-904">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-904">Requirement</span></span>| <span data-ttu-id="04352-905">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-906">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-907">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-907">1.0</span></span>|
|[<span data-ttu-id="04352-908">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-909">ReadItem</span></span>|
|[<span data-ttu-id="04352-910">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-911">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-912">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-912">Returns:</span></span>

<span data-ttu-id="04352-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="04352-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="04352-915">Type : Objet</span><span class="sxs-lookup"><span data-stu-id="04352-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="04352-916">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-916">Example</span></span>

<span data-ttu-id="04352-917">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="04352-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="04352-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="04352-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="04352-919">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="04352-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-920">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="04352-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04352-921">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="04352-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="04352-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="04352-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-924">Parameters</span><span class="sxs-lookup"><span data-stu-id="04352-924">Parameters</span></span>

|<span data-ttu-id="04352-925">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-925">Name</span></span>| <span data-ttu-id="04352-926">Type</span><span class="sxs-lookup"><span data-stu-id="04352-926">Type</span></span>| <span data-ttu-id="04352-927">Description</span><span class="sxs-lookup"><span data-stu-id="04352-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="04352-928">String</span><span class="sxs-lookup"><span data-stu-id="04352-928">String</span></span>|<span data-ttu-id="04352-929">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="04352-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04352-930">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-930">Requirements</span></span>

|<span data-ttu-id="04352-931">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-931">Requirement</span></span>| <span data-ttu-id="04352-932">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-933">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-934">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-934">1.0</span></span>|
|[<span data-ttu-id="04352-935">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-936">ReadItem</span></span>|
|[<span data-ttu-id="04352-937">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-938">Lecture</span><span class="sxs-lookup"><span data-stu-id="04352-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-939">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-939">Returns:</span></span>

<span data-ttu-id="04352-940">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="04352-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="04352-941">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="04352-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="04352-942">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="04352-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="04352-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="04352-944">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="04352-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="04352-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="04352-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="04352-947">Dans Outlook sur le Web, la méthode renvoie la chaîne « NULL » si aucun texte n’est sélectionné, mais que le curseur se trouve dans le corps.</span><span class="sxs-lookup"><span data-stu-id="04352-947">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="04352-948">Pour vérifier cette situation, incluez un code similaire à celui-ci :</span><span class="sxs-lookup"><span data-stu-id="04352-948">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="04352-949">Parameters</span><span class="sxs-lookup"><span data-stu-id="04352-949">Parameters</span></span>

|<span data-ttu-id="04352-950">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-950">Name</span></span>| <span data-ttu-id="04352-951">Type</span><span class="sxs-lookup"><span data-stu-id="04352-951">Type</span></span>| <span data-ttu-id="04352-952">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-952">Attributes</span></span>| <span data-ttu-id="04352-953">Description</span><span class="sxs-lookup"><span data-stu-id="04352-953">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="04352-954">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="04352-954">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="04352-p167">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="04352-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="04352-958">Object</span><span class="sxs-lookup"><span data-stu-id="04352-958">Object</span></span>| <span data-ttu-id="04352-959">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-959">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-960">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="04352-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04352-961">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-961">Object</span></span>| <span data-ttu-id="04352-962">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-962">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-963">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04352-964">fonction</span><span class="sxs-lookup"><span data-stu-id="04352-964">function</span></span>||<span data-ttu-id="04352-965">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-965">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="04352-966">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="04352-966">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="04352-967">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="04352-967">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04352-968">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-968">Requirements</span></span>

|<span data-ttu-id="04352-969">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-969">Requirement</span></span>| <span data-ttu-id="04352-970">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-971">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-972">1.2</span><span class="sxs-lookup"><span data-stu-id="04352-972">1.2</span></span>|
|[<span data-ttu-id="04352-973">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-974">ReadItem</span></span>|
|[<span data-ttu-id="04352-975">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-976">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-976">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="04352-977">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="04352-977">Returns:</span></span>

<span data-ttu-id="04352-978">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="04352-978">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="04352-979">Type : String</span><span class="sxs-lookup"><span data-stu-id="04352-979">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="04352-980">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-980">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="04352-981">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="04352-981">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="04352-982">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="04352-982">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="04352-p169">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="04352-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-986">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-986">Parameters</span></span>

|<span data-ttu-id="04352-987">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-987">Name</span></span>| <span data-ttu-id="04352-988">Type</span><span class="sxs-lookup"><span data-stu-id="04352-988">Type</span></span>| <span data-ttu-id="04352-989">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-989">Attributes</span></span>| <span data-ttu-id="04352-990">Description</span><span class="sxs-lookup"><span data-stu-id="04352-990">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="04352-991">function</span><span class="sxs-lookup"><span data-stu-id="04352-991">function</span></span>||<span data-ttu-id="04352-992">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="04352-993">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04352-993">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="04352-994">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="04352-994">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="04352-995">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-995">Object</span></span>| <span data-ttu-id="04352-996">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-996">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-997">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-997">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="04352-998">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-998">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04352-999">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-999">Requirements</span></span>

|<span data-ttu-id="04352-1000">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-1000">Requirement</span></span>| <span data-ttu-id="04352-1001">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-1002">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="04352-1003">1.0</span></span>|
|[<span data-ttu-id="04352-1004">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04352-1005">ReadItem</span></span>|
|[<span data-ttu-id="04352-1006">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-1007">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="04352-1007">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-1008">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-1008">Example</span></span>

<span data-ttu-id="04352-p172">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="04352-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="04352-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04352-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="04352-1013">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="04352-1013">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="04352-1014">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="04352-1014">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="04352-1015">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="04352-1015">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="04352-1016">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="04352-1016">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="04352-1017">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="04352-1017">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-1018">Paramètres</span><span class="sxs-lookup"><span data-stu-id="04352-1018">Parameters</span></span>

|<span data-ttu-id="04352-1019">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-1019">Name</span></span>| <span data-ttu-id="04352-1020">Type</span><span class="sxs-lookup"><span data-stu-id="04352-1020">Type</span></span>| <span data-ttu-id="04352-1021">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-1021">Attributes</span></span>| <span data-ttu-id="04352-1022">Description</span><span class="sxs-lookup"><span data-stu-id="04352-1022">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="04352-1023">Chaîne</span><span class="sxs-lookup"><span data-stu-id="04352-1023">String</span></span>||<span data-ttu-id="04352-1024">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="04352-1024">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="04352-1025">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-1025">Object</span></span>| <span data-ttu-id="04352-1026">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1027">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="04352-1027">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04352-1028">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-1028">Object</span></span>| <span data-ttu-id="04352-1029">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1029">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1030">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-1030">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04352-1031">fonction</span><span class="sxs-lookup"><span data-stu-id="04352-1031">function</span></span>| <span data-ttu-id="04352-1032">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1032">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1033">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-1033">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04352-1034">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="04352-1034">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04352-1035">Erreurs</span><span class="sxs-lookup"><span data-stu-id="04352-1035">Errors</span></span>

| <span data-ttu-id="04352-1036">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="04352-1036">Error code</span></span> | <span data-ttu-id="04352-1037">Description</span><span class="sxs-lookup"><span data-stu-id="04352-1037">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="04352-1038">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="04352-1038">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-1039">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-1039">Requirements</span></span>

|<span data-ttu-id="04352-1040">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-1040">Requirement</span></span>| <span data-ttu-id="04352-1041">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-1042">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-1043">1.1</span><span class="sxs-lookup"><span data-stu-id="04352-1043">1.1</span></span>|
|[<span data-ttu-id="04352-1044">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-1044">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-1045">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04352-1045">ReadWriteItem</span></span>|
|[<span data-ttu-id="04352-1046">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-1046">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-1047">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-1047">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-1048">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-1048">Example</span></span>

<span data-ttu-id="04352-1049">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="04352-1049">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="04352-1050">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="04352-1050">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="04352-1051">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="04352-1051">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="04352-p174">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="04352-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04352-1055">Parameters</span><span class="sxs-lookup"><span data-stu-id="04352-1055">Parameters</span></span>

|<span data-ttu-id="04352-1056">Nom</span><span class="sxs-lookup"><span data-stu-id="04352-1056">Name</span></span>| <span data-ttu-id="04352-1057">Type</span><span class="sxs-lookup"><span data-stu-id="04352-1057">Type</span></span>| <span data-ttu-id="04352-1058">Attributs</span><span class="sxs-lookup"><span data-stu-id="04352-1058">Attributes</span></span>| <span data-ttu-id="04352-1059">Description</span><span class="sxs-lookup"><span data-stu-id="04352-1059">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="04352-1060">String</span><span class="sxs-lookup"><span data-stu-id="04352-1060">String</span></span>||<span data-ttu-id="04352-p175">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="04352-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="04352-1064">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-1064">Object</span></span>| <span data-ttu-id="04352-1065">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1065">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1066">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="04352-1066">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04352-1067">Objet</span><span class="sxs-lookup"><span data-stu-id="04352-1067">Object</span></span>| <span data-ttu-id="04352-1068">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1069">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="04352-1069">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="04352-1070">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="04352-1070">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="04352-1071">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="04352-1071">&lt;optional&gt;</span></span>|<span data-ttu-id="04352-1072">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="04352-1072">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="04352-1073">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="04352-1073">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="04352-1074">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="04352-1074">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="04352-1075">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="04352-1075">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="04352-1076">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="04352-1076">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="04352-1077">fonction</span><span class="sxs-lookup"><span data-stu-id="04352-1077">function</span></span>||<span data-ttu-id="04352-1078">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04352-1078">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04352-1079">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="04352-1079">Requirements</span></span>

|<span data-ttu-id="04352-1080">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="04352-1080">Requirement</span></span>| <span data-ttu-id="04352-1081">Valeur</span><span class="sxs-lookup"><span data-stu-id="04352-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="04352-1082">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="04352-1082">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04352-1083">1.2</span><span class="sxs-lookup"><span data-stu-id="04352-1083">1.2</span></span>|
|[<span data-ttu-id="04352-1084">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="04352-1084">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04352-1085">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04352-1085">ReadWriteItem</span></span>|
|[<span data-ttu-id="04352-1086">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="04352-1086">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04352-1087">Composition</span><span class="sxs-lookup"><span data-stu-id="04352-1087">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04352-1088">Exemple</span><span class="sxs-lookup"><span data-stu-id="04352-1088">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
