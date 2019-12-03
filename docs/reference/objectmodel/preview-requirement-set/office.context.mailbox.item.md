---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 2ebcacb1f99df047b5f5c5ebe82c012e21e45d3c
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670138"
---
# <a name="item"></a><span data-ttu-id="f30a2-102">élément</span><span class="sxs-lookup"><span data-stu-id="f30a2-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f30a2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f30a2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f30a2-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-mailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f30a2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-mailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-106">Requirements</span></span>

|<span data-ttu-id="f30a2-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-107">Requirement</span></span>|<span data-ttu-id="f30a2-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-110">1.0</span></span>|
|[<span data-ttu-id="f30a2-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f30a2-112">Restricted</span></span>|
|[<span data-ttu-id="f30a2-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-114">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f30a2-115">Propriétés</span><span class="sxs-lookup"><span data-stu-id="f30a2-115">Properties</span></span>

| <span data-ttu-id="f30a2-116">Propriété</span><span class="sxs-lookup"><span data-stu-id="f30a2-116">Property</span></span> | <span data-ttu-id="f30a2-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="f30a2-117">Minimum</span></span><br><span data-ttu-id="f30a2-118">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f30a2-118">permission level</span></span> | <span data-ttu-id="f30a2-119">Modes</span><span class="sxs-lookup"><span data-stu-id="f30a2-119">Modes</span></span> | <span data-ttu-id="f30a2-120">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f30a2-120">Return type</span></span> | <span data-ttu-id="f30a2-121">Minimale</span><span class="sxs-lookup"><span data-stu-id="f30a2-121">Minimum</span></span><br><span data-ttu-id="f30a2-122">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-122">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="f30a2-123">attachments</span><span class="sxs-lookup"><span data-stu-id="f30a2-123">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="f30a2-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-124">ReadItem</span></span> | <span data-ttu-id="f30a2-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-125">Read</span></span> | <span data-ttu-id="f30a2-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span> | <span data-ttu-id="f30a2-127">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-127">1.0</span></span> |
| [<span data-ttu-id="f30a2-128">bcc</span><span class="sxs-lookup"><span data-stu-id="f30a2-128">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="f30a2-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-129">ReadItem</span></span> | <span data-ttu-id="f30a2-130">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-130">Message Compose</span></span> | [<span data-ttu-id="f30a2-131">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-131">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="f30a2-132">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-132">1.1</span></span> |
| [<span data-ttu-id="f30a2-133">body</span><span class="sxs-lookup"><span data-stu-id="f30a2-133">body</span></span>](#body-body) | <span data-ttu-id="f30a2-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-134">ReadItem</span></span> | <span data-ttu-id="f30a2-135">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-135">Compose</span></span> | [<span data-ttu-id="f30a2-136">Body</span><span class="sxs-lookup"><span data-stu-id="f30a2-136">Body</span></span>](/javascript/api/outlook/office.body) | <span data-ttu-id="f30a2-137">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-137">1.1</span></span> |
| | | <span data-ttu-id="f30a2-138">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-138">Read</span></span> | | |
| [<span data-ttu-id="f30a2-139">categories</span><span class="sxs-lookup"><span data-stu-id="f30a2-139">categories</span></span>](#categories-categories) | <span data-ttu-id="f30a2-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-140">ReadItem</span></span> | <span data-ttu-id="f30a2-141">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-141">Compose</span></span> | [<span data-ttu-id="f30a2-142">Categories</span><span class="sxs-lookup"><span data-stu-id="f30a2-142">Categories</span></span>](/javascript/api/outlook/office.categories) | <span data-ttu-id="f30a2-143">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-143">1.8</span></span> |
| | | <span data-ttu-id="f30a2-144">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-144">Read</span></span> | | |
| [<span data-ttu-id="f30a2-145">cc</span><span class="sxs-lookup"><span data-stu-id="f30a2-145">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f30a2-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-146">ReadItem</span></span> | <span data-ttu-id="f30a2-147">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-147">Message Compose</span></span> | [<span data-ttu-id="f30a2-148">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-148">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="f30a2-149">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-149">1.0</span></span> |
| | | <span data-ttu-id="f30a2-150">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-150">Message Read</span></span> | <span data-ttu-id="f30a2-151">Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-151">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="f30a2-152">conversationId</span><span class="sxs-lookup"><span data-stu-id="f30a2-152">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f30a2-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-153">ReadItem</span></span> | <span data-ttu-id="f30a2-154">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-154">Message Compose</span></span> | <span data-ttu-id="f30a2-155">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-155">String</span></span> | <span data-ttu-id="f30a2-156">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-156">1.0</span></span> |
| | | <span data-ttu-id="f30a2-157">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-157">Message Read</span></span> | | |
| [<span data-ttu-id="f30a2-158">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f30a2-158">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f30a2-159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-159">ReadItem</span></span> | <span data-ttu-id="f30a2-160">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-160">Read</span></span> | <span data-ttu-id="f30a2-161">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-161">Date</span></span> | <span data-ttu-id="f30a2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-162">1.0</span></span> |
| [<span data-ttu-id="f30a2-163">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f30a2-163">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f30a2-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-164">ReadItem</span></span> | <span data-ttu-id="f30a2-165">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-165">Read</span></span> | <span data-ttu-id="f30a2-166">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-166">Date</span></span> | <span data-ttu-id="f30a2-167">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-167">1.0</span></span> |
| [<span data-ttu-id="f30a2-168">end</span><span class="sxs-lookup"><span data-stu-id="f30a2-168">end</span></span>](#end-datetime) | <span data-ttu-id="f30a2-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-169">ReadItem</span></span> | <span data-ttu-id="f30a2-170">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-170">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-171">Heure</span><span class="sxs-lookup"><span data-stu-id="f30a2-171">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="f30a2-172">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-172">1.0</span></span> |
| | | <span data-ttu-id="f30a2-173">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-173">Appointment Attendee</span></span> | <span data-ttu-id="f30a2-174">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-174">Date</span></span> | |
| | | <span data-ttu-id="f30a2-175">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-175">Message Read</span></span><br><span data-ttu-id="f30a2-176">(Demande de réunion)</span><span class="sxs-lookup"><span data-stu-id="f30a2-176">(Meeting Request)</span></span> | <span data-ttu-id="f30a2-177">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-177">Date</span></span> | |
| [<span data-ttu-id="f30a2-178">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f30a2-178">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="f30a2-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-179">ReadItem</span></span> | <span data-ttu-id="f30a2-180">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-180">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-181">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f30a2-181">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation) | <span data-ttu-id="f30a2-182">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-182">1.8</span></span> |
| | | <span data-ttu-id="f30a2-183">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-183">Appointment Attendee</span></span> | | |
| [<span data-ttu-id="f30a2-184">from</span><span class="sxs-lookup"><span data-stu-id="f30a2-184">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="f30a2-185">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-185">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-186">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-186">Message Compose</span></span> | [<span data-ttu-id="f30a2-187">From</span><span class="sxs-lookup"><span data-stu-id="f30a2-187">From</span></span>](/javascript/api/outlook/office.from) | <span data-ttu-id="f30a2-188">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-188">1.7</span></span> |
| | <span data-ttu-id="f30a2-189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-189">ReadItem</span></span> | <span data-ttu-id="f30a2-190">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-190">Message Read</span></span> | [<span data-ttu-id="f30a2-191">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-191">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="f30a2-192">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-192">1.0</span></span> |
| [<span data-ttu-id="f30a2-193">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="f30a2-193">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="f30a2-194">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-194">ReadItem</span></span> | <span data-ttu-id="f30a2-195">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-195">Message Compose</span></span> | [<span data-ttu-id="f30a2-196">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f30a2-196">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders) | <span data-ttu-id="f30a2-197">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-197">1.8</span></span> |
| [<span data-ttu-id="f30a2-198">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f30a2-198">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f30a2-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-199">ReadItem</span></span> | <span data-ttu-id="f30a2-200">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-200">Message Read</span></span> | <span data-ttu-id="f30a2-201">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-201">String</span></span> | <span data-ttu-id="f30a2-202">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-202">1.0</span></span> |
| [<span data-ttu-id="f30a2-203">itemClass</span><span class="sxs-lookup"><span data-stu-id="f30a2-203">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f30a2-204">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-204">ReadItem</span></span> | <span data-ttu-id="f30a2-205">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-205">Read</span></span> | <span data-ttu-id="f30a2-206">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-206">String</span></span> | <span data-ttu-id="f30a2-207">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-207">1.0</span></span> |
| [<span data-ttu-id="f30a2-208">itemId</span><span class="sxs-lookup"><span data-stu-id="f30a2-208">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f30a2-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-209">ReadItem</span></span> | <span data-ttu-id="f30a2-210">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-210">Read</span></span> | <span data-ttu-id="f30a2-211">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-211">String</span></span> | <span data-ttu-id="f30a2-212">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-212">1.0</span></span> |
| [<span data-ttu-id="f30a2-213">itemType</span><span class="sxs-lookup"><span data-stu-id="f30a2-213">itemType</span></span>](#itemtype-mailboxenumsitemtype) | <span data-ttu-id="f30a2-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-214">ReadItem</span></span> | <span data-ttu-id="f30a2-215">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-215">Compose</span></span> | [<span data-ttu-id="f30a2-216">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="f30a2-216">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype) | <span data-ttu-id="f30a2-217">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-217">1.0</span></span> |
| | | <span data-ttu-id="f30a2-218">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-218">Read</span></span> | | |
| [<span data-ttu-id="f30a2-219">location</span><span class="sxs-lookup"><span data-stu-id="f30a2-219">location</span></span>](#location-stringlocation) | <span data-ttu-id="f30a2-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-220">ReadItem</span></span> | <span data-ttu-id="f30a2-221">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-221">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-222">Location</span><span class="sxs-lookup"><span data-stu-id="f30a2-222">Location</span></span>](/javascript/api/outlook/office.location) | <span data-ttu-id="f30a2-223">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-223">1.0</span></span> |
| | | <span data-ttu-id="f30a2-224">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-224">Appointment Attendee</span></span> | <span data-ttu-id="f30a2-225">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-225">String</span></span> | |
| | | <span data-ttu-id="f30a2-226">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-226">Message Read</span></span><br><span data-ttu-id="f30a2-227">(Demande de réunion)</span><span class="sxs-lookup"><span data-stu-id="f30a2-227">(Meeting Request)</span></span> | <span data-ttu-id="f30a2-228">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-228">String</span></span> | |
| [<span data-ttu-id="f30a2-229">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f30a2-229">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f30a2-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-230">ReadItem</span></span> | <span data-ttu-id="f30a2-231">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-231">Read</span></span> | <span data-ttu-id="f30a2-232">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-232">String</span></span> | <span data-ttu-id="f30a2-233">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-233">1.0</span></span> |
| [<span data-ttu-id="f30a2-234">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f30a2-234">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="f30a2-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-235">ReadItem</span></span> | <span data-ttu-id="f30a2-236">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-236">Message Compose</span></span> | [<span data-ttu-id="f30a2-237">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f30a2-237">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages) | <span data-ttu-id="f30a2-238">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-238">1.3</span></span> |
| | <span data-ttu-id="f30a2-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-239">ReadItem</span></span> | <span data-ttu-id="f30a2-240">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-240">Message Read</span></span> | | |
| [<span data-ttu-id="f30a2-241">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f30a2-241">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f30a2-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-242">ReadItem</span></span> | <span data-ttu-id="f30a2-243">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-243">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-244">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-244">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="f30a2-245">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-245">1.0</span></span> |
| | | <span data-ttu-id="f30a2-246">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-246">Appointment Attendee</span></span> | <span data-ttu-id="f30a2-247">Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-247">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="f30a2-248">organizer</span><span class="sxs-lookup"><span data-stu-id="f30a2-248">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="f30a2-249">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-249">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-250">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-250">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-251">Organizer</span><span class="sxs-lookup"><span data-stu-id="f30a2-251">Organizer</span></span>](/javascript/api/outlook/office.organizer) | <span data-ttu-id="f30a2-252">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-252">1.7</span></span> |
| | <span data-ttu-id="f30a2-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-253">ReadItem</span></span> | <span data-ttu-id="f30a2-254">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-254">Appointment Attendee</span></span> | [<span data-ttu-id="f30a2-255">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-255">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="f30a2-256">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-256">1.0</span></span> |
| [<span data-ttu-id="f30a2-257">recurrence</span><span class="sxs-lookup"><span data-stu-id="f30a2-257">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="f30a2-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-258">ReadItem</span></span> | <span data-ttu-id="f30a2-259">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-259">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-260">Instances</span><span class="sxs-lookup"><span data-stu-id="f30a2-260">Recurrence</span></span>](/javascript/api/outlook/office.recurrence) | <span data-ttu-id="f30a2-261">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-261">1.7</span></span> |
| | | <span data-ttu-id="f30a2-262">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-262">Appointment Attendee</span></span> | | |
| | | <span data-ttu-id="f30a2-263">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-263">Message Read</span></span><br><span data-ttu-id="f30a2-264">(Demande de réunion)</span><span class="sxs-lookup"><span data-stu-id="f30a2-264">(Meeting Request)</span></span> | | |
| [<span data-ttu-id="f30a2-265">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f30a2-265">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f30a2-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-266">ReadItem</span></span> | <span data-ttu-id="f30a2-267">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-267">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-268">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-268">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="f30a2-269">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-269">1.0</span></span> |
| | | <span data-ttu-id="f30a2-270">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-270">Appointment Attendee</span></span> | <span data-ttu-id="f30a2-271">Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-271">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="f30a2-272">sender</span><span class="sxs-lookup"><span data-stu-id="f30a2-272">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="f30a2-273">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-273">ReadItem</span></span> | <span data-ttu-id="f30a2-274">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-274">Message Read</span></span> | [<span data-ttu-id="f30a2-275">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-275">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="f30a2-276">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-276">1.0</span></span> |
| [<span data-ttu-id="f30a2-277">seriesId</span><span class="sxs-lookup"><span data-stu-id="f30a2-277">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="f30a2-278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-278">ReadItem</span></span> | <span data-ttu-id="f30a2-279">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-279">Compose</span></span> | <span data-ttu-id="f30a2-280">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-280">String</span></span> | <span data-ttu-id="f30a2-281">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-281">1.7</span></span> |
| | | <span data-ttu-id="f30a2-282">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-282">Read</span></span> | | |
| [<span data-ttu-id="f30a2-283">start</span><span class="sxs-lookup"><span data-stu-id="f30a2-283">start</span></span>](#start-datetime) | <span data-ttu-id="f30a2-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-284">ReadItem</span></span> | <span data-ttu-id="f30a2-285">Organisateur de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-285">Appointment Organizer</span></span> | [<span data-ttu-id="f30a2-286">Heure</span><span class="sxs-lookup"><span data-stu-id="f30a2-286">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="f30a2-287">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-287">1.0</span></span> |
| | | <span data-ttu-id="f30a2-288">Participant à un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-288">Appointment Attendee</span></span> | <span data-ttu-id="f30a2-289">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-289">Date</span></span> | |
| | | <span data-ttu-id="f30a2-290">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-290">Message Read</span></span><br><span data-ttu-id="f30a2-291">(Demande de réunion)</span><span class="sxs-lookup"><span data-stu-id="f30a2-291">(Meeting Request)</span></span> | <span data-ttu-id="f30a2-292">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-292">Date</span></span> | |
| [<span data-ttu-id="f30a2-293">subject</span><span class="sxs-lookup"><span data-stu-id="f30a2-293">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="f30a2-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-294">ReadItem</span></span> | <span data-ttu-id="f30a2-295">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-295">Compose</span></span> | [<span data-ttu-id="f30a2-296">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-296">Subject</span></span>](/javascript/api/outlook/office.subject) | <span data-ttu-id="f30a2-297">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-297">1.0</span></span> |
| | | <span data-ttu-id="f30a2-298">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-298">Read</span></span> | <span data-ttu-id="f30a2-299">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-299">String</span></span> | |
| [<span data-ttu-id="f30a2-300">to</span><span class="sxs-lookup"><span data-stu-id="f30a2-300">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f30a2-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-301">ReadItem</span></span> | <span data-ttu-id="f30a2-302">Composition de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-302">Message Compose</span></span> | [<span data-ttu-id="f30a2-303">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-303">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="f30a2-304">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-304">1.0</span></span> |
| | | <span data-ttu-id="f30a2-305">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-305">Message Read</span></span> | <span data-ttu-id="f30a2-306">Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-306">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |

##### <a name="methods"></a><span data-ttu-id="f30a2-307">Méthodes</span><span class="sxs-lookup"><span data-stu-id="f30a2-307">Methods</span></span>

| <span data-ttu-id="f30a2-308">Méthode</span><span class="sxs-lookup"><span data-stu-id="f30a2-308">Method</span></span> | <span data-ttu-id="f30a2-309">Minimale</span><span class="sxs-lookup"><span data-stu-id="f30a2-309">Minimum</span></span><br><span data-ttu-id="f30a2-310">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f30a2-310">permission level</span></span> | <span data-ttu-id="f30a2-311">Modes</span><span class="sxs-lookup"><span data-stu-id="f30a2-311">Modes</span></span> | <span data-ttu-id="f30a2-312">Minimale</span><span class="sxs-lookup"><span data-stu-id="f30a2-312">Minimum</span></span><br><span data-ttu-id="f30a2-313">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-313">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="f30a2-314">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-314">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f30a2-315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-315">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-316">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-316">Compose</span></span> | <span data-ttu-id="f30a2-317">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-317">1.1</span></span> |
| [<span data-ttu-id="f30a2-318">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f30a2-318">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="f30a2-319">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-319">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-320">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-320">Compose</span></span> | <span data-ttu-id="f30a2-321">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-321">1.8</span></span> |
| [<span data-ttu-id="f30a2-322">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-322">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f30a2-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-323">ReadItem</span></span> | <span data-ttu-id="f30a2-324">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-324">Compose</span></span><br><span data-ttu-id="f30a2-325">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-325">Read</span></span> | <span data-ttu-id="f30a2-326">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-326">1.7</span></span> |
| [<span data-ttu-id="f30a2-327">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-327">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f30a2-328">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-328">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-329">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-329">Compose</span></span> | <span data-ttu-id="f30a2-330">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-330">1.1</span></span> |
| [<span data-ttu-id="f30a2-331">close</span><span class="sxs-lookup"><span data-stu-id="f30a2-331">close</span></span>](#close) | <span data-ttu-id="f30a2-332">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f30a2-332">Restricted</span></span> | <span data-ttu-id="f30a2-333">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-333">Compose</span></span> | <span data-ttu-id="f30a2-334">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-334">1.3</span></span> |
| [<span data-ttu-id="f30a2-335">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f30a2-335">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="f30a2-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-336">ReadItem</span></span> | <span data-ttu-id="f30a2-337">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-337">Read</span></span> | <span data-ttu-id="f30a2-338">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-338">1.0</span></span> |
| [<span data-ttu-id="f30a2-339">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f30a2-339">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="f30a2-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-340">ReadItem</span></span> | <span data-ttu-id="f30a2-341">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-341">Read</span></span> | <span data-ttu-id="f30a2-342">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-342">1.0</span></span> |
| [<span data-ttu-id="f30a2-343">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-343">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="f30a2-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-344">ReadItem</span></span> | <span data-ttu-id="f30a2-345">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-345">Message Read</span></span> | <span data-ttu-id="f30a2-346">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-346">1.8</span></span> |
| [<span data-ttu-id="f30a2-347">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-347">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="f30a2-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-348">ReadItem</span></span> | <span data-ttu-id="f30a2-349">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-349">Compose</span></span><br><span data-ttu-id="f30a2-350">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-350">Read</span></span> | <span data-ttu-id="f30a2-351">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-351">1.8</span></span> |
| [<span data-ttu-id="f30a2-352">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-352">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="f30a2-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-353">ReadItem</span></span> | <span data-ttu-id="f30a2-354">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-354">Compose</span></span> | <span data-ttu-id="f30a2-355">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-355">1.8</span></span> |
| [<span data-ttu-id="f30a2-356">getEntities</span><span class="sxs-lookup"><span data-stu-id="f30a2-356">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="f30a2-357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-357">ReadItem</span></span> | <span data-ttu-id="f30a2-358">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-358">Read</span></span> | <span data-ttu-id="f30a2-359">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-359">1.0</span></span> |
| [<span data-ttu-id="f30a2-360">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f30a2-360">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f30a2-361">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f30a2-361">Restricted</span></span> | <span data-ttu-id="f30a2-362">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-362">Read</span></span> | <span data-ttu-id="f30a2-363">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-363">1.0</span></span> |
| [<span data-ttu-id="f30a2-364">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f30a2-364">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f30a2-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-365">ReadItem</span></span> | <span data-ttu-id="f30a2-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-366">Read</span></span> | <span data-ttu-id="f30a2-367">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-367">1.0</span></span> |
| [<span data-ttu-id="f30a2-368">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-368">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="f30a2-369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-369">ReadItem</span></span> | <span data-ttu-id="f30a2-370">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-370">Read</span></span> | <span data-ttu-id="f30a2-371">Aperçu</span><span class="sxs-lookup"><span data-stu-id="f30a2-371">Preview</span></span> |
| [<span data-ttu-id="f30a2-372">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-372">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="f30a2-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-373">ReadItem</span></span> | <span data-ttu-id="f30a2-374">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-374">Compose</span></span> | <span data-ttu-id="f30a2-375">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-375">1.8</span></span> |
| [<span data-ttu-id="f30a2-376">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f30a2-376">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f30a2-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-377">ReadItem</span></span> | <span data-ttu-id="f30a2-378">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-378">Read</span></span> | <span data-ttu-id="f30a2-379">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-379">1.0</span></span> |
| [<span data-ttu-id="f30a2-380">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f30a2-380">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f30a2-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-381">ReadItem</span></span> | <span data-ttu-id="f30a2-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-382">Read</span></span> | <span data-ttu-id="f30a2-383">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-383">1.0</span></span> |
| [<span data-ttu-id="f30a2-384">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-384">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f30a2-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-385">ReadItem</span></span> | <span data-ttu-id="f30a2-386">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-386">Compose</span></span> | <span data-ttu-id="f30a2-387">1.2</span><span class="sxs-lookup"><span data-stu-id="f30a2-387">1.2</span></span> |
| [<span data-ttu-id="f30a2-388">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="f30a2-388">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="f30a2-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-389">ReadItem</span></span> | <span data-ttu-id="f30a2-390">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-390">Read</span></span> | <span data-ttu-id="f30a2-391">1.6</span><span class="sxs-lookup"><span data-stu-id="f30a2-391">1.6</span></span> |
| [<span data-ttu-id="f30a2-392">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f30a2-392">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f30a2-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-393">ReadItem</span></span> | <span data-ttu-id="f30a2-394">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-394">Read</span></span> | <span data-ttu-id="f30a2-395">1.6</span><span class="sxs-lookup"><span data-stu-id="f30a2-395">1.6</span></span> |
| [<span data-ttu-id="f30a2-396">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-396">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="f30a2-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-397">ReadItem</span></span> | <span data-ttu-id="f30a2-398">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-398">Compose</span></span><br><span data-ttu-id="f30a2-399">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-399">Read</span></span> | <span data-ttu-id="f30a2-400">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-400">1.8</span></span> |
| [<span data-ttu-id="f30a2-401">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-401">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f30a2-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-402">ReadItem</span></span> | <span data-ttu-id="f30a2-403">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-403">Compose</span></span><br><span data-ttu-id="f30a2-404">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-404">Read</span></span> | <span data-ttu-id="f30a2-405">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-405">1.0</span></span> |
| [<span data-ttu-id="f30a2-406">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-406">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f30a2-407">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-407">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-408">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-408">Compose</span></span> | <span data-ttu-id="f30a2-409">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-409">1.1</span></span> |
| [<span data-ttu-id="f30a2-410">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-410">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="f30a2-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-411">ReadItem</span></span> | <span data-ttu-id="f30a2-412">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-412">Compose</span></span><br><span data-ttu-id="f30a2-413">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-413">Read</span></span> | <span data-ttu-id="f30a2-414">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-414">1.7</span></span> |
| [<span data-ttu-id="f30a2-415">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-415">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f30a2-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-416">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-417">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-417">Compose</span></span> | <span data-ttu-id="f30a2-418">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-418">1.3</span></span> |
| [<span data-ttu-id="f30a2-419">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f30a2-419">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f30a2-420">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-420">ReadWriteItem</span></span> | <span data-ttu-id="f30a2-421">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-421">Compose</span></span> | <span data-ttu-id="f30a2-422">1.2</span><span class="sxs-lookup"><span data-stu-id="f30a2-422">1.2</span></span> |

##### <a name="events"></a><span data-ttu-id="f30a2-423">Événements</span><span class="sxs-lookup"><span data-stu-id="f30a2-423">Events</span></span>

<span data-ttu-id="f30a2-424">Vous pouvez vous abonner et annuler l’abonnement aux événements suivants à l’aide de [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) et [removeHandlerAsync](#removehandlerasynceventtype-options-callback) , respectivement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-424">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="f30a2-425">Événement</span><span class="sxs-lookup"><span data-stu-id="f30a2-425">Event</span></span> | <span data-ttu-id="f30a2-426">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-426">Description</span></span> | <span data-ttu-id="f30a2-427">Minimale</span><span class="sxs-lookup"><span data-stu-id="f30a2-427">Minimum</span></span><br><span data-ttu-id="f30a2-428">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-428">requirement set</span></span> |
|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f30a2-429">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-429">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f30a2-430">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-430">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f30a2-431">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="f30a2-431">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f30a2-432">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-432">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f30a2-433">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="f30a2-433">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f30a2-434">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-434">1.8</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f30a2-435">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="f30a2-435">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f30a2-436">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-436">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f30a2-437">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-437">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f30a2-438">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-438">1.7</span></span> |

### <a name="example"></a><span data-ttu-id="f30a2-439">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-439">Example</span></span>

<span data-ttu-id="f30a2-440">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f30a2-440">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

## <a name="property-details"></a><span data-ttu-id="f30a2-441">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="f30a2-441">Property details</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="f30a2-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="f30a2-443">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="f30a2-443">Gets the item's attachments as an array.</span></span> <span data-ttu-id="f30a2-444">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-444">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-445">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="f30a2-445">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f30a2-446">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f30a2-446">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-447">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-447">Type</span></span>

*   <span data-ttu-id="f30a2-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-449">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-449">Requirements</span></span>

|<span data-ttu-id="f30a2-450">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-450">Requirement</span></span>|<span data-ttu-id="f30a2-451">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-451">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-452">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-453">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-453">1.0</span></span>|
|[<span data-ttu-id="f30a2-454">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-455">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-456">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-457">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-457">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-458">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-458">Example</span></span>

<span data-ttu-id="f30a2-459">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-459">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f30a2-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f30a2-461">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-461">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f30a2-462">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-462">Compose mode only.</span></span>

<span data-ttu-id="f30a2-463">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-463">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-464">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-464">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f30a2-465">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-465">Get 500 members maximum.</span></span>
- <span data-ttu-id="f30a2-466">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="f30a2-466">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-467">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-467">Type</span></span>

*   [<span data-ttu-id="f30a2-468">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f30a2-468">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f30a2-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-469">Requirements</span></span>

|<span data-ttu-id="f30a2-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-470">Requirement</span></span>|<span data-ttu-id="f30a2-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-473">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-473">1.1</span></span>|
|[<span data-ttu-id="f30a2-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-475">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-477">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-477">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-478">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="f30a2-479">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="f30a2-479">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="f30a2-480">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-480">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-481">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-481">Type</span></span>

*   [<span data-ttu-id="f30a2-482">Body</span><span class="sxs-lookup"><span data-stu-id="f30a2-482">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="f30a2-483">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-483">Requirements</span></span>

|<span data-ttu-id="f30a2-484">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-484">Requirement</span></span>|<span data-ttu-id="f30a2-485">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-486">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-487">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-487">1.1</span></span>|
|[<span data-ttu-id="f30a2-488">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-488">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-489">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-490">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-490">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-491">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-491">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-492">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-492">Example</span></span>

<span data-ttu-id="f30a2-493">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="f30a2-493">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f30a2-494">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-494">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="f30a2-495">Catégories : [catégories](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="f30a2-495">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="f30a2-496">Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-496">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-497">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-497">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-498">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-498">Type</span></span>

*   [<span data-ttu-id="f30a2-499">Categories</span><span class="sxs-lookup"><span data-stu-id="f30a2-499">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="f30a2-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-500">Requirements</span></span>

|<span data-ttu-id="f30a2-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-501">Requirement</span></span>|<span data-ttu-id="f30a2-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-504">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-504">1.8</span></span>|
|[<span data-ttu-id="f30a2-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-506">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-508">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-508">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-509">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-509">Example</span></span>

<span data-ttu-id="f30a2-510">Cet exemple obtient les catégories de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-510">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f30a2-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f30a2-512">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-512">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f30a2-513">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-513">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-514">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-514">Read mode</span></span>

<span data-ttu-id="f30a2-515">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-515">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="f30a2-516">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-516">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-517">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-517">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-518">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-518">Compose mode</span></span>

<span data-ttu-id="f30a2-519">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-519">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="f30a2-520">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-520">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-521">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-521">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f30a2-522">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-522">Get 500 members maximum.</span></span>
- <span data-ttu-id="f30a2-523">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="f30a2-523">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-524">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-524">Type</span></span>

*   <span data-ttu-id="f30a2-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-526">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-526">Requirements</span></span>

|<span data-ttu-id="f30a2-527">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-527">Requirement</span></span>|<span data-ttu-id="f30a2-528">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-529">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-530">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-530">1.0</span></span>|
|[<span data-ttu-id="f30a2-531">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-531">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-532">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-533">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-533">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-534">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-534">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f30a2-535">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="f30a2-535">(nullable) conversationId: String</span></span>

<span data-ttu-id="f30a2-536">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="f30a2-536">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f30a2-p109">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f30a2-p110">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-541">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-541">Type</span></span>

*   <span data-ttu-id="f30a2-542">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-542">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-543">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-543">Requirements</span></span>

|<span data-ttu-id="f30a2-544">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-544">Requirement</span></span>|<span data-ttu-id="f30a2-545">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-547">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-547">1.0</span></span>|
|[<span data-ttu-id="f30a2-548">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-549">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-550">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-551">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-552">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-552">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="f30a2-553">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-553">dateTimeCreated: Date</span></span>

<span data-ttu-id="f30a2-p111">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-556">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-556">Type</span></span>

*   <span data-ttu-id="f30a2-557">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-557">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-558">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-558">Requirements</span></span>

|<span data-ttu-id="f30a2-559">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-559">Requirement</span></span>|<span data-ttu-id="f30a2-560">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-561">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-562">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-562">1.0</span></span>|
|[<span data-ttu-id="f30a2-563">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-564">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-565">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-566">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-566">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-567">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-567">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="f30a2-568">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-568">dateTimeModified: Date</span></span>

<span data-ttu-id="f30a2-p112">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-571">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-571">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-572">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-572">Type</span></span>

*   <span data-ttu-id="f30a2-573">Date</span><span class="sxs-lookup"><span data-stu-id="f30a2-573">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-574">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-574">Requirements</span></span>

|<span data-ttu-id="f30a2-575">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-575">Requirement</span></span>|<span data-ttu-id="f30a2-576">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-577">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-578">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-578">1.0</span></span>|
|[<span data-ttu-id="f30a2-579">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-580">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-581">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-582">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-582">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-583">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-583">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f30a2-584">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f30a2-584">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f30a2-585">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-585">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f30a2-p113">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-588">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-588">Read mode</span></span>

<span data-ttu-id="f30a2-589">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-589">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-590">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-590">Compose mode</span></span>

<span data-ttu-id="f30a2-591">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-591">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f30a2-592">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-592">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f30a2-593">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-593">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f30a2-594">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-594">Type</span></span>

*   <span data-ttu-id="f30a2-595">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f30a2-595">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-596">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-596">Requirements</span></span>

|<span data-ttu-id="f30a2-597">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-597">Requirement</span></span>|<span data-ttu-id="f30a2-598">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-599">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-600">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-600">1.0</span></span>|
|[<span data-ttu-id="f30a2-601">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-602">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-603">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-604">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-604">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="f30a2-605">enhancedLocation : [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="f30a2-605">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="f30a2-606">Obtient ou définit les emplacements d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-606">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-607">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-607">Read mode</span></span>

<span data-ttu-id="f30a2-608">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-608">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-609">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-609">Compose mode</span></span>

<span data-ttu-id="f30a2-610">La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-610">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-611">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-611">Type</span></span>

*   [<span data-ttu-id="f30a2-612">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f30a2-612">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="f30a2-613">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-613">Requirements</span></span>

|<span data-ttu-id="f30a2-614">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-614">Requirement</span></span>|<span data-ttu-id="f30a2-615">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-616">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-617">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-617">1.8</span></span>|
|[<span data-ttu-id="f30a2-618">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-619">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-620">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-621">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-622">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-622">Example</span></span>

<span data-ttu-id="f30a2-623">L’exemple suivant obtient les emplacements actuels associés au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-623">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="f30a2-624">from : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f30a2-624">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="f30a2-625">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-625">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="f30a2-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-628">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-628">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-629">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-629">Read mode</span></span>

<span data-ttu-id="f30a2-630">La `from` propriété renvoie un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="f30a2-630">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-631">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-631">Compose mode</span></span>

<span data-ttu-id="f30a2-632">La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.</span><span class="sxs-lookup"><span data-stu-id="f30a2-632">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-633">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-633">Type</span></span>

*   <span data-ttu-id="f30a2-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f30a2-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-635">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-635">Requirements</span></span>

|<span data-ttu-id="f30a2-636">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-636">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f30a2-637">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-638">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-638">1.0</span></span>|<span data-ttu-id="f30a2-639">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-639">1.7</span></span>|
|[<span data-ttu-id="f30a2-640">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-640">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-641">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-641">ReadItem</span></span>|<span data-ttu-id="f30a2-642">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-642">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-643">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-643">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-644">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-644">Read</span></span>|<span data-ttu-id="f30a2-645">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-645">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="f30a2-646">internetHeaders : [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="f30a2-646">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="f30a2-647">Obtient ou définit les en-têtes Internet personnalisés d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-647">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="f30a2-648">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-648">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-649">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-649">Type</span></span>

*   [<span data-ttu-id="f30a2-650">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f30a2-650">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="f30a2-651">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-651">Requirements</span></span>

|<span data-ttu-id="f30a2-652">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-652">Requirement</span></span>|<span data-ttu-id="f30a2-653">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-654">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-655">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-655">1.8</span></span>|
|[<span data-ttu-id="f30a2-656">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-657">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-658">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-659">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-659">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-660">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-660">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="f30a2-661">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="f30a2-661">internetMessageId: String</span></span>

<span data-ttu-id="f30a2-p116">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-664">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-664">Type</span></span>

*   <span data-ttu-id="f30a2-665">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-665">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-666">Requirements</span></span>

|<span data-ttu-id="f30a2-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-667">Requirement</span></span>|<span data-ttu-id="f30a2-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-670">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-670">1.0</span></span>|
|[<span data-ttu-id="f30a2-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-672">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-674">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-674">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-675">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-675">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="f30a2-676">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="f30a2-676">itemClass: String</span></span>

<span data-ttu-id="f30a2-p117">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f30a2-p118">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="f30a2-681">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-681">Type</span></span>|<span data-ttu-id="f30a2-682">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-682">Description</span></span>|<span data-ttu-id="f30a2-683">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="f30a2-683">item class</span></span>|
|---|---|---|
|<span data-ttu-id="f30a2-684">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f30a2-684">Appointment items</span></span>|<span data-ttu-id="f30a2-685">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-685">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="f30a2-686">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="f30a2-686">Message items</span></span>|<span data-ttu-id="f30a2-687">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="f30a2-687">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="f30a2-688">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-688">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-689">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-689">Type</span></span>

*   <span data-ttu-id="f30a2-690">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-690">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-691">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-691">Requirements</span></span>

|<span data-ttu-id="f30a2-692">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-692">Requirement</span></span>|<span data-ttu-id="f30a2-693">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-694">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-695">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-695">1.0</span></span>|
|[<span data-ttu-id="f30a2-696">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-697">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-697">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-698">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-699">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-699">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-700">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-700">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f30a2-701">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="f30a2-701">(nullable) itemId: String</span></span>

<span data-ttu-id="f30a2-p119">Permet d’obtenir l’[identificateur de l’élément des services web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-704">L’identificateur renvoyé par la propriété `itemId` est identique à l’[identificateur d’élément des services web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="f30a2-704">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="f30a2-705">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="f30a2-705">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f30a2-706">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f30a2-706">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f30a2-707">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f30a2-707">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f30a2-p121">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-710">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-710">Type</span></span>

*   <span data-ttu-id="f30a2-711">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-711">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-712">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-712">Requirements</span></span>

|<span data-ttu-id="f30a2-713">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-713">Requirement</span></span>|<span data-ttu-id="f30a2-714">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-715">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-716">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-716">1.0</span></span>|
|[<span data-ttu-id="f30a2-717">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-718">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-719">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-720">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-720">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-721">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-721">Example</span></span>

<span data-ttu-id="f30a2-p122">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-mailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="f30a2-724">itemType : [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f30a2-724">itemType: [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f30a2-725">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="f30a2-725">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f30a2-726">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-726">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-727">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-727">Type</span></span>

*   [<span data-ttu-id="f30a2-728">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="f30a2-728">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f30a2-729">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-729">Requirements</span></span>

|<span data-ttu-id="f30a2-730">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-730">Requirement</span></span>|<span data-ttu-id="f30a2-731">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-731">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-732">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-732">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-733">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-733">1.0</span></span>|
|[<span data-ttu-id="f30a2-734">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-734">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-735">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-735">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-736">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-736">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-737">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-737">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-738">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-738">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="f30a2-739">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f30a2-739">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="f30a2-740">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-740">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-741">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-741">Read mode</span></span>

<span data-ttu-id="f30a2-742">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-742">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-743">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-743">Compose mode</span></span>

<span data-ttu-id="f30a2-744">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-744">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-745">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-745">Type</span></span>

*   <span data-ttu-id="f30a2-746">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f30a2-746">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-747">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-747">Requirements</span></span>

|<span data-ttu-id="f30a2-748">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-748">Requirement</span></span>|<span data-ttu-id="f30a2-749">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-750">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-751">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-751">1.0</span></span>|
|[<span data-ttu-id="f30a2-752">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-752">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-753">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-754">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-754">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-755">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-755">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f30a2-756">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="f30a2-756">normalizedSubject: String</span></span>

<span data-ttu-id="f30a2-p123">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f30a2-p124">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="f30a2-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-761">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-761">Type</span></span>

*   <span data-ttu-id="f30a2-762">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-762">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-763">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-763">Requirements</span></span>

|<span data-ttu-id="f30a2-764">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-764">Requirement</span></span>|<span data-ttu-id="f30a2-765">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-765">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-766">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-766">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-767">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-767">1.0</span></span>|
|[<span data-ttu-id="f30a2-768">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-768">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-769">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-769">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-770">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-770">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-771">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-771">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-772">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-772">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="f30a2-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f30a2-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="f30a2-774">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-774">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-775">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-775">Type</span></span>

*   [<span data-ttu-id="f30a2-776">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f30a2-776">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f30a2-777">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-777">Requirements</span></span>

|<span data-ttu-id="f30a2-778">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-778">Requirement</span></span>|<span data-ttu-id="f30a2-779">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-780">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-781">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-781">1.3</span></span>|
|[<span data-ttu-id="f30a2-782">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-782">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-783">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-783">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-784">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-784">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-785">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-785">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-786">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-786">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f30a2-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f30a2-788">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-788">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f30a2-789">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-789">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-790">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-790">Read mode</span></span>

<span data-ttu-id="f30a2-791">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-791">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="f30a2-792">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-792">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-793">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-793">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-794">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-794">Compose mode</span></span>

<span data-ttu-id="f30a2-795">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-795">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="f30a2-796">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-796">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-797">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-797">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f30a2-798">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-798">Get 500 members maximum.</span></span>
- <span data-ttu-id="f30a2-799">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="f30a2-799">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-800">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-800">Type</span></span>

*   <span data-ttu-id="f30a2-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-802">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-802">Requirements</span></span>

|<span data-ttu-id="f30a2-803">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-803">Requirement</span></span>|<span data-ttu-id="f30a2-804">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-804">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-805">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-805">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-806">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-806">1.0</span></span>|
|[<span data-ttu-id="f30a2-807">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-807">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-808">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-808">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-809">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-809">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-810">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-810">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="f30a2-811">Organisateur : [](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-811">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="f30a2-812">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-812">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-813">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-813">Read mode</span></span>

<span data-ttu-id="f30a2-814">La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-814">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-815">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-815">Compose mode</span></span>

<span data-ttu-id="f30a2-816">La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-816">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="f30a2-817">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-817">Type</span></span>

*   <span data-ttu-id="f30a2-818">[](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-819">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-819">Requirements</span></span>

|<span data-ttu-id="f30a2-820">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-820">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f30a2-821">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-822">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-822">1.0</span></span>|<span data-ttu-id="f30a2-823">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-823">1.7</span></span>|
|[<span data-ttu-id="f30a2-824">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-824">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-825">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-825">ReadItem</span></span>|<span data-ttu-id="f30a2-826">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-826">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-827">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-827">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-828">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-828">Read</span></span>|<span data-ttu-id="f30a2-829">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-829">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="f30a2-830">(Nullable) récurrence : [périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="f30a2-830">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="f30a2-831">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-831">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="f30a2-832">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-832">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="f30a2-833">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-833">Read and compose modes for appointment items.</span></span> <span data-ttu-id="f30a2-834">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-834">Read mode for meeting request items.</span></span>

<span data-ttu-id="f30a2-835">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="f30a2-835">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="f30a2-836">`null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="f30a2-836">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="f30a2-837">`undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-837">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="f30a2-838">Remarque : les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="f30a2-838">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="f30a2-839">Remarque : si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.</span><span class="sxs-lookup"><span data-stu-id="f30a2-839">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-840">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-840">Read mode</span></span>

<span data-ttu-id="f30a2-841">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-841">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="f30a2-842">Elle est disponible pour les rendez-vous et les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-842">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-843">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-843">Compose mode</span></span>

<span data-ttu-id="f30a2-844">La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-844">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="f30a2-845">Elle est disponible pour les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-845">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-846">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-846">Type</span></span>

* [<span data-ttu-id="f30a2-847">Instances</span><span class="sxs-lookup"><span data-stu-id="f30a2-847">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="f30a2-848">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-848">Requirement</span></span>|<span data-ttu-id="f30a2-849">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-850">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-851">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-851">1.7</span></span>|
|[<span data-ttu-id="f30a2-852">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-853">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-854">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-855">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-855">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f30a2-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f30a2-857">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-857">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f30a2-858">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-858">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-859">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-859">Read mode</span></span>

<span data-ttu-id="f30a2-860">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-860">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="f30a2-861">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-861">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-862">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-862">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-863">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-863">Compose mode</span></span>

<span data-ttu-id="f30a2-864">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-864">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="f30a2-865">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-865">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-866">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-866">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f30a2-867">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-867">Get 500 members maximum.</span></span>
- <span data-ttu-id="f30a2-868">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="f30a2-868">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-869">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-869">Type</span></span>

*   <span data-ttu-id="f30a2-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-871">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-871">Requirements</span></span>

|<span data-ttu-id="f30a2-872">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-872">Requirement</span></span>|<span data-ttu-id="f30a2-873">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-874">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-875">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-875">1.0</span></span>|
|[<span data-ttu-id="f30a2-876">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-876">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-877">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-877">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-878">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-878">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-879">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-879">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="f30a2-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f30a2-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="f30a2-p135">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f30a2-p136">Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-885">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-885">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-886">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-886">Type</span></span>

*   [<span data-ttu-id="f30a2-887">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f30a2-887">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f30a2-888">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-888">Requirements</span></span>

|<span data-ttu-id="f30a2-889">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-889">Requirement</span></span>|<span data-ttu-id="f30a2-890">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-891">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-892">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-892">1.0</span></span>|
|[<span data-ttu-id="f30a2-893">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-894">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-894">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-895">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-896">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-896">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-897">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-897">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="f30a2-898">(Nullable) seriesId : chaîne</span><span class="sxs-lookup"><span data-stu-id="f30a2-898">(nullable) seriesId: String</span></span>

<span data-ttu-id="f30a2-899">Obtient l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="f30a2-899">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="f30a2-900">Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="f30a2-900">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="f30a2-901">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-901">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-902">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="f30a2-902">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f30a2-903">La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="f30a2-903">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="f30a2-904">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f30a2-904">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f30a2-905">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="f30a2-905">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="f30a2-906">La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-906">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="f30a2-907">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-907">Type</span></span>

* <span data-ttu-id="f30a2-908">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-908">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-909">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-909">Requirements</span></span>

|<span data-ttu-id="f30a2-910">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-910">Requirement</span></span>|<span data-ttu-id="f30a2-911">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-911">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-912">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-912">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-913">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-913">1.7</span></span>|
|[<span data-ttu-id="f30a2-914">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-914">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-915">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-915">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-916">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-916">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-917">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-917">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-918">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-918">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f30a2-919">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f30a2-919">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f30a2-920">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-920">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f30a2-p139">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-923">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-923">Read mode</span></span>

<span data-ttu-id="f30a2-924">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-924">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-925">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-925">Compose mode</span></span>

<span data-ttu-id="f30a2-926">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-926">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f30a2-927">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-927">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f30a2-928">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-928">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f30a2-929">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-929">Type</span></span>

*   <span data-ttu-id="f30a2-930">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f30a2-930">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-931">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-931">Requirements</span></span>

|<span data-ttu-id="f30a2-932">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-932">Requirement</span></span>|<span data-ttu-id="f30a2-933">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-934">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-935">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-935">1.0</span></span>|
|[<span data-ttu-id="f30a2-936">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-936">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-937">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-938">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-938">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-939">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-939">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="f30a2-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f30a2-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="f30a2-941">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-941">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f30a2-942">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="f30a2-942">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-943">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-943">Read mode</span></span>

<span data-ttu-id="f30a2-p140">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="f30a2-946">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f30a2-946">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-947">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-947">Compose mode</span></span>
<span data-ttu-id="f30a2-948">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="f30a2-948">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-949">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-949">Type</span></span>

*   <span data-ttu-id="f30a2-950">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f30a2-950">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-951">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-951">Requirements</span></span>

|<span data-ttu-id="f30a2-952">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-952">Requirement</span></span>|<span data-ttu-id="f30a2-953">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-953">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-954">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-954">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-955">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-955">1.0</span></span>|
|[<span data-ttu-id="f30a2-956">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-956">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-957">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-957">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-958">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-958">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-959">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-959">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f30a2-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f30a2-961">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-961">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f30a2-962">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-962">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f30a2-963">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-963">Read mode</span></span>

<span data-ttu-id="f30a2-964">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-964">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="f30a2-965">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-965">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-966">Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-966">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f30a2-967">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-967">Compose mode</span></span>

<span data-ttu-id="f30a2-968">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-968">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="f30a2-969">Par défaut, la collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-969">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f30a2-970">Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="f30a2-970">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f30a2-971">Obtenez 500 membres au maximum.</span><span class="sxs-lookup"><span data-stu-id="f30a2-971">Get 500 members maximum.</span></span>
- <span data-ttu-id="f30a2-972">Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.</span><span class="sxs-lookup"><span data-stu-id="f30a2-972">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f30a2-973">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-973">Type</span></span>

*   <span data-ttu-id="f30a2-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f30a2-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-975">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-975">Requirements</span></span>

|<span data-ttu-id="f30a2-976">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-976">Requirement</span></span>|<span data-ttu-id="f30a2-977">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-977">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-978">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-978">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-979">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-979">1.0</span></span>|
|[<span data-ttu-id="f30a2-980">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-980">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-981">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-981">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-982">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-982">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-983">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-983">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="f30a2-984">Détails de méthodes</span><span class="sxs-lookup"><span data-stu-id="f30a2-984">Method details</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f30a2-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f30a2-986">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f30a2-986">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f30a2-987">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="f30a2-987">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f30a2-988">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-988">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-989">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-989">Parameters</span></span>
|<span data-ttu-id="f30a2-990">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-990">Name</span></span>|<span data-ttu-id="f30a2-991">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-991">Type</span></span>|<span data-ttu-id="f30a2-992">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-992">Attributes</span></span>|<span data-ttu-id="f30a2-993">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-993">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="f30a2-994">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-994">String</span></span>||<span data-ttu-id="f30a2-p144">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f30a2-997">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-997">String</span></span>||<span data-ttu-id="f30a2-p145">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f30a2-1000">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1000">Object</span></span>|<span data-ttu-id="f30a2-1001">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1002">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1003">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1003">Object</span></span>|<span data-ttu-id="f30a2-1004">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1005">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f30a2-1006">Boolean</span><span class="sxs-lookup"><span data-stu-id="f30a2-1006">Boolean</span></span>|<span data-ttu-id="f30a2-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1008">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1008">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1009">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1009">function</span></span>|<span data-ttu-id="f30a2-1010">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1011">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1011">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f30a2-1012">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1012">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f30a2-1013">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1013">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f30a2-1014">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1014">Errors</span></span>

|<span data-ttu-id="f30a2-1015">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1015">Error code</span></span>|<span data-ttu-id="f30a2-1016">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1016">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f30a2-1017">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1017">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f30a2-1018">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1018">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f30a2-1019">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1019">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1020">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1020">Requirements</span></span>

|<span data-ttu-id="f30a2-1021">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1021">Requirement</span></span>|<span data-ttu-id="f30a2-1022">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1023">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1024">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-1024">1.1</span></span>|
|[<span data-ttu-id="f30a2-1025">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1026">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1026">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1027">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1028">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1028">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1029">Exemples</span><span class="sxs-lookup"><span data-stu-id="f30a2-1029">Examples</span></span>

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

<span data-ttu-id="f30a2-1030">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1030">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="f30a2-1031">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1031">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f30a2-1032">Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1032">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f30a2-1033">La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1033">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="f30a2-1034">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1034">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="f30a2-1035">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1035">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1036">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1036">Parameters</span></span>

|<span data-ttu-id="f30a2-1037">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1037">Name</span></span>|<span data-ttu-id="f30a2-1038">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1038">Type</span></span>|<span data-ttu-id="f30a2-1039">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1039">Attributes</span></span>|<span data-ttu-id="f30a2-1040">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1040">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="f30a2-1041">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f30a2-1041">String</span></span>||<span data-ttu-id="f30a2-1042">Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1042">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="f30a2-1043">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1043">String</span></span>||<span data-ttu-id="f30a2-p147">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f30a2-1046">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1046">Object</span></span>|<span data-ttu-id="f30a2-1047">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1048">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1049">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1049">Object</span></span>|<span data-ttu-id="f30a2-1050">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1051">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f30a2-1052">Boolean</span><span class="sxs-lookup"><span data-stu-id="f30a2-1052">Boolean</span></span>|<span data-ttu-id="f30a2-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1054">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1054">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1055">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1055">function</span></span>|<span data-ttu-id="f30a2-1056">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1057">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f30a2-1058">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1058">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f30a2-1059">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1059">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f30a2-1060">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1060">Errors</span></span>

|<span data-ttu-id="f30a2-1061">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1061">Error code</span></span>|<span data-ttu-id="f30a2-1062">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1062">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f30a2-1063">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1063">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f30a2-1064">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1064">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f30a2-1065">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1065">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1066">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1066">Requirements</span></span>

|<span data-ttu-id="f30a2-1067">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1067">Requirement</span></span>|<span data-ttu-id="f30a2-1068">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1069">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1070">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1070">1.8</span></span>|
|[<span data-ttu-id="f30a2-1071">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1072">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1072">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1073">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1074">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1074">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1075">Exemples</span><span class="sxs-lookup"><span data-stu-id="f30a2-1075">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f30a2-1076">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1076">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f30a2-1077">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1077">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f30a2-1078">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1078">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1079">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1079">Parameters</span></span>

| <span data-ttu-id="f30a2-1080">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1080">Name</span></span> | <span data-ttu-id="f30a2-1081">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1081">Type</span></span> | <span data-ttu-id="f30a2-1082">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1082">Attributes</span></span> | <span data-ttu-id="f30a2-1083">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1083">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f30a2-1084">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f30a2-1084">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f30a2-1085">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1085">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f30a2-1086">Fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1086">Function</span></span> || <span data-ttu-id="f30a2-p148">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f30a2-1090">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1090">Object</span></span> | <span data-ttu-id="f30a2-1091">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1091">&lt;optional&gt;</span></span> | <span data-ttu-id="f30a2-1092">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1092">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f30a2-1093">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1093">Object</span></span> | <span data-ttu-id="f30a2-1094">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1094">&lt;optional&gt;</span></span> | <span data-ttu-id="f30a2-1095">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1095">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f30a2-1096">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1096">function</span></span>| <span data-ttu-id="f30a2-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1098">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1099">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1099">Requirements</span></span>

|<span data-ttu-id="f30a2-1100">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1100">Requirement</span></span>| <span data-ttu-id="f30a2-1101">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1101">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1102">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1102">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f30a2-1103">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-1103">1.7</span></span> |
|[<span data-ttu-id="f30a2-1104">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1104">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f30a2-1105">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1105">ReadItem</span></span> |
|[<span data-ttu-id="f30a2-1106">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1106">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f30a2-1107">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1107">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="f30a2-1108">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1108">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f30a2-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f30a2-1110">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1110">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f30a2-p149">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f30a2-1114">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1114">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f30a2-1115">Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1115">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1116">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1116">Parameters</span></span>

|<span data-ttu-id="f30a2-1117">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1117">Name</span></span>|<span data-ttu-id="f30a2-1118">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1118">Type</span></span>|<span data-ttu-id="f30a2-1119">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1119">Attributes</span></span>|<span data-ttu-id="f30a2-1120">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1120">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="f30a2-1121">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1121">String</span></span>||<span data-ttu-id="f30a2-p150">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f30a2-1124">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1124">String</span></span>||<span data-ttu-id="f30a2-1125">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1125">The subject of the item to be attached.</span></span> <span data-ttu-id="f30a2-1126">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1126">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f30a2-1127">Object</span><span class="sxs-lookup"><span data-stu-id="f30a2-1127">Object</span></span>|<span data-ttu-id="f30a2-1128">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1129">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1129">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1130">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1130">Object</span></span>|<span data-ttu-id="f30a2-1131">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1131">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1132">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1132">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1133">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1133">function</span></span>|<span data-ttu-id="f30a2-1134">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1134">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1135">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f30a2-1136">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1136">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f30a2-1137">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1137">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f30a2-1138">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1138">Errors</span></span>

|<span data-ttu-id="f30a2-1139">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1139">Error code</span></span>|<span data-ttu-id="f30a2-1140">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1140">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f30a2-1141">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1141">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1142">Requirements</span></span>

|<span data-ttu-id="f30a2-1143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1143">Requirement</span></span>|<span data-ttu-id="f30a2-1144">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1144">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1146">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-1146">1.1</span></span>|
|[<span data-ttu-id="f30a2-1147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1148">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1148">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1150">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1151">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1151">Example</span></span>

<span data-ttu-id="f30a2-1152">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1152">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="f30a2-1153">close()</span><span class="sxs-lookup"><span data-stu-id="f30a2-1153">close()</span></span>

<span data-ttu-id="f30a2-1154">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1154">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f30a2-p152">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1157">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1157">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f30a2-1158">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1158">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-1159">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1159">Requirements</span></span>

|<span data-ttu-id="f30a2-1160">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1160">Requirement</span></span>|<span data-ttu-id="f30a2-1161">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1162">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1163">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-1163">1.3</span></span>|
|[<span data-ttu-id="f30a2-1164">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1165">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f30a2-1165">Restricted</span></span>|
|[<span data-ttu-id="f30a2-1166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1167">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1167">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f30a2-1168">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1168">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f30a2-1169">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1169">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1170">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1170">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-1171">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1171">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f30a2-1172">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1172">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f30a2-p153">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1176">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1176">Parameters</span></span>

|<span data-ttu-id="f30a2-1177">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1177">Name</span></span>|<span data-ttu-id="f30a2-1178">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1178">Type</span></span>|<span data-ttu-id="f30a2-1179">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1179">Attributes</span></span>|<span data-ttu-id="f30a2-1180">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1180">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f30a2-1181">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f30a2-1181">String &#124; Object</span></span>||<span data-ttu-id="f30a2-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f30a2-1184">**OU**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1184">**OR**</span></span><br/><span data-ttu-id="f30a2-p155">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f30a2-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f30a2-1187">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1187">String</span></span>|<span data-ttu-id="f30a2-1188">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1188">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-p156">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f30a2-1191">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1191">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f30a2-1192">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1192">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1193">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1193">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f30a2-1194">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1194">String</span></span>||<span data-ttu-id="f30a2-p157">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f30a2-1197">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1197">String</span></span>||<span data-ttu-id="f30a2-1198">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1198">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f30a2-1199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f30a2-1199">String</span></span>||<span data-ttu-id="f30a2-p158">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f30a2-1202">Booléen</span><span class="sxs-lookup"><span data-stu-id="f30a2-1202">Boolean</span></span>||<span data-ttu-id="f30a2-p159">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f30a2-1205">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1205">String</span></span>||<span data-ttu-id="f30a2-p160">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1209">function</span><span class="sxs-lookup"><span data-stu-id="f30a2-1209">function</span></span>|<span data-ttu-id="f30a2-1210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1210">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1211">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1212">Requirements</span></span>

|<span data-ttu-id="f30a2-1213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1213">Requirement</span></span>|<span data-ttu-id="f30a2-1214">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1216">1.0</span></span>|
|[<span data-ttu-id="f30a2-1217">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1218">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1220">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1220">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1221">Exemples</span><span class="sxs-lookup"><span data-stu-id="f30a2-1221">Examples</span></span>

<span data-ttu-id="f30a2-1222">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1222">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f30a2-1223">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1223">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f30a2-1224">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1224">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f30a2-1225">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1225">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f30a2-1226">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1226">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f30a2-1227">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1227">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f30a2-1228">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1228">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f30a2-1229">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1229">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1230">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-1231">Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1231">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f30a2-1232">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1232">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f30a2-p161">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1236">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1236">Parameters</span></span>

|<span data-ttu-id="f30a2-1237">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1237">Name</span></span>|<span data-ttu-id="f30a2-1238">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1238">Type</span></span>|<span data-ttu-id="f30a2-1239">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1239">Attributes</span></span>|<span data-ttu-id="f30a2-1240">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1240">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f30a2-1241">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f30a2-1241">String &#124; Object</span></span>||<span data-ttu-id="f30a2-p162">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f30a2-1244">**OU**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1244">**OR**</span></span><br/><span data-ttu-id="f30a2-p163">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f30a2-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f30a2-1247">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1247">String</span></span>|<span data-ttu-id="f30a2-1248">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1248">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-p164">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f30a2-1251">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1251">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f30a2-1252">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1252">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1253">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1253">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f30a2-1254">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1254">String</span></span>||<span data-ttu-id="f30a2-p165">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f30a2-1257">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1257">String</span></span>||<span data-ttu-id="f30a2-1258">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1258">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f30a2-1259">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f30a2-1259">String</span></span>||<span data-ttu-id="f30a2-p166">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f30a2-1262">Booléen</span><span class="sxs-lookup"><span data-stu-id="f30a2-1262">Boolean</span></span>||<span data-ttu-id="f30a2-p167">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f30a2-1265">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1265">String</span></span>||<span data-ttu-id="f30a2-p168">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1269">function</span><span class="sxs-lookup"><span data-stu-id="f30a2-1269">function</span></span>|<span data-ttu-id="f30a2-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1271">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1272">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1272">Requirements</span></span>

|<span data-ttu-id="f30a2-1273">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1273">Requirement</span></span>|<span data-ttu-id="f30a2-1274">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1274">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1275">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1275">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1276">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1276">1.0</span></span>|
|[<span data-ttu-id="f30a2-1277">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1277">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1278">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1279">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1279">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1280">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1280">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1281">Exemples</span><span class="sxs-lookup"><span data-stu-id="f30a2-1281">Examples</span></span>

<span data-ttu-id="f30a2-1282">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1282">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f30a2-1283">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1283">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f30a2-1284">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1284">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f30a2-1285">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1285">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f30a2-1286">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1286">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f30a2-1287">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1287">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="f30a2-1288">getAllInternetHeadersAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1288">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="f30a2-1289">Obtient tous les en-têtes Internet pour le message sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1289">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="f30a2-1290">Mode Lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1290">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1291">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1291">Parameters</span></span>

|<span data-ttu-id="f30a2-1292">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1292">Name</span></span>|<span data-ttu-id="f30a2-1293">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1293">Type</span></span>|<span data-ttu-id="f30a2-1294">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1294">Attributes</span></span>|<span data-ttu-id="f30a2-1295">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1295">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1296">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1296">Object</span></span>|<span data-ttu-id="f30a2-1297">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1297">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1298">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1298">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1299">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1299">Object</span></span>|<span data-ttu-id="f30a2-1300">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1301">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1301">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1302">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1302">function</span></span>|<span data-ttu-id="f30a2-1303">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1304">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1304">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="f30a2-1305">En cas de réussite, les données des en-têtes Internet sont fournies dans la propriété asyncResult. Value sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1305">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="f30a2-1306">Reportez-vous à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183) pour les informations de mise en forme de la valeur de chaîne renvoyée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1306">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="f30a2-1307">En cas d’échec de l’appel, la propriété asyncResult. Error contient un code d’erreur correspondant à la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1307">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1308">Requirements</span></span>

|<span data-ttu-id="f30a2-1309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1309">Requirement</span></span>|<span data-ttu-id="f30a2-1310">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1310">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1312">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1312">1.8</span></span>|
|[<span data-ttu-id="f30a2-1313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1314">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1316">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1316">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1317">Returns:</span></span>

<span data-ttu-id="f30a2-1318">Les données des en-têtes Internet sous forme de chaîne formatée conformément à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1318">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="f30a2-1319">Type : String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1319">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1320">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1320">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="f30a2-1321">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1321">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="f30a2-1322">Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1322">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="f30a2-1323">La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1323">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f30a2-1324">Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="f30a2-1324">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="f30a2-1325">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1325">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f30a2-1326">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1326">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1327">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1327">Parameters</span></span>

|<span data-ttu-id="f30a2-1328">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1328">Name</span></span>|<span data-ttu-id="f30a2-1329">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1329">Type</span></span>|<span data-ttu-id="f30a2-1330">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1330">Attributes</span></span>|<span data-ttu-id="f30a2-1331">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1331">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f30a2-1332">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1332">String</span></span>||<span data-ttu-id="f30a2-1333">Identificateur de la pièce jointe que vous souhaitez obtenir.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1333">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="f30a2-1334">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1334">Object</span></span>|<span data-ttu-id="f30a2-1335">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1335">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1336">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1336">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1337">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1337">Object</span></span>|<span data-ttu-id="f30a2-1338">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1338">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1339">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1339">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1340">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1340">function</span></span>|<span data-ttu-id="f30a2-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1342">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1343">Requirements</span></span>

|<span data-ttu-id="f30a2-1344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1344">Requirement</span></span>|<span data-ttu-id="f30a2-1345">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1347">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1347">1.8</span></span>|
|[<span data-ttu-id="f30a2-1348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1349">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1351">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1351">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1352">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1352">Returns:</span></span>

<span data-ttu-id="f30a2-1353">Type : [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1353">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1354">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1354">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="f30a2-1355">getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-1355">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="f30a2-1356">Obtient les pièces jointes de l’élément sous la forme d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1356">Gets the item's attachments as an array.</span></span> <span data-ttu-id="f30a2-1357">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1357">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1358">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1358">Parameters</span></span>

|<span data-ttu-id="f30a2-1359">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1359">Name</span></span>|<span data-ttu-id="f30a2-1360">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1360">Type</span></span>|<span data-ttu-id="f30a2-1361">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1361">Attributes</span></span>|<span data-ttu-id="f30a2-1362">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1362">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1363">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1363">Object</span></span>|<span data-ttu-id="f30a2-1364">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1364">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1365">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1365">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1366">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1366">Object</span></span>|<span data-ttu-id="f30a2-1367">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1367">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1368">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1368">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1369">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1369">function</span></span>|<span data-ttu-id="f30a2-1370">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1371">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1371">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1372">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1372">Requirements</span></span>

|<span data-ttu-id="f30a2-1373">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1373">Requirement</span></span>|<span data-ttu-id="f30a2-1374">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1374">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1375">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1376">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1376">1.8</span></span>|
|[<span data-ttu-id="f30a2-1377">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1378">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1380">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1380">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1381">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1381">Returns:</span></span>

<span data-ttu-id="f30a2-1382">Type : Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f30a2-1382">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1383">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1383">Example</span></span>

<span data-ttu-id="f30a2-1384">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1384">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f30a2-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f30a2-1386">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1386">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1387">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1387">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-1388">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1388">Requirements</span></span>

|<span data-ttu-id="f30a2-1389">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1389">Requirement</span></span>|<span data-ttu-id="f30a2-1390">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1390">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1391">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1392">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1392">1.0</span></span>|
|[<span data-ttu-id="f30a2-1393">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1393">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1394">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1395">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1395">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1396">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1396">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1397">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1397">Returns:</span></span>

<span data-ttu-id="f30a2-1398">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1398">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1399">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1399">Example</span></span>

<span data-ttu-id="f30a2-1400">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1400">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f30a2-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f30a2-1402">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1402">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1403">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1403">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1404">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1404">Parameters</span></span>

|<span data-ttu-id="f30a2-1405">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1405">Name</span></span>|<span data-ttu-id="f30a2-1406">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1406">Type</span></span>|<span data-ttu-id="f30a2-1407">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1407">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="f30a2-1408">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f30a2-1408">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="f30a2-1409">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1409">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1410">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1410">Requirements</span></span>

|<span data-ttu-id="f30a2-1411">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1411">Requirement</span></span>|<span data-ttu-id="f30a2-1412">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1412">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1413">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1414">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1414">1.0</span></span>|
|[<span data-ttu-id="f30a2-1415">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1416">Restreinte</span><span class="sxs-lookup"><span data-stu-id="f30a2-1416">Restricted</span></span>|
|[<span data-ttu-id="f30a2-1417">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1418">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1418">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1419">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1419">Returns:</span></span>

<span data-ttu-id="f30a2-1420">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1420">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f30a2-1421">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1421">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f30a2-1422">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1422">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f30a2-1423">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1423">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="f30a2-1424">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f30a2-1424">Value of `entityType`</span></span>|<span data-ttu-id="f30a2-1425">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="f30a2-1425">Type of objects in returned array</span></span>|<span data-ttu-id="f30a2-1426">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="f30a2-1426">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="f30a2-1427">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1427">String</span></span>|<span data-ttu-id="f30a2-1428">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1428">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="f30a2-1429">Contact</span><span class="sxs-lookup"><span data-stu-id="f30a2-1429">Contact</span></span>|<span data-ttu-id="f30a2-1430">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1430">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="f30a2-1431">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1431">String</span></span>|<span data-ttu-id="f30a2-1432">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1432">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="f30a2-1433">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f30a2-1433">MeetingSuggestion</span></span>|<span data-ttu-id="f30a2-1434">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1434">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="f30a2-1435">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f30a2-1435">PhoneNumber</span></span>|<span data-ttu-id="f30a2-1436">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1436">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="f30a2-1437">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f30a2-1437">TaskSuggestion</span></span>|<span data-ttu-id="f30a2-1438">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1438">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="f30a2-1439">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1439">String</span></span>|<span data-ttu-id="f30a2-1440">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f30a2-1440">**Restricted**</span></span>|

<span data-ttu-id="f30a2-1441">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f30a2-1441">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1442">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1442">Example</span></span>

<span data-ttu-id="f30a2-1443">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1443">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f30a2-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f30a2-1445">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1445">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1446">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1446">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-1447">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1447">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1448">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1448">Parameters</span></span>

|<span data-ttu-id="f30a2-1449">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1449">Name</span></span>|<span data-ttu-id="f30a2-1450">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1450">Type</span></span>|<span data-ttu-id="f30a2-1451">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1451">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f30a2-1452">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f30a2-1452">String</span></span>|<span data-ttu-id="f30a2-1453">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1453">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1454">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1454">Requirements</span></span>

|<span data-ttu-id="f30a2-1455">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1455">Requirement</span></span>|<span data-ttu-id="f30a2-1456">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1456">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1457">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1457">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1458">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1458">1.0</span></span>|
|[<span data-ttu-id="f30a2-1459">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1459">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1460">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1460">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1461">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1461">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1462">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1462">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1463">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1463">Returns:</span></span>

<span data-ttu-id="f30a2-p174">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f30a2-1466">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f30a2-1466">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="f30a2-1467">getInitializationContextAsync ([options], [Rappel])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1467">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="f30a2-1468">Obtient les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1468">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1469">Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure sur Windows (versions « démarrer en un clic » ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1469">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1470">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1470">Parameters</span></span>

|<span data-ttu-id="f30a2-1471">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1471">Name</span></span>|<span data-ttu-id="f30a2-1472">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1472">Type</span></span>|<span data-ttu-id="f30a2-1473">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1473">Attributes</span></span>|<span data-ttu-id="f30a2-1474">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1474">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1475">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1475">Object</span></span>|<span data-ttu-id="f30a2-1476">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1476">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1477">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1477">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1478">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1478">Object</span></span>|<span data-ttu-id="f30a2-1479">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1479">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1480">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1480">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1481">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1481">function</span></span>|<span data-ttu-id="f30a2-1482">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1482">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1483">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1483">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f30a2-1484">En cas de réussite, les données d’initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1484">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="f30a2-1485">S’il n’existe pas de contexte d’initialisation `asyncResult` , l’objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1485">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1486">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1486">Requirements</span></span>

|<span data-ttu-id="f30a2-1487">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1487">Requirement</span></span>|<span data-ttu-id="f30a2-1488">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1489">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1490">Aperçu</span><span class="sxs-lookup"><span data-stu-id="f30a2-1490">Preview</span></span>|
|[<span data-ttu-id="f30a2-1491">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1492">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1493">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1494">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1494">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1495">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1495">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="f30a2-1496">getItemIdAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1496">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="f30a2-1497">Obtient de manière asynchrone l’ID d’un élément enregistré.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1497">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="f30a2-1498">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1498">Compose mode only.</span></span>

<span data-ttu-id="f30a2-1499">Lorsqu’elle est appelée, cette méthode renvoie l’ID de l’élément par le biais de la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1499">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1500">Si votre complément appelle `getItemIdAsync` sur un élément en mode composition (par exemple, pour obtenir un à utiliser avec `itemId` EWS ou l’API REST), sachez que lorsque Outlook est en mode mis en cache, l’élément peut prendre un certain temps avant la synchronisation de l’élément avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1500">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="f30a2-1501">Tant que l’élément n’est pas synchronisé `itemId` , le n’est pas reconnu et son utilisation renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1501">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1502">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1502">Parameters</span></span>

|<span data-ttu-id="f30a2-1503">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1503">Name</span></span>|<span data-ttu-id="f30a2-1504">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1504">Type</span></span>|<span data-ttu-id="f30a2-1505">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1505">Attributes</span></span>|<span data-ttu-id="f30a2-1506">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1506">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1507">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1507">Object</span></span>|<span data-ttu-id="f30a2-1508">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1508">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1509">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1509">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1510">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1510">Object</span></span>|<span data-ttu-id="f30a2-1511">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1511">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1512">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1512">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1513">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1513">function</span></span>||<span data-ttu-id="f30a2-1514">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1514">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f30a2-1515">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1515">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f30a2-1516">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1516">Errors</span></span>

|<span data-ttu-id="f30a2-1517">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1517">Error code</span></span>|<span data-ttu-id="f30a2-1518">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1518">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="f30a2-1519">L’ID ne peut pas être récupéré tant que l’élément n’est pas enregistré.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1519">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1520">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1520">Requirements</span></span>

|<span data-ttu-id="f30a2-1521">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1521">Requirement</span></span>|<span data-ttu-id="f30a2-1522">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1522">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1523">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1524">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1524">1.8</span></span>|
|[<span data-ttu-id="f30a2-1525">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1526">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1527">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1528">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1528">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1529">Exemples</span><span class="sxs-lookup"><span data-stu-id="f30a2-1529">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f30a2-1530">L’exemple suivant montre la structure du `result` paramètre transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1530">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="f30a2-1531">La `value` propriété contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1531">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="f30a2-1532">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1532">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f30a2-1533">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1533">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1534">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1534">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-p178">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f30a2-1538">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1538">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f30a2-1539">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1539">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f30a2-p179">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-1543">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1543">Requirements</span></span>

|<span data-ttu-id="f30a2-1544">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1544">Requirement</span></span>|<span data-ttu-id="f30a2-1545">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1547">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1547">1.0</span></span>|
|[<span data-ttu-id="f30a2-1548">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1549">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1550">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1551">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1551">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1552">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1552">Returns:</span></span>

<span data-ttu-id="f30a2-p180">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f30a2-1555">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f30a2-1555">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f30a2-1556">Object</span><span class="sxs-lookup"><span data-stu-id="f30a2-1556">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f30a2-1557">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1557">Example</span></span>

<span data-ttu-id="f30a2-1558">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1558">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f30a2-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f30a2-1560">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1560">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1561">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1561">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-1562">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1562">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f30a2-p181">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1565">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1565">Parameters</span></span>

|<span data-ttu-id="f30a2-1566">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1566">Name</span></span>|<span data-ttu-id="f30a2-1567">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1567">Type</span></span>|<span data-ttu-id="f30a2-1568">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1568">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f30a2-1569">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1569">String</span></span>|<span data-ttu-id="f30a2-1570">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1570">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1571">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1571">Requirements</span></span>

|<span data-ttu-id="f30a2-1572">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1572">Requirement</span></span>|<span data-ttu-id="f30a2-1573">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1573">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1574">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1575">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1575">1.0</span></span>|
|[<span data-ttu-id="f30a2-1576">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1577">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1578">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1579">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1579">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1580">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1580">Returns:</span></span>

<span data-ttu-id="f30a2-1581">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1581">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="f30a2-1582">Type : Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f30a2-1582">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1583">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1583">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f30a2-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f30a2-1585">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1585">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f30a2-p182">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie une chaîne vide pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1588">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1588">Parameters</span></span>

|<span data-ttu-id="f30a2-1589">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1589">Name</span></span>|<span data-ttu-id="f30a2-1590">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1590">Type</span></span>|<span data-ttu-id="f30a2-1591">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1591">Attributes</span></span>|<span data-ttu-id="f30a2-1592">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1592">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="f30a2-1593">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f30a2-1593">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f30a2-p183">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="f30a2-1597">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1597">Object</span></span>|<span data-ttu-id="f30a2-1598">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1599">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1599">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1600">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1600">Object</span></span>|<span data-ttu-id="f30a2-1601">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1602">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1602">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1603">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1603">function</span></span>||<span data-ttu-id="f30a2-1604">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1604">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f30a2-1605">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1605">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f30a2-1606">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1606">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1607">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1607">Requirements</span></span>

|<span data-ttu-id="f30a2-1608">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1608">Requirement</span></span>|<span data-ttu-id="f30a2-1609">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1609">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1610">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1611">1.2</span><span class="sxs-lookup"><span data-stu-id="f30a2-1611">1.2</span></span>|
|[<span data-ttu-id="f30a2-1612">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1613">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1614">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1615">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1615">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1616">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1616">Returns:</span></span>

<span data-ttu-id="f30a2-1617">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1617">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="f30a2-1618">Type : String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1618">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1619">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1619">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f30a2-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f30a2-1621">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1621">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="f30a2-1622">Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1622">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1623">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1623">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-1624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1624">Requirements</span></span>

|<span data-ttu-id="f30a2-1625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1625">Requirement</span></span>|<span data-ttu-id="f30a2-1626">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1626">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1628">1.6</span><span class="sxs-lookup"><span data-stu-id="f30a2-1628">1.6</span></span>|
|[<span data-ttu-id="f30a2-1629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1630">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1632">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1632">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1633">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1633">Returns:</span></span>

<span data-ttu-id="f30a2-1634">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1634">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1635">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1635">Example</span></span>

<span data-ttu-id="f30a2-1636">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1636">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f30a2-1637">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f30a2-1637">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f30a2-p186">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f30a2-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1640">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f30a2-p187">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f30a2-1644">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1644">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f30a2-1645">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1645">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f30a2-p188">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f30a2-1649">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1649">Requirements</span></span>

|<span data-ttu-id="f30a2-1650">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1650">Requirement</span></span>|<span data-ttu-id="f30a2-1651">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1651">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1652">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1653">1.6</span><span class="sxs-lookup"><span data-stu-id="f30a2-1653">1.6</span></span>|
|[<span data-ttu-id="f30a2-1654">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1655">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1656">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1657">Lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1657">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f30a2-1658">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1658">Returns:</span></span>

<span data-ttu-id="f30a2-p189">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f30a2-1661">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1661">Example</span></span>

<span data-ttu-id="f30a2-1662">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1662">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="f30a2-1663">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1663">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="f30a2-1664">Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1664">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1665">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1665">Parameters</span></span>

|<span data-ttu-id="f30a2-1666">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1666">Name</span></span>|<span data-ttu-id="f30a2-1667">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1667">Type</span></span>|<span data-ttu-id="f30a2-1668">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1668">Attributes</span></span>|<span data-ttu-id="f30a2-1669">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1669">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1670">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1670">Object</span></span>|<span data-ttu-id="f30a2-1671">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1671">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1672">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1672">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1673">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1673">Object</span></span>|<span data-ttu-id="f30a2-1674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1674">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1675">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1675">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1676">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1676">function</span></span>||<span data-ttu-id="f30a2-1677">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f30a2-1678">Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1678">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f30a2-1679">Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1679">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1680">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1680">Requirements</span></span>

|<span data-ttu-id="f30a2-1681">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1681">Requirement</span></span>|<span data-ttu-id="f30a2-1682">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1682">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1683">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1684">1.8</span><span class="sxs-lookup"><span data-stu-id="f30a2-1684">1.8</span></span>|
|[<span data-ttu-id="f30a2-1685">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1686">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1687">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1688">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1688">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1689">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1689">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f30a2-1690">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1690">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f30a2-1691">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1691">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f30a2-p191">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1695">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1695">Parameters</span></span>

|<span data-ttu-id="f30a2-1696">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1696">Name</span></span>|<span data-ttu-id="f30a2-1697">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1697">Type</span></span>|<span data-ttu-id="f30a2-1698">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1698">Attributes</span></span>|<span data-ttu-id="f30a2-1699">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1699">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="f30a2-1700">function</span><span class="sxs-lookup"><span data-stu-id="f30a2-1700">function</span></span>||<span data-ttu-id="f30a2-1701">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f30a2-1702">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1702">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f30a2-1703">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1703">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="f30a2-1704">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1704">Object</span></span>|<span data-ttu-id="f30a2-1705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1705">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1706">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1706">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f30a2-1707">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1707">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1708">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1708">Requirements</span></span>

|<span data-ttu-id="f30a2-1709">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1709">Requirement</span></span>|<span data-ttu-id="f30a2-1710">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1710">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1711">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1712">1.0</span><span class="sxs-lookup"><span data-stu-id="f30a2-1712">1.0</span></span>|
|[<span data-ttu-id="f30a2-1713">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1714">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1714">ReadItem</span></span>|
|[<span data-ttu-id="f30a2-1715">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1716">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1716">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1717">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1717">Example</span></span>

<span data-ttu-id="f30a2-p194">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f30a2-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f30a2-1722">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1722">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f30a2-1723">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1723">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f30a2-1724">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1724">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f30a2-1725">Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1725">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f30a2-1726">Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1726">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1727">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1727">Parameters</span></span>

|<span data-ttu-id="f30a2-1728">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1728">Name</span></span>|<span data-ttu-id="f30a2-1729">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1729">Type</span></span>|<span data-ttu-id="f30a2-1730">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1730">Attributes</span></span>|<span data-ttu-id="f30a2-1731">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1731">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f30a2-1732">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1732">String</span></span>||<span data-ttu-id="f30a2-1733">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1733">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="f30a2-1734">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1734">Object</span></span>|<span data-ttu-id="f30a2-1735">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1735">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1736">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1736">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1737">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1737">Object</span></span>|<span data-ttu-id="f30a2-1738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1738">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1739">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1739">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1740">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1740">function</span></span>|<span data-ttu-id="f30a2-1741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1741">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1742">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1742">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f30a2-1743">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1743">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f30a2-1744">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1744">Errors</span></span>

|<span data-ttu-id="f30a2-1745">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1745">Error code</span></span>|<span data-ttu-id="f30a2-1746">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1746">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="f30a2-1747">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1747">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1748">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1748">Requirements</span></span>

|<span data-ttu-id="f30a2-1749">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1749">Requirement</span></span>|<span data-ttu-id="f30a2-1750">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1750">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1751">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1752">1.1</span><span class="sxs-lookup"><span data-stu-id="f30a2-1752">1.1</span></span>|
|[<span data-ttu-id="f30a2-1753">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1754">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1754">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1755">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1756">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1756">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1757">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1757">Example</span></span>

<span data-ttu-id="f30a2-1758">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="f30a2-1758">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="f30a2-1759">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f30a2-1759">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="f30a2-1760">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1760">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="f30a2-1761">Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1761">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1762">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1762">Parameters</span></span>

| <span data-ttu-id="f30a2-1763">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1763">Name</span></span> | <span data-ttu-id="f30a2-1764">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1764">Type</span></span> | <span data-ttu-id="f30a2-1765">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1765">Attributes</span></span> | <span data-ttu-id="f30a2-1766">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1766">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f30a2-1767">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f30a2-1767">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f30a2-1768">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1768">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="f30a2-1769">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1769">Object</span></span> | <span data-ttu-id="f30a2-1770">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1770">&lt;optional&gt;</span></span> | <span data-ttu-id="f30a2-1771">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1771">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f30a2-1772">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1772">Object</span></span> | <span data-ttu-id="f30a2-1773">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1773">&lt;optional&gt;</span></span> | <span data-ttu-id="f30a2-1774">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1774">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f30a2-1775">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1775">function</span></span>| <span data-ttu-id="f30a2-1776">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1776">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1777">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1778">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1778">Requirements</span></span>

|<span data-ttu-id="f30a2-1779">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1779">Requirement</span></span>| <span data-ttu-id="f30a2-1780">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1780">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1781">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f30a2-1782">1.7</span><span class="sxs-lookup"><span data-stu-id="f30a2-1782">1.7</span></span> |
|[<span data-ttu-id="f30a2-1783">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f30a2-1784">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1784">ReadItem</span></span> |
|[<span data-ttu-id="f30a2-1785">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f30a2-1786">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f30a2-1786">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f30a2-1787">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1787">saveAsync([options], callback)</span></span>

<span data-ttu-id="f30a2-1788">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1788">Asynchronously saves an item.</span></span>

<span data-ttu-id="f30a2-1789">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1789">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f30a2-1790">Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1790">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f30a2-1791">Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1791">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1792">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1792">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f30a2-1793">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1793">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f30a2-p198">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f30a2-1797">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="f30a2-1797">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f30a2-1798">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1798">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f30a2-1799">La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1799">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f30a2-1800">Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1800">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f30a2-1801">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1801">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1802">Parameters</span><span class="sxs-lookup"><span data-stu-id="f30a2-1802">Parameters</span></span>

|<span data-ttu-id="f30a2-1803">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1803">Name</span></span>|<span data-ttu-id="f30a2-1804">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1804">Type</span></span>|<span data-ttu-id="f30a2-1805">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1805">Attributes</span></span>|<span data-ttu-id="f30a2-1806">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1806">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f30a2-1807">Object</span><span class="sxs-lookup"><span data-stu-id="f30a2-1807">Object</span></span>|<span data-ttu-id="f30a2-1808">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1808">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1809">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1809">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1810">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1810">Object</span></span>|<span data-ttu-id="f30a2-1811">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1811">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1812">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1812">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1813">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1813">function</span></span>||<span data-ttu-id="f30a2-1814">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1814">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f30a2-1815">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1815">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1816">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1816">Requirements</span></span>

|<span data-ttu-id="f30a2-1817">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1817">Requirement</span></span>|<span data-ttu-id="f30a2-1818">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1818">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1819">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1819">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1820">1.3</span><span class="sxs-lookup"><span data-stu-id="f30a2-1820">1.3</span></span>|
|[<span data-ttu-id="f30a2-1821">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1821">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1822">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1822">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1823">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1823">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1824">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1824">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f30a2-1825">範例</span><span class="sxs-lookup"><span data-stu-id="f30a2-1825">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f30a2-p200">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f30a2-1828">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f30a2-1828">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f30a2-1829">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1829">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f30a2-p201">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f30a2-1833">Paramètres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1833">Parameters</span></span>

|<span data-ttu-id="f30a2-1834">Nom</span><span class="sxs-lookup"><span data-stu-id="f30a2-1834">Name</span></span>|<span data-ttu-id="f30a2-1835">Type</span><span class="sxs-lookup"><span data-stu-id="f30a2-1835">Type</span></span>|<span data-ttu-id="f30a2-1836">Attributs</span><span class="sxs-lookup"><span data-stu-id="f30a2-1836">Attributes</span></span>|<span data-ttu-id="f30a2-1837">Description</span><span class="sxs-lookup"><span data-stu-id="f30a2-1837">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="f30a2-1838">String</span><span class="sxs-lookup"><span data-stu-id="f30a2-1838">String</span></span>||<span data-ttu-id="f30a2-p202">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="f30a2-1842">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1842">Object</span></span>|<span data-ttu-id="f30a2-1843">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1843">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1844">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1844">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f30a2-1845">Objet</span><span class="sxs-lookup"><span data-stu-id="f30a2-1845">Object</span></span>|<span data-ttu-id="f30a2-1846">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1846">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1847">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1847">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f30a2-1848">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f30a2-1848">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f30a2-1849">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f30a2-1849">&lt;optional&gt;</span></span>|<span data-ttu-id="f30a2-1850">Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1850">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f30a2-1851">Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1851">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f30a2-1852">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1852">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f30a2-1853">Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1853">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f30a2-1854">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="f30a2-1854">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="f30a2-1855">fonction</span><span class="sxs-lookup"><span data-stu-id="f30a2-1855">function</span></span>||<span data-ttu-id="f30a2-1856">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f30a2-1856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f30a2-1857">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f30a2-1857">Requirements</span></span>

|<span data-ttu-id="f30a2-1858">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f30a2-1858">Requirement</span></span>|<span data-ttu-id="f30a2-1859">Valeur</span><span class="sxs-lookup"><span data-stu-id="f30a2-1859">Value</span></span>|
|---|---|
|[<span data-ttu-id="f30a2-1860">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f30a2-1860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f30a2-1861">1.2</span><span class="sxs-lookup"><span data-stu-id="f30a2-1861">1.2</span></span>|
|[<span data-ttu-id="f30a2-1862">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f30a2-1862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f30a2-1863">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f30a2-1863">ReadWriteItem</span></span>|
|[<span data-ttu-id="f30a2-1864">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f30a2-1864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f30a2-1865">Composition</span><span class="sxs-lookup"><span data-stu-id="f30a2-1865">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f30a2-1866">Exemple</span><span class="sxs-lookup"><span data-stu-id="f30a2-1866">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
