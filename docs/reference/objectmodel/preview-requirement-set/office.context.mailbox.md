---
title: Office. Context. Mailbox-Preview-ensemble de conditions requises
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 864c4f2931762ff6d8a02abb8da1a03e1abcab80
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670117"
---
# <a name="mailbox"></a><span data-ttu-id="dc38c-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="dc38c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="dc38c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="dc38c-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc38c-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc38c-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-105">Requirements</span></span>

|<span data-ttu-id="dc38c-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-106">Requirement</span></span>| <span data-ttu-id="dc38c-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-109">1.0</span></span>|
|[<span data-ttu-id="dc38c-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dc38c-111">Restricted</span></span>|
|[<span data-ttu-id="dc38c-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dc38c-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="dc38c-114">Properties</span></span>

| <span data-ttu-id="dc38c-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="dc38c-115">Property</span></span> | <span data-ttu-id="dc38c-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc38c-116">Minimum</span></span><br><span data-ttu-id="dc38c-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="dc38c-117">permission level</span></span> | <span data-ttu-id="dc38c-118">Modes</span><span class="sxs-lookup"><span data-stu-id="dc38c-118">Modes</span></span> | <span data-ttu-id="dc38c-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dc38c-119">Return type</span></span> | <span data-ttu-id="dc38c-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc38c-120">Minimum</span></span><br><span data-ttu-id="dc38c-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="dc38c-122">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="dc38c-122">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="dc38c-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-123">ReadItem</span></span> | <span data-ttu-id="dc38c-124">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-124">Compose</span></span><br><span data-ttu-id="dc38c-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-125">Read</span></span> | <span data-ttu-id="dc38c-126">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-126">String</span></span> | <span data-ttu-id="dc38c-127">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-127">1.0</span></span> |
| [<span data-ttu-id="dc38c-128">masterCategories</span><span class="sxs-lookup"><span data-stu-id="dc38c-128">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="dc38c-129">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dc38c-129">ReadWriteMailbox</span></span> | <span data-ttu-id="dc38c-130">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-130">Compose</span></span><br><span data-ttu-id="dc38c-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-131">Read</span></span> | [<span data-ttu-id="dc38c-132">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="dc38c-132">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories) | <span data-ttu-id="dc38c-133">1.8</span><span class="sxs-lookup"><span data-stu-id="dc38c-133">1.8</span></span> |
| [<span data-ttu-id="dc38c-134">restUrl</span><span class="sxs-lookup"><span data-stu-id="dc38c-134">restUrl</span></span>](#resturl-string) | <span data-ttu-id="dc38c-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-135">ReadItem</span></span> | <span data-ttu-id="dc38c-136">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-136">Compose</span></span><br><span data-ttu-id="dc38c-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-137">Read</span></span> | <span data-ttu-id="dc38c-138">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-138">String</span></span> | <span data-ttu-id="dc38c-139">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-139">1.5</span></span> |

##### <a name="methods"></a><span data-ttu-id="dc38c-140">Méthodes</span><span class="sxs-lookup"><span data-stu-id="dc38c-140">Methods</span></span>

| <span data-ttu-id="dc38c-141">Méthode</span><span class="sxs-lookup"><span data-stu-id="dc38c-141">Method</span></span> | <span data-ttu-id="dc38c-142">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc38c-142">Minimum</span></span><br><span data-ttu-id="dc38c-143">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="dc38c-143">permission level</span></span> | <span data-ttu-id="dc38c-144">Modes</span><span class="sxs-lookup"><span data-stu-id="dc38c-144">Modes</span></span> | <span data-ttu-id="dc38c-145">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc38c-145">Minimum</span></span><br><span data-ttu-id="dc38c-146">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-146">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="dc38c-147">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-147">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="dc38c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-148">ReadItem</span></span> | <span data-ttu-id="dc38c-149">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-149">Compose</span></span><br><span data-ttu-id="dc38c-150">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-150">Read</span></span> | <span data-ttu-id="dc38c-151">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-151">1.5</span></span> |
| [<span data-ttu-id="dc38c-152">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="dc38c-152">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="dc38c-153">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dc38c-153">Restricted</span></span> | <span data-ttu-id="dc38c-154">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-154">Compose</span></span><br><span data-ttu-id="dc38c-155">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-155">Read</span></span> | <span data-ttu-id="dc38c-156">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-156">1.3</span></span> |
| [<span data-ttu-id="dc38c-157">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dc38c-157">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="dc38c-158">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-158">ReadItem</span></span> | <span data-ttu-id="dc38c-159">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-159">Compose</span></span><br><span data-ttu-id="dc38c-160">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-160">Read</span></span> | <span data-ttu-id="dc38c-161">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-161">1.0</span></span> |
| [<span data-ttu-id="dc38c-162">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="dc38c-162">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="dc38c-163">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dc38c-163">Restricted</span></span> | <span data-ttu-id="dc38c-164">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-164">Compose</span></span><br><span data-ttu-id="dc38c-165">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-165">Read</span></span> | <span data-ttu-id="dc38c-166">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-166">1.3</span></span> |
| [<span data-ttu-id="dc38c-167">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="dc38c-167">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="dc38c-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-168">ReadItem</span></span> | <span data-ttu-id="dc38c-169">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-169">Compose</span></span><br><span data-ttu-id="dc38c-170">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-170">Read</span></span> | <span data-ttu-id="dc38c-171">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-171">1.0</span></span> |
| [<span data-ttu-id="dc38c-172">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dc38c-172">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="dc38c-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-173">ReadItem</span></span> | <span data-ttu-id="dc38c-174">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-174">Compose</span></span><br><span data-ttu-id="dc38c-175">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-175">Read</span></span> | <span data-ttu-id="dc38c-176">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-176">1.0</span></span> |
| [<span data-ttu-id="dc38c-177">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="dc38c-177">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="dc38c-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-178">ReadItem</span></span> | <span data-ttu-id="dc38c-179">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-179">Compose</span></span><br><span data-ttu-id="dc38c-180">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-180">Read</span></span> | <span data-ttu-id="dc38c-181">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-181">1.0</span></span> |
| [<span data-ttu-id="dc38c-182">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dc38c-182">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="dc38c-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-183">ReadItem</span></span> | <span data-ttu-id="dc38c-184">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-184">Read</span></span> | <span data-ttu-id="dc38c-185">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-185">1.0</span></span> |
| [<span data-ttu-id="dc38c-186">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="dc38c-186">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="dc38c-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-187">ReadItem</span></span> | <span data-ttu-id="dc38c-188">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-188">Compose</span></span><br><span data-ttu-id="dc38c-189">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-189">Read</span></span> | <span data-ttu-id="dc38c-190">1.6</span><span class="sxs-lookup"><span data-stu-id="dc38c-190">1.6</span></span> |
| [<span data-ttu-id="dc38c-191">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-191">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="dc38c-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-192">ReadItem</span></span> | <span data-ttu-id="dc38c-193">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-193">Compose</span></span><br><span data-ttu-id="dc38c-194">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-194">Read</span></span> | <span data-ttu-id="dc38c-195">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-195">1.5</span></span> |
| [<span data-ttu-id="dc38c-196">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-196">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="dc38c-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-197">ReadItem</span></span> | <span data-ttu-id="dc38c-198">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-198">Compose</span></span><br><span data-ttu-id="dc38c-199">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-199">Read</span></span> | <span data-ttu-id="dc38c-200">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-200">1.3</span></span><br><span data-ttu-id="dc38c-201">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-201">1.0</span></span> |
| [<span data-ttu-id="dc38c-202">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-202">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="dc38c-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-203">ReadItem</span></span> | <span data-ttu-id="dc38c-204">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-204">Compose</span></span><br><span data-ttu-id="dc38c-205">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-205">Read</span></span> | <span data-ttu-id="dc38c-206">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-206">1.0</span></span> |
| [<span data-ttu-id="dc38c-207">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-207">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="dc38c-208">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dc38c-208">ReadWriteMailbox</span></span> | <span data-ttu-id="dc38c-209">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-209">Compose</span></span><br><span data-ttu-id="dc38c-210">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-210">Read</span></span> | <span data-ttu-id="dc38c-211">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-211">1.0</span></span> |
| [<span data-ttu-id="dc38c-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dc38c-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="dc38c-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-213">ReadItem</span></span> | <span data-ttu-id="dc38c-214">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-214">Compose</span></span><br><span data-ttu-id="dc38c-215">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-215">Read</span></span> | <span data-ttu-id="dc38c-216">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-216">1.5</span></span> |

##### <a name="events"></a><span data-ttu-id="dc38c-217">Événements</span><span class="sxs-lookup"><span data-stu-id="dc38c-217">Events</span></span>

<span data-ttu-id="dc38c-218">Vous pouvez vous abonner et annuler l’abonnement aux événements suivants à l’aide de [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) et [removeHandlerAsync](#removehandlerasynceventtype-options-callback) , respectivement.</span><span class="sxs-lookup"><span data-stu-id="dc38c-218">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="dc38c-219">Événement</span><span class="sxs-lookup"><span data-stu-id="dc38c-219">Event</span></span> | <span data-ttu-id="dc38c-220">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-220">Description</span></span> | <span data-ttu-id="dc38c-221">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc38c-221">Minimum</span></span><br><span data-ttu-id="dc38c-222">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-222">requirement set</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="dc38c-223">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="dc38c-223">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dc38c-224">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-224">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="dc38c-225">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="dc38c-225">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="dc38c-226">Aperçu</span><span class="sxs-lookup"><span data-stu-id="dc38c-226">Preview</span></span> |

### <a name="namespaces"></a><span data-ttu-id="dc38c-227">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="dc38c-227">Namespaces</span></span>

<span data-ttu-id="dc38c-228">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc38c-228">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="dc38c-229">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc38c-229">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="dc38c-230">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc38c-230">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

## <a name="property-details"></a><span data-ttu-id="dc38c-231">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="dc38c-231">Property details</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="dc38c-232">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="dc38c-232">ewsUrl: String</span></span>

<span data-ttu-id="dc38c-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-235">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-235">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dc38c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dc38c-238">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="dc38c-238">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="dc38c-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc38c-241">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-241">Type</span></span>

*   <span data-ttu-id="dc38c-242">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-242">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc38c-243">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-243">Requirements</span></span>

|<span data-ttu-id="dc38c-244">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-244">Requirement</span></span>| <span data-ttu-id="dc38c-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-246">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-247">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-247">1.0</span></span>|
|[<span data-ttu-id="dc38c-248">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-248">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-249">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-250">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-251">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-251">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="dc38c-252">masterCategories : [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="dc38c-252">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="dc38c-253">Obtient un objet qui fournit des méthodes pour gérer la liste de formes de base des catégories sur cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="dc38c-253">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-254">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-254">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="dc38c-255">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-255">Type</span></span>

*   [<span data-ttu-id="dc38c-256">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="dc38c-256">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="dc38c-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-257">Requirements</span></span>

|<span data-ttu-id="dc38c-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-258">Requirement</span></span>| <span data-ttu-id="dc38c-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-261">1.8</span><span class="sxs-lookup"><span data-stu-id="dc38c-261">1.8</span></span> |
|[<span data-ttu-id="dc38c-262">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-262">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-263">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dc38c-263">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="dc38c-264">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-264">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-265">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-265">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="dc38c-266">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-266">Example</span></span>

<span data-ttu-id="dc38c-267">Cet exemple obtient la liste principale des catégories pour cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="dc38c-267">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="dc38c-268">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="dc38c-268">restUrl: String</span></span>

<span data-ttu-id="dc38c-269">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="dc38c-269">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="dc38c-270">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-270">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dc38c-271">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-271">Type</span></span>

*   <span data-ttu-id="dc38c-272">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc38c-273">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-273">Requirements</span></span>

|<span data-ttu-id="dc38c-274">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-274">Requirement</span></span>| <span data-ttu-id="dc38c-275">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-276">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-277">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-277">1.5</span></span> |
|[<span data-ttu-id="dc38c-278">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-279">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-280">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-281">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-281">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="dc38c-282">Détails de méthodes</span><span class="sxs-lookup"><span data-stu-id="dc38c-282">Method details</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="dc38c-283">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dc38c-283">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="dc38c-284">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="dc38c-284">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="dc38c-285">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="dc38c-285">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-286">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-286">Parameters</span></span>

| <span data-ttu-id="dc38c-287">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-287">Name</span></span> | <span data-ttu-id="dc38c-288">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-288">Type</span></span> | <span data-ttu-id="dc38c-289">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-289">Attributes</span></span> | <span data-ttu-id="dc38c-290">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-290">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dc38c-291">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dc38c-291">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dc38c-292">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="dc38c-292">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="dc38c-293">Fonction</span><span class="sxs-lookup"><span data-stu-id="dc38c-293">Function</span></span> || <span data-ttu-id="dc38c-p104">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="dc38c-297">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-297">Object</span></span> | <span data-ttu-id="dc38c-298">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-298">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-299">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-299">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dc38c-300">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-300">Object</span></span> | <span data-ttu-id="dc38c-301">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-301">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-302">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="dc38c-302">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dc38c-303">fonction</span><span class="sxs-lookup"><span data-stu-id="dc38c-303">function</span></span>| <span data-ttu-id="dc38c-304">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-304">&lt;optional&gt;</span></span>|<span data-ttu-id="dc38c-305">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-305">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-306">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-306">Requirements</span></span>

|<span data-ttu-id="dc38c-307">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-307">Requirement</span></span>| <span data-ttu-id="dc38c-308">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-309">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-310">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-310">1.5</span></span> |
|[<span data-ttu-id="dc38c-311">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-312">ReadItem</span></span> |
|[<span data-ttu-id="dc38c-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-314">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-314">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-315">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-315">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="dc38c-316">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dc38c-316">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dc38c-317">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="dc38c-317">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-318">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-318">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-p105">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-321">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dc38c-321">Parameters</span></span>

|<span data-ttu-id="dc38c-322">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-322">Name</span></span>| <span data-ttu-id="dc38c-323">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-323">Type</span></span>| <span data-ttu-id="dc38c-324">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-324">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dc38c-325">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-325">String</span></span>|<span data-ttu-id="dc38c-326">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="dc38c-326">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="dc38c-327">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dc38c-327">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="dc38c-328">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-328">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-329">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-329">Requirements</span></span>

|<span data-ttu-id="dc38c-330">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-330">Requirement</span></span>| <span data-ttu-id="dc38c-331">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-331">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-332">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-332">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-333">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-333">1.3</span></span>|
|[<span data-ttu-id="dc38c-334">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-334">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-335">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dc38c-335">Restricted</span></span>|
|[<span data-ttu-id="dc38c-336">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-336">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-337">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-337">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dc38c-338">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dc38c-338">Returns:</span></span>

<span data-ttu-id="dc38c-339">Type : String</span><span class="sxs-lookup"><span data-stu-id="dc38c-339">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dc38c-340">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-340">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="dc38c-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="dc38c-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="dc38c-342">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="dc38c-342">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="dc38c-p106">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="dc38c-p107">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-348">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dc38c-348">Parameters</span></span>

|<span data-ttu-id="dc38c-349">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-349">Name</span></span>| <span data-ttu-id="dc38c-350">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-350">Type</span></span>| <span data-ttu-id="dc38c-351">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-351">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="dc38c-352">Date</span><span class="sxs-lookup"><span data-stu-id="dc38c-352">Date</span></span>|<span data-ttu-id="dc38c-353">Objet Date</span><span class="sxs-lookup"><span data-stu-id="dc38c-353">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-354">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-354">Requirements</span></span>

|<span data-ttu-id="dc38c-355">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-355">Requirement</span></span>| <span data-ttu-id="dc38c-356">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-357">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-358">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-358">1.0</span></span>|
|[<span data-ttu-id="dc38c-359">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-360">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-361">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-362">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-362">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dc38c-363">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dc38c-363">Returns:</span></span>

<span data-ttu-id="dc38c-364">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="dc38c-364">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="dc38c-365">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dc38c-365">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dc38c-366">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="dc38c-366">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-367">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-p108">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-370">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-370">Parameters</span></span>

|<span data-ttu-id="dc38c-371">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-371">Name</span></span>| <span data-ttu-id="dc38c-372">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-372">Type</span></span>| <span data-ttu-id="dc38c-373">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-373">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dc38c-374">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-374">String</span></span>|<span data-ttu-id="dc38c-375">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="dc38c-375">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="dc38c-376">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dc38c-376">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="dc38c-377">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="dc38c-377">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-378">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-378">Requirements</span></span>

|<span data-ttu-id="dc38c-379">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-379">Requirement</span></span>| <span data-ttu-id="dc38c-380">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-381">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-382">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-382">1.3</span></span>|
|[<span data-ttu-id="dc38c-383">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-384">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dc38c-384">Restricted</span></span>|
|[<span data-ttu-id="dc38c-385">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-386">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-386">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dc38c-387">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dc38c-387">Returns:</span></span>

<span data-ttu-id="dc38c-388">Type : String</span><span class="sxs-lookup"><span data-stu-id="dc38c-388">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dc38c-389">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-389">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="dc38c-390">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="dc38c-390">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="dc38c-391">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="dc38c-391">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="dc38c-392">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="dc38c-392">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-393">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-393">Parameters</span></span>

|<span data-ttu-id="dc38c-394">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-394">Name</span></span>| <span data-ttu-id="dc38c-395">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-395">Type</span></span>| <span data-ttu-id="dc38c-396">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-396">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="dc38c-397">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dc38c-397">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="dc38c-398">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="dc38c-398">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-399">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-399">Requirements</span></span>

|<span data-ttu-id="dc38c-400">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-400">Requirement</span></span>| <span data-ttu-id="dc38c-401">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-402">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-403">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-403">1.0</span></span>|
|[<span data-ttu-id="dc38c-404">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-405">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-406">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-407">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-407">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dc38c-408">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dc38c-408">Returns:</span></span>

<span data-ttu-id="dc38c-409">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="dc38c-409">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="dc38c-410">Type : Date</span><span class="sxs-lookup"><span data-stu-id="dc38c-410">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="dc38c-411">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-411">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="dc38c-412">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dc38c-412">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="dc38c-413">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="dc38c-413">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-414">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-414">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-415">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="dc38c-415">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dc38c-p109">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="dc38c-418">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-418">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="dc38c-419">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="dc38c-419">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-420">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-420">Parameters</span></span>

|<span data-ttu-id="dc38c-421">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-421">Name</span></span>| <span data-ttu-id="dc38c-422">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-422">Type</span></span>| <span data-ttu-id="dc38c-423">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-423">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dc38c-424">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-424">String</span></span>|<span data-ttu-id="dc38c-425">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="dc38c-425">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-426">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-426">Requirements</span></span>

|<span data-ttu-id="dc38c-427">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-427">Requirement</span></span>| <span data-ttu-id="dc38c-428">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-429">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-430">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-430">1.0</span></span>|
|[<span data-ttu-id="dc38c-431">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-432">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-433">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-434">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-435">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-435">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="dc38c-436">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dc38c-436">displayMessageForm(itemId)</span></span>

<span data-ttu-id="dc38c-437">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="dc38c-437">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-438">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-438">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-439">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="dc38c-439">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dc38c-440">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-440">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="dc38c-441">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="dc38c-441">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="dc38c-p110">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-444">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-444">Parameters</span></span>

|<span data-ttu-id="dc38c-445">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-445">Name</span></span>| <span data-ttu-id="dc38c-446">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-446">Type</span></span>| <span data-ttu-id="dc38c-447">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-447">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dc38c-448">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dc38c-448">String</span></span>|<span data-ttu-id="dc38c-449">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="dc38c-449">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-450">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-450">Requirements</span></span>

|<span data-ttu-id="dc38c-451">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-451">Requirement</span></span>| <span data-ttu-id="dc38c-452">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-453">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-454">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-454">1.0</span></span>|
|[<span data-ttu-id="dc38c-455">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-456">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-457">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-458">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-458">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-459">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-459">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="dc38c-460">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="dc38c-460">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="dc38c-461">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="dc38c-461">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-462">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dc38c-462">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dc38c-p111">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dc38c-p112">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="dc38c-p113">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="dc38c-470">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="dc38c-470">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-471">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dc38c-471">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-472">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="dc38c-472">All parameters are optional.</span></span>

|<span data-ttu-id="dc38c-473">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-473">Name</span></span>| <span data-ttu-id="dc38c-474">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-474">Type</span></span>| <span data-ttu-id="dc38c-475">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-475">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dc38c-476">Object</span><span class="sxs-lookup"><span data-stu-id="dc38c-476">Object</span></span> | <span data-ttu-id="dc38c-477">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dc38c-477">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="dc38c-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dc38c-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="dc38c-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dc38c-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="dc38c-484">Date</span><span class="sxs-lookup"><span data-stu-id="dc38c-484">Date</span></span> | <span data-ttu-id="dc38c-485">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dc38c-485">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="dc38c-486">Date</span><span class="sxs-lookup"><span data-stu-id="dc38c-486">Date</span></span> | <span data-ttu-id="dc38c-487">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dc38c-487">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="dc38c-488">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dc38c-488">String</span></span> | <span data-ttu-id="dc38c-p116">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="dc38c-491">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-491">Array.&lt;String&gt;</span></span> | <span data-ttu-id="dc38c-p117">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dc38c-494">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-494">String</span></span> | <span data-ttu-id="dc38c-p118">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="dc38c-497">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-497">String</span></span> | <span data-ttu-id="dc38c-p119">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc38c-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-500">Requirements</span></span>

|<span data-ttu-id="dc38c-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-501">Requirement</span></span>| <span data-ttu-id="dc38c-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-504">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-504">1.0</span></span>|
|[<span data-ttu-id="dc38c-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-506">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-508">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-509">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-509">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="dc38c-510">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="dc38c-510">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="dc38c-511">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-511">Displays a form for creating a new message.</span></span>

<span data-ttu-id="dc38c-512">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-512">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="dc38c-513">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="dc38c-513">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dc38c-514">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="dc38c-514">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-515">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dc38c-515">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-516">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="dc38c-516">All parameters are optional.</span></span>

|<span data-ttu-id="dc38c-517">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-517">Name</span></span>| <span data-ttu-id="dc38c-518">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-518">Type</span></span>| <span data-ttu-id="dc38c-519">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-519">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dc38c-520">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-520">Object</span></span> | <span data-ttu-id="dc38c-521">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-521">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="dc38c-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dc38c-523">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="dc38c-523">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="dc38c-524">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-524">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="dc38c-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dc38c-526">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="dc38c-526">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="dc38c-527">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-527">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="dc38c-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dc38c-529">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="dc38c-529">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="dc38c-530">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-530">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dc38c-531">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-531">String</span></span> | <span data-ttu-id="dc38c-532">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-532">A string containing the subject of the message.</span></span> <span data-ttu-id="dc38c-533">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dc38c-533">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="dc38c-534">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dc38c-534">String</span></span> | <span data-ttu-id="dc38c-535">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-535">The HTML body of the message.</span></span> <span data-ttu-id="dc38c-536">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="dc38c-536">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="dc38c-537">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-537">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dc38c-538">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-538">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="dc38c-539">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-539">String</span></span> | <span data-ttu-id="dc38c-p126">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="dc38c-542">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-542">String</span></span> | <span data-ttu-id="dc38c-543">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="dc38c-543">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="dc38c-544">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dc38c-544">String</span></span> | <span data-ttu-id="dc38c-p127">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="dc38c-547">Booléen</span><span class="sxs-lookup"><span data-stu-id="dc38c-547">Boolean</span></span> | <span data-ttu-id="dc38c-p128">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="dc38c-550">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-550">String</span></span> | <span data-ttu-id="dc38c-551">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-551">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="dc38c-552">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="dc38c-552">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="dc38c-553">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="dc38c-553">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="dc38c-554">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-554">Requirements</span></span>

|<span data-ttu-id="dc38c-555">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-555">Requirement</span></span>| <span data-ttu-id="dc38c-556">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-557">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-558">1.6</span><span class="sxs-lookup"><span data-stu-id="dc38c-558">1.6</span></span> |
|[<span data-ttu-id="dc38c-559">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-560">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-561">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-562">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-562">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-563">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-563">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="dc38c-564">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="dc38c-564">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="dc38c-565">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="dc38c-565">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="dc38c-p130">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-568">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="dc38c-568">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="dc38c-569">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-569">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dc38c-570">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-570">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dc38c-571">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-571">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="dc38c-572">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="dc38c-572">**REST Tokens**</span></span>

<span data-ttu-id="dc38c-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="dc38c-576">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="dc38c-576">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="dc38c-577">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="dc38c-577">**EWS Tokens**</span></span>

<span data-ttu-id="dc38c-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="dc38c-580">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="dc38c-580">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="dc38c-581">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="dc38c-581">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dc38c-582">Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) des services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour récupérer une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-582">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="dc38c-583">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dc38c-583">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-584">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-584">Parameters</span></span>

|<span data-ttu-id="dc38c-585">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-585">Name</span></span>| <span data-ttu-id="dc38c-586">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-586">Type</span></span>| <span data-ttu-id="dc38c-587">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-587">Attributes</span></span>| <span data-ttu-id="dc38c-588">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-588">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="dc38c-589">Object</span><span class="sxs-lookup"><span data-stu-id="dc38c-589">Object</span></span> | <span data-ttu-id="dc38c-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-590">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-591">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="dc38c-592">Boolean</span><span class="sxs-lookup"><span data-stu-id="dc38c-592">Boolean</span></span> |  <span data-ttu-id="dc38c-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-593">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-p135">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dc38c-596">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-596">Object</span></span> |  <span data-ttu-id="dc38c-597">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-597">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-598">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dc38c-598">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="dc38c-599">fonction</span><span class="sxs-lookup"><span data-stu-id="dc38c-599">function</span></span>||<span data-ttu-id="dc38c-600">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-600">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dc38c-601">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-601">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dc38c-602">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dc38c-602">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dc38c-603">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dc38c-603">Errors</span></span>

|<span data-ttu-id="dc38c-604">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dc38c-604">Error code</span></span>|<span data-ttu-id="dc38c-605">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-605">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dc38c-606">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dc38c-606">The request has failed.</span></span> <span data-ttu-id="dc38c-607">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-607">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dc38c-608">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-608">The Exchange server returned an error.</span></span> <span data-ttu-id="dc38c-609">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-609">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dc38c-610">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dc38c-610">The user is no longer connected to the network.</span></span> <span data-ttu-id="dc38c-611">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dc38c-611">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-612">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-612">Requirements</span></span>

|<span data-ttu-id="dc38c-613">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-613">Requirement</span></span>| <span data-ttu-id="dc38c-614">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-615">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-616">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-616">1.5</span></span> |
|[<span data-ttu-id="dc38c-617">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-618">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-619">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-620">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-620">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-621">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-621">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="dc38c-622">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dc38c-622">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dc38c-623">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="dc38c-623">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="dc38c-p139">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="dc38c-626">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="dc38c-626">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dc38c-627">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-627">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="dc38c-628">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dc38c-628">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dc38c-629">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-629">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dc38c-630">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-630">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dc38c-631">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dc38c-631">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-632">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-632">Parameters</span></span>

|<span data-ttu-id="dc38c-633">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-633">Name</span></span>| <span data-ttu-id="dc38c-634">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-634">Type</span></span>| <span data-ttu-id="dc38c-635">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-635">Attributes</span></span>| <span data-ttu-id="dc38c-636">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-636">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dc38c-637">function</span><span class="sxs-lookup"><span data-stu-id="dc38c-637">function</span></span>||<span data-ttu-id="dc38c-638">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dc38c-639">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-639">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dc38c-640">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dc38c-640">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dc38c-641">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-641">Object</span></span>| <span data-ttu-id="dc38c-642">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-642">&lt;optional&gt;</span></span>|<span data-ttu-id="dc38c-643">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dc38c-643">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dc38c-644">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dc38c-644">Errors</span></span>

|<span data-ttu-id="dc38c-645">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dc38c-645">Error code</span></span>|<span data-ttu-id="dc38c-646">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-646">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dc38c-647">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dc38c-647">The request has failed.</span></span> <span data-ttu-id="dc38c-648">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-648">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dc38c-649">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-649">The Exchange server returned an error.</span></span> <span data-ttu-id="dc38c-650">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-650">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dc38c-651">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dc38c-651">The user is no longer connected to the network.</span></span> <span data-ttu-id="dc38c-652">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dc38c-652">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-653">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-653">Requirements</span></span>

|<span data-ttu-id="dc38c-654">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-654">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="dc38c-655">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-656">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-656">1.0</span></span> | <span data-ttu-id="dc38c-657">1.3</span><span class="sxs-lookup"><span data-stu-id="dc38c-657">1.3</span></span> |
|[<span data-ttu-id="dc38c-658">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-658">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-659">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-659">ReadItem</span></span> | <span data-ttu-id="dc38c-660">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-660">ReadItem</span></span> |
|[<span data-ttu-id="dc38c-661">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-661">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-662">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-662">Read</span></span> | <span data-ttu-id="dc38c-663">Composition</span><span class="sxs-lookup"><span data-stu-id="dc38c-663">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="dc38c-664">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-664">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="dc38c-665">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dc38c-665">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dc38c-666">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dc38c-666">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="dc38c-667">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="dc38c-667">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-668">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dc38c-668">Parameters</span></span>

|<span data-ttu-id="dc38c-669">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-669">Name</span></span>| <span data-ttu-id="dc38c-670">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-670">Type</span></span>| <span data-ttu-id="dc38c-671">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-671">Attributes</span></span>| <span data-ttu-id="dc38c-672">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-672">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dc38c-673">fonction</span><span class="sxs-lookup"><span data-stu-id="dc38c-673">function</span></span>||<span data-ttu-id="dc38c-674">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-674">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dc38c-675">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-675">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dc38c-676">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dc38c-676">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dc38c-677">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-677">Object</span></span>| <span data-ttu-id="dc38c-678">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-678">&lt;optional&gt;</span></span>|<span data-ttu-id="dc38c-679">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dc38c-679">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dc38c-680">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dc38c-680">Errors</span></span>

|<span data-ttu-id="dc38c-681">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dc38c-681">Error code</span></span>|<span data-ttu-id="dc38c-682">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-682">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dc38c-683">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dc38c-683">The request has failed.</span></span> <span data-ttu-id="dc38c-684">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-684">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dc38c-685">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-685">The Exchange server returned an error.</span></span> <span data-ttu-id="dc38c-686">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dc38c-686">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dc38c-687">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dc38c-687">The user is no longer connected to the network.</span></span> <span data-ttu-id="dc38c-688">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dc38c-688">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-689">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-689">Requirements</span></span>

|<span data-ttu-id="dc38c-690">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-690">Requirement</span></span>| <span data-ttu-id="dc38c-691">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-691">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-692">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-692">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-693">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-693">1.0</span></span>|
|[<span data-ttu-id="dc38c-694">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-694">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-695">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-695">ReadItem</span></span>|
|[<span data-ttu-id="dc38c-696">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-696">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-697">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-697">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-698">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-698">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="dc38c-699">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dc38c-699">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="dc38c-700">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-700">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-701">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="dc38c-701">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="dc38c-702">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="dc38c-702">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="dc38c-703">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="dc38c-703">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="dc38c-704">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dc38c-704">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="dc38c-705">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-705">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="dc38c-706">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="dc38c-706">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="dc38c-707">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-707">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="dc38c-708">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="dc38c-708">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="dc38c-p149">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="dc38c-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="dc38c-711">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="dc38c-711">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="dc38c-712">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="dc38c-712">Version differences</span></span>

<span data-ttu-id="dc38c-713">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-713">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="dc38c-714">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage.</span><span class="sxs-lookup"><span data-stu-id="dc38c-714">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="dc38c-715">Vous pouvez déterminer si votre application de messagerie est en cours d’exécution dans Outlook sur le Web ou sur un client de bureau à l’aide de la propriété Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="dc38c-715">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="dc38c-716">Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="dc38c-716">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-717">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-717">Parameters</span></span>

|<span data-ttu-id="dc38c-718">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-718">Name</span></span>| <span data-ttu-id="dc38c-719">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-719">Type</span></span>| <span data-ttu-id="dc38c-720">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-720">Attributes</span></span>| <span data-ttu-id="dc38c-721">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-721">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dc38c-722">String</span><span class="sxs-lookup"><span data-stu-id="dc38c-722">String</span></span>||<span data-ttu-id="dc38c-723">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="dc38c-723">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="dc38c-724">function</span><span class="sxs-lookup"><span data-stu-id="dc38c-724">function</span></span>||<span data-ttu-id="dc38c-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dc38c-726">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dc38c-726">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="dc38c-727">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="dc38c-727">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="dc38c-728">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-728">Object</span></span>| <span data-ttu-id="dc38c-729">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-729">&lt;optional&gt;</span></span>|<span data-ttu-id="dc38c-730">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dc38c-730">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-731">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-731">Requirements</span></span>

|<span data-ttu-id="dc38c-732">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-732">Requirement</span></span>| <span data-ttu-id="dc38c-733">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-734">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-735">1.0</span><span class="sxs-lookup"><span data-stu-id="dc38c-735">1.0</span></span>|
|[<span data-ttu-id="dc38c-736">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-737">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dc38c-737">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="dc38c-738">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-739">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-739">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc38c-740">Exemple</span><span class="sxs-lookup"><span data-stu-id="dc38c-740">Example</span></span>

<span data-ttu-id="dc38c-741">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="dc38c-741">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="dc38c-742">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dc38c-742">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="dc38c-743">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="dc38c-743">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="dc38c-744">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="dc38c-744">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dc38c-745">Parameters</span><span class="sxs-lookup"><span data-stu-id="dc38c-745">Parameters</span></span>

| <span data-ttu-id="dc38c-746">Nom</span><span class="sxs-lookup"><span data-stu-id="dc38c-746">Name</span></span> | <span data-ttu-id="dc38c-747">Type</span><span class="sxs-lookup"><span data-stu-id="dc38c-747">Type</span></span> | <span data-ttu-id="dc38c-748">Attributs</span><span class="sxs-lookup"><span data-stu-id="dc38c-748">Attributes</span></span> | <span data-ttu-id="dc38c-749">Description</span><span class="sxs-lookup"><span data-stu-id="dc38c-749">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dc38c-750">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dc38c-750">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dc38c-751">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="dc38c-751">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="dc38c-752">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-752">Object</span></span> | <span data-ttu-id="dc38c-753">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-753">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-754">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dc38c-754">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dc38c-755">Objet</span><span class="sxs-lookup"><span data-stu-id="dc38c-755">Object</span></span> | <span data-ttu-id="dc38c-756">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-756">&lt;optional&gt;</span></span> | <span data-ttu-id="dc38c-757">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="dc38c-757">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dc38c-758">fonction</span><span class="sxs-lookup"><span data-stu-id="dc38c-758">function</span></span>| <span data-ttu-id="dc38c-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dc38c-759">&lt;optional&gt;</span></span>|<span data-ttu-id="dc38c-760">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dc38c-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc38c-761">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc38c-761">Requirements</span></span>

|<span data-ttu-id="dc38c-762">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc38c-762">Requirement</span></span>| <span data-ttu-id="dc38c-763">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc38c-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc38c-764">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc38c-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc38c-765">1,5</span><span class="sxs-lookup"><span data-stu-id="dc38c-765">1.5</span></span> |
|[<span data-ttu-id="dc38c-766">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dc38c-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc38c-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc38c-767">ReadItem</span></span> |
|[<span data-ttu-id="dc38c-768">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc38c-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc38c-769">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc38c-769">Compose or Read</span></span>|
