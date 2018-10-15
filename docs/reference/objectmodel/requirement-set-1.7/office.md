 

# <a name="office"></a><span data-ttu-id="4bead-101">Office</span><span class="sxs-lookup"><span data-stu-id="4bead-101">Office</span></span>

<span data-ttu-id="4bead-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4bead-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bead-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-104">Requirements</span></span>

|<span data-ttu-id="4bead-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4bead-105">Requirement</span></span>| <span data-ttu-id="4bead-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bead-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bead-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bead-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bead-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4bead-108">1.0</span></span>|
|[<span data-ttu-id="4bead-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bead-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bead-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bead-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4bead-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4bead-111">Members and methods</span></span>

| <span data-ttu-id="4bead-112">Membre</span><span class="sxs-lookup"><span data-stu-id="4bead-112">Member</span></span> | <span data-ttu-id="4bead-113">Type</span><span class="sxs-lookup"><span data-stu-id="4bead-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4bead-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4bead-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4bead-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4bead-115">Member</span></span> |
| [<span data-ttu-id="4bead-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bead-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4bead-117">Membre</span><span class="sxs-lookup"><span data-stu-id="4bead-117">Member</span></span> |
| [<span data-ttu-id="4bead-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4bead-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4bead-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4bead-119">Member</span></span> |
| [<span data-ttu-id="4bead-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4bead-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4bead-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4bead-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4bead-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4bead-122">Namespaces</span></span>

<span data-ttu-id="4bead-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bead-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4bead-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4bead-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4bead-125">Membres</span><span class="sxs-lookup"><span data-stu-id="4bead-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4bead-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4bead-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4bead-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bead-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4bead-128">Type :</span><span class="sxs-lookup"><span data-stu-id="4bead-128">Type:</span></span>

*   <span data-ttu-id="4bead-129">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bead-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4bead-130">Properties:</span></span>

|<span data-ttu-id="4bead-131">Nom</span><span class="sxs-lookup"><span data-stu-id="4bead-131">Name</span></span>| <span data-ttu-id="4bead-132">Type</span><span class="sxs-lookup"><span data-stu-id="4bead-132">Type</span></span>| <span data-ttu-id="4bead-133">Description</span><span class="sxs-lookup"><span data-stu-id="4bead-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4bead-134">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-134">String</span></span>|<span data-ttu-id="4bead-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4bead-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4bead-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-136">String</span></span>|<span data-ttu-id="4bead-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4bead-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bead-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-138">Requirements</span></span>

|<span data-ttu-id="4bead-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4bead-139">Requirement</span></span>| <span data-ttu-id="4bead-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bead-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bead-141">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bead-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bead-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4bead-142">1.0</span></span>|
|[<span data-ttu-id="4bead-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bead-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bead-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bead-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4bead-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4bead-145">CoercionType :String</span></span>

<span data-ttu-id="4bead-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4bead-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4bead-147">Type :</span><span class="sxs-lookup"><span data-stu-id="4bead-147">Type:</span></span>

*   <span data-ttu-id="4bead-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bead-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4bead-149">Properties:</span></span>

|<span data-ttu-id="4bead-150">Nom</span><span class="sxs-lookup"><span data-stu-id="4bead-150">Name</span></span>| <span data-ttu-id="4bead-151">Type</span><span class="sxs-lookup"><span data-stu-id="4bead-151">Type</span></span>| <span data-ttu-id="4bead-152">Description</span><span class="sxs-lookup"><span data-stu-id="4bead-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4bead-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-153">String</span></span>|<span data-ttu-id="4bead-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4bead-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4bead-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-155">String</span></span>|<span data-ttu-id="4bead-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4bead-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bead-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-157">Requirements</span></span>

|<span data-ttu-id="4bead-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4bead-158">Requirement</span></span>| <span data-ttu-id="4bead-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bead-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bead-160">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bead-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bead-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4bead-161">1.0</span></span>|
|[<span data-ttu-id="4bead-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bead-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bead-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bead-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4bead-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4bead-164">EventType :String</span></span>

<span data-ttu-id="4bead-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4bead-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4bead-166">Type :</span><span class="sxs-lookup"><span data-stu-id="4bead-166">Type:</span></span>

*   <span data-ttu-id="4bead-167">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bead-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4bead-168">Properties:</span></span>

| <span data-ttu-id="4bead-169">Nom</span><span class="sxs-lookup"><span data-stu-id="4bead-169">Name</span></span> | <span data-ttu-id="4bead-170">Type</span><span class="sxs-lookup"><span data-stu-id="4bead-170">Type</span></span> | <span data-ttu-id="4bead-171">Description</span><span class="sxs-lookup"><span data-stu-id="4bead-171">Description</span></span> | <span data-ttu-id="4bead-172">Ensemble minimal de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="4bead-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-173">String</span></span> | <span data-ttu-id="4bead-174">La date ou l’heure du rendez-vous sélectionné ou de la série a changé.</span><span class="sxs-lookup"><span data-stu-id="4bead-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4bead-175">1.7</span><span class="sxs-lookup"><span data-stu-id="4bead-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="4bead-176">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-176">String</span></span> | <span data-ttu-id="4bead-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4bead-177">The selected item has changed.</span></span> | <span data-ttu-id="4bead-178">1.5</span><span class="sxs-lookup"><span data-stu-id="4bead-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4bead-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-179">String</span></span> | <span data-ttu-id="4bead-180">La liste de destinataires de l’élément sélectionné ou l’emplacement du rendez-vous a changé.</span><span class="sxs-lookup"><span data-stu-id="4bead-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4bead-181">1.7</span><span class="sxs-lookup"><span data-stu-id="4bead-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4bead-182">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-182">String</span></span> | <span data-ttu-id="4bead-183">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="4bead-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4bead-184">1.7</span><span class="sxs-lookup"><span data-stu-id="4bead-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bead-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-185">Requirements</span></span>

|<span data-ttu-id="4bead-186">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4bead-186">Requirement</span></span>| <span data-ttu-id="4bead-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bead-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bead-188">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bead-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bead-189">1,5</span><span class="sxs-lookup"><span data-stu-id="4bead-189">1.5</span></span> |
|[<span data-ttu-id="4bead-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bead-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bead-191">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bead-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4bead-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4bead-192">SourceProperty :String</span></span>

<span data-ttu-id="4bead-193">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4bead-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4bead-194">Type :</span><span class="sxs-lookup"><span data-stu-id="4bead-194">Type:</span></span>

*   <span data-ttu-id="4bead-195">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bead-196">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4bead-196">Properties:</span></span>

|<span data-ttu-id="4bead-197">Nom</span><span class="sxs-lookup"><span data-stu-id="4bead-197">Name</span></span>| <span data-ttu-id="4bead-198">Type</span><span class="sxs-lookup"><span data-stu-id="4bead-198">Type</span></span>| <span data-ttu-id="4bead-199">Description</span><span class="sxs-lookup"><span data-stu-id="4bead-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4bead-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-200">String</span></span>|<span data-ttu-id="4bead-201">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bead-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4bead-202">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4bead-202">String</span></span>|<span data-ttu-id="4bead-203">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bead-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bead-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4bead-204">Requirements</span></span>

|<span data-ttu-id="4bead-205">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4bead-205">Requirement</span></span>| <span data-ttu-id="4bead-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="4bead-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bead-207">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4bead-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bead-208">1.0</span><span class="sxs-lookup"><span data-stu-id="4bead-208">1.0</span></span>|
|[<span data-ttu-id="4bead-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4bead-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4bead-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4bead-210">Compose or read</span></span>|