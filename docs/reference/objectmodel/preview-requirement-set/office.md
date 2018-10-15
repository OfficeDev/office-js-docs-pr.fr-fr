 

# <a name="office"></a><span data-ttu-id="f6f07-101">Office</span><span class="sxs-lookup"><span data-stu-id="f6f07-101">Office</span></span>

<span data-ttu-id="f6f07-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f6f07-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6f07-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-104">Requirements</span></span>

|<span data-ttu-id="f6f07-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f6f07-105">Requirement</span></span>| <span data-ttu-id="f6f07-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="f6f07-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6f07-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f6f07-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6f07-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f6f07-108">1.0</span></span>|
|[<span data-ttu-id="f6f07-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f6f07-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6f07-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f6f07-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f6f07-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="f6f07-111">Members and methods</span></span>

| <span data-ttu-id="f6f07-112">Membre</span><span class="sxs-lookup"><span data-stu-id="f6f07-112">Member</span></span> | <span data-ttu-id="f6f07-113">Type</span><span class="sxs-lookup"><span data-stu-id="f6f07-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f6f07-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f6f07-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f6f07-115">Membre</span><span class="sxs-lookup"><span data-stu-id="f6f07-115">Member</span></span> |
| [<span data-ttu-id="f6f07-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f6f07-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f6f07-117">Membre</span><span class="sxs-lookup"><span data-stu-id="f6f07-117">Member</span></span> |
| [<span data-ttu-id="f6f07-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f6f07-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f6f07-119">Membre</span><span class="sxs-lookup"><span data-stu-id="f6f07-119">Member</span></span> |
| [<span data-ttu-id="f6f07-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f6f07-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f6f07-121">Membre</span><span class="sxs-lookup"><span data-stu-id="f6f07-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f6f07-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f6f07-122">Namespaces</span></span>

<span data-ttu-id="f6f07-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6f07-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f6f07-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f6f07-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f6f07-125">Membres</span><span class="sxs-lookup"><span data-stu-id="f6f07-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f6f07-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f6f07-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f6f07-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f6f07-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f6f07-128">Type :</span><span class="sxs-lookup"><span data-stu-id="f6f07-128">Type:</span></span>

*   <span data-ttu-id="f6f07-129">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f6f07-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f6f07-130">Properties:</span></span>

|<span data-ttu-id="f6f07-131">Nom</span><span class="sxs-lookup"><span data-stu-id="f6f07-131">Name</span></span>| <span data-ttu-id="f6f07-132">Type</span><span class="sxs-lookup"><span data-stu-id="f6f07-132">Type</span></span>| <span data-ttu-id="f6f07-133">Description</span><span class="sxs-lookup"><span data-stu-id="f6f07-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f6f07-134">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-134">String</span></span>|<span data-ttu-id="f6f07-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f6f07-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f6f07-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-136">String</span></span>|<span data-ttu-id="f6f07-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f6f07-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6f07-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-138">Requirements</span></span>

|<span data-ttu-id="f6f07-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f6f07-139">Requirement</span></span>| <span data-ttu-id="f6f07-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="f6f07-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6f07-141">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f6f07-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6f07-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f6f07-142">1.0</span></span>|
|[<span data-ttu-id="f6f07-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f6f07-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6f07-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f6f07-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f6f07-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f6f07-145">CoercionType :String</span></span>

<span data-ttu-id="f6f07-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f6f07-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f6f07-147">Type :</span><span class="sxs-lookup"><span data-stu-id="f6f07-147">Type:</span></span>

*   <span data-ttu-id="f6f07-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f6f07-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f6f07-149">Properties:</span></span>

|<span data-ttu-id="f6f07-150">Nom</span><span class="sxs-lookup"><span data-stu-id="f6f07-150">Name</span></span>| <span data-ttu-id="f6f07-151">Type</span><span class="sxs-lookup"><span data-stu-id="f6f07-151">Type</span></span>| <span data-ttu-id="f6f07-152">Description</span><span class="sxs-lookup"><span data-stu-id="f6f07-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f6f07-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-153">String</span></span>|<span data-ttu-id="f6f07-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f6f07-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f6f07-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-155">String</span></span>|<span data-ttu-id="f6f07-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f6f07-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6f07-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-157">Requirements</span></span>

|<span data-ttu-id="f6f07-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f6f07-158">Requirement</span></span>| <span data-ttu-id="f6f07-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="f6f07-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6f07-160">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f6f07-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6f07-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f6f07-161">1.0</span></span>|
|[<span data-ttu-id="f6f07-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f6f07-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6f07-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f6f07-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f6f07-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f6f07-164">EventType :String</span></span>

<span data-ttu-id="f6f07-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="f6f07-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f6f07-166">Type :</span><span class="sxs-lookup"><span data-stu-id="f6f07-166">Type:</span></span>

*   <span data-ttu-id="f6f07-167">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f6f07-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f6f07-168">Properties:</span></span>

| <span data-ttu-id="f6f07-169">Nom</span><span class="sxs-lookup"><span data-stu-id="f6f07-169">Name</span></span> | <span data-ttu-id="f6f07-170">Type</span><span class="sxs-lookup"><span data-stu-id="f6f07-170">Type</span></span> | <span data-ttu-id="f6f07-171">Description</span><span class="sxs-lookup"><span data-stu-id="f6f07-171">Description</span></span> | <span data-ttu-id="f6f07-172">Ensemble minimal de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f6f07-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-173">String</span></span> | <span data-ttu-id="f6f07-174">La date ou l’heure du rendez-vous sélectionné ou de la série a changé.</span><span class="sxs-lookup"><span data-stu-id="f6f07-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f6f07-175">1.7</span><span class="sxs-lookup"><span data-stu-id="f6f07-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="f6f07-176">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-176">String</span></span> | <span data-ttu-id="f6f07-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="f6f07-177">The selected item has changed.</span></span> | <span data-ttu-id="f6f07-178">1.5</span><span class="sxs-lookup"><span data-stu-id="f6f07-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f6f07-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-179">String</span></span> | <span data-ttu-id="f6f07-180">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="f6f07-180">The selected item has changed.</span></span> | <span data-ttu-id="f6f07-181">Aperçu</span><span class="sxs-lookup"><span data-stu-id="f6f07-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f6f07-182">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-182">String</span></span> | <span data-ttu-id="f6f07-183">La liste de destinataires de l’élément sélectionné ou l’emplacement du rendez-vous a changé.</span><span class="sxs-lookup"><span data-stu-id="f6f07-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f6f07-184">1.7</span><span class="sxs-lookup"><span data-stu-id="f6f07-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f6f07-185">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-185">String</span></span> | <span data-ttu-id="f6f07-186">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="f6f07-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f6f07-187">1.7</span><span class="sxs-lookup"><span data-stu-id="f6f07-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f6f07-188">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-188">Requirements</span></span>

|<span data-ttu-id="f6f07-189">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f6f07-189">Requirement</span></span>| <span data-ttu-id="f6f07-190">Valeur</span><span class="sxs-lookup"><span data-stu-id="f6f07-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6f07-191">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f6f07-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6f07-192">1.5</span><span class="sxs-lookup"><span data-stu-id="f6f07-192">1.5</span></span> |
|[<span data-ttu-id="f6f07-193">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f6f07-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6f07-194">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f6f07-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f6f07-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f6f07-195">SourceProperty :String</span></span>

<span data-ttu-id="f6f07-196">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f6f07-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f6f07-197">Type :</span><span class="sxs-lookup"><span data-stu-id="f6f07-197">Type:</span></span>

*   <span data-ttu-id="f6f07-198">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f6f07-199">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f6f07-199">Properties:</span></span>

|<span data-ttu-id="f6f07-200">Nom</span><span class="sxs-lookup"><span data-stu-id="f6f07-200">Name</span></span>| <span data-ttu-id="f6f07-201">Type</span><span class="sxs-lookup"><span data-stu-id="f6f07-201">Type</span></span>| <span data-ttu-id="f6f07-202">Description</span><span class="sxs-lookup"><span data-stu-id="f6f07-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f6f07-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-203">String</span></span>|<span data-ttu-id="f6f07-204">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6f07-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f6f07-205">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f6f07-205">String</span></span>|<span data-ttu-id="f6f07-206">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6f07-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6f07-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f6f07-207">Requirements</span></span>

|<span data-ttu-id="f6f07-208">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f6f07-208">Requirement</span></span>| <span data-ttu-id="f6f07-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="f6f07-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6f07-210">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f6f07-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6f07-211">1.0</span><span class="sxs-lookup"><span data-stu-id="f6f07-211">1.0</span></span>|
|[<span data-ttu-id="f6f07-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f6f07-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6f07-213">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f6f07-213">Compose or read</span></span>|