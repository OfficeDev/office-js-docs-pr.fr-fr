 

# <a name="office"></a><span data-ttu-id="002ec-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="002ec-101">Office</span></span>

<span data-ttu-id="002ec-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="002ec-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="002ec-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="002ec-104">Requirements</span></span>

|<span data-ttu-id="002ec-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="002ec-105">Requirement</span></span>| <span data-ttu-id="002ec-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="002ec-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="002ec-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="002ec-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="002ec-108">1.0</span><span class="sxs-lookup"><span data-stu-id="002ec-108">1.0</span></span>|
|[<span data-ttu-id="002ec-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="002ec-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="002ec-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="002ec-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="002ec-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="002ec-111">Members and methods</span></span>

| <span data-ttu-id="002ec-112">Membre</span><span class="sxs-lookup"><span data-stu-id="002ec-112">Member</span></span> | <span data-ttu-id="002ec-113">Type</span><span class="sxs-lookup"><span data-stu-id="002ec-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="002ec-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="002ec-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="002ec-115">Membre</span><span class="sxs-lookup"><span data-stu-id="002ec-115">Member</span></span> |
| [<span data-ttu-id="002ec-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="002ec-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="002ec-117">Membre</span><span class="sxs-lookup"><span data-stu-id="002ec-117">Member</span></span> |
| [<span data-ttu-id="002ec-118">EventType</span><span class="sxs-lookup"><span data-stu-id="002ec-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="002ec-119">Membre</span><span class="sxs-lookup"><span data-stu-id="002ec-119">Member</span></span> |
| [<span data-ttu-id="002ec-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="002ec-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="002ec-121">Membre</span><span class="sxs-lookup"><span data-stu-id="002ec-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="002ec-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="002ec-122">Namespaces</span></span>

<span data-ttu-id="002ec-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="002ec-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="002ec-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="002ec-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="002ec-125">Membres</span><span class="sxs-lookup"><span data-stu-id="002ec-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="002ec-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="002ec-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="002ec-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="002ec-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="002ec-128">Type :</span><span class="sxs-lookup"><span data-stu-id="002ec-128">Type:</span></span>

*   <span data-ttu-id="002ec-129">String</span><span class="sxs-lookup"><span data-stu-id="002ec-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="002ec-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="002ec-130">Properties:</span></span>

|<span data-ttu-id="002ec-131">Nom</span><span class="sxs-lookup"><span data-stu-id="002ec-131">Name</span></span>| <span data-ttu-id="002ec-132">Type</span><span class="sxs-lookup"><span data-stu-id="002ec-132">Type</span></span>| <span data-ttu-id="002ec-133">Description</span><span class="sxs-lookup"><span data-stu-id="002ec-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="002ec-134">String</span><span class="sxs-lookup"><span data-stu-id="002ec-134">String</span></span>|<span data-ttu-id="002ec-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="002ec-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="002ec-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-136">String</span></span>|<span data-ttu-id="002ec-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="002ec-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="002ec-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="002ec-138">Requirements</span></span>

|<span data-ttu-id="002ec-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="002ec-139">Requirement</span></span>| <span data-ttu-id="002ec-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="002ec-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="002ec-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="002ec-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="002ec-142">1.0</span><span class="sxs-lookup"><span data-stu-id="002ec-142">1.0</span></span>|
|[<span data-ttu-id="002ec-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="002ec-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="002ec-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="002ec-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="002ec-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="002ec-145">CoercionType :String</span></span>

<span data-ttu-id="002ec-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="002ec-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="002ec-147">Type :</span><span class="sxs-lookup"><span data-stu-id="002ec-147">Type:</span></span>

*   <span data-ttu-id="002ec-148">String</span><span class="sxs-lookup"><span data-stu-id="002ec-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="002ec-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="002ec-149">Properties:</span></span>

|<span data-ttu-id="002ec-150">Nom</span><span class="sxs-lookup"><span data-stu-id="002ec-150">Name</span></span>| <span data-ttu-id="002ec-151">Type</span><span class="sxs-lookup"><span data-stu-id="002ec-151">Type</span></span>| <span data-ttu-id="002ec-152">Description</span><span class="sxs-lookup"><span data-stu-id="002ec-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="002ec-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-153">String</span></span>|<span data-ttu-id="002ec-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="002ec-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="002ec-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-155">String</span></span>|<span data-ttu-id="002ec-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="002ec-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="002ec-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="002ec-157">Requirements</span></span>

|<span data-ttu-id="002ec-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="002ec-158">Requirement</span></span>| <span data-ttu-id="002ec-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="002ec-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="002ec-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="002ec-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="002ec-161">1.0</span><span class="sxs-lookup"><span data-stu-id="002ec-161">1.0</span></span>|
|[<span data-ttu-id="002ec-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="002ec-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="002ec-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="002ec-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="002ec-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="002ec-164">EventType :String</span></span>

<span data-ttu-id="002ec-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="002ec-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="002ec-166">Type :</span><span class="sxs-lookup"><span data-stu-id="002ec-166">Type:</span></span>

*   <span data-ttu-id="002ec-167">String</span><span class="sxs-lookup"><span data-stu-id="002ec-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="002ec-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="002ec-168">Properties:</span></span>

| <span data-ttu-id="002ec-169">Nom</span><span class="sxs-lookup"><span data-stu-id="002ec-169">Name</span></span> | <span data-ttu-id="002ec-170">Type</span><span class="sxs-lookup"><span data-stu-id="002ec-170">Type</span></span> | <span data-ttu-id="002ec-171">Description</span><span class="sxs-lookup"><span data-stu-id="002ec-171">Description</span></span> | <span data-ttu-id="002ec-172">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="002ec-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="002ec-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-173">String</span></span> | <span data-ttu-id="002ec-174">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="002ec-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="002ec-175">1.7</span><span class="sxs-lookup"><span data-stu-id="002ec-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="002ec-176">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-176">String</span></span> | <span data-ttu-id="002ec-177">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="002ec-177">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="002ec-178">1,5</span><span class="sxs-lookup"><span data-stu-id="002ec-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="002ec-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-179">String</span></span> | <span data-ttu-id="002ec-180">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="002ec-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="002ec-181">1.7</span><span class="sxs-lookup"><span data-stu-id="002ec-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="002ec-182">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-182">String</span></span> | <span data-ttu-id="002ec-183">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="002ec-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="002ec-184">1.7</span><span class="sxs-lookup"><span data-stu-id="002ec-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="002ec-185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="002ec-185">Requirements</span></span>

|<span data-ttu-id="002ec-186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="002ec-186">Requirement</span></span>| <span data-ttu-id="002ec-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="002ec-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="002ec-188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="002ec-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="002ec-189">1,5</span><span class="sxs-lookup"><span data-stu-id="002ec-189">1.5</span></span> |
|[<span data-ttu-id="002ec-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="002ec-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="002ec-191">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="002ec-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="002ec-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="002ec-192">SourceProperty :String</span></span>

<span data-ttu-id="002ec-193">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="002ec-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="002ec-194">Type :</span><span class="sxs-lookup"><span data-stu-id="002ec-194">Type:</span></span>

*   <span data-ttu-id="002ec-195">String</span><span class="sxs-lookup"><span data-stu-id="002ec-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="002ec-196">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="002ec-196">Properties:</span></span>

|<span data-ttu-id="002ec-197">Nom</span><span class="sxs-lookup"><span data-stu-id="002ec-197">Name</span></span>| <span data-ttu-id="002ec-198">Type</span><span class="sxs-lookup"><span data-stu-id="002ec-198">Type</span></span>| <span data-ttu-id="002ec-199">Description</span><span class="sxs-lookup"><span data-stu-id="002ec-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="002ec-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="002ec-200">String</span></span>|<span data-ttu-id="002ec-201">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="002ec-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="002ec-202">String</span><span class="sxs-lookup"><span data-stu-id="002ec-202">String</span></span>|<span data-ttu-id="002ec-203">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="002ec-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="002ec-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="002ec-204">Requirements</span></span>

|<span data-ttu-id="002ec-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="002ec-205">Requirement</span></span>| <span data-ttu-id="002ec-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="002ec-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="002ec-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="002ec-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="002ec-208">1.0</span><span class="sxs-lookup"><span data-stu-id="002ec-208">1.0</span></span>|
|[<span data-ttu-id="002ec-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="002ec-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="002ec-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="002ec-210">Compose or read</span></span>|