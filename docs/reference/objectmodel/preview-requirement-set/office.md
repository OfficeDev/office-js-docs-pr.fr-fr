 

# <a name="office"></a><span data-ttu-id="4e6f8-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="4e6f8-101">Office</span></span>

<span data-ttu-id="4e6f8-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4e6f8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e6f8-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4e6f8-104">Requirements</span></span>

|<span data-ttu-id="4e6f8-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4e6f8-105">Requirement</span></span>| <span data-ttu-id="4e6f8-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="4e6f8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e6f8-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e6f8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4e6f8-108">1.0</span></span>|
|[<span data-ttu-id="4e6f8-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4e6f8-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e6f8-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4e6f8-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4e6f8-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4e6f8-111">Members and methods</span></span>

| <span data-ttu-id="4e6f8-112">Membre</span><span class="sxs-lookup"><span data-stu-id="4e6f8-112">Member</span></span> | <span data-ttu-id="4e6f8-113">Type</span><span class="sxs-lookup"><span data-stu-id="4e6f8-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4e6f8-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4e6f8-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4e6f8-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4e6f8-115">Member</span></span> |
| [<span data-ttu-id="4e6f8-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4e6f8-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4e6f8-117">Membre</span><span class="sxs-lookup"><span data-stu-id="4e6f8-117">Member</span></span> |
| [<span data-ttu-id="4e6f8-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4e6f8-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4e6f8-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4e6f8-119">Member</span></span> |
| [<span data-ttu-id="4e6f8-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4e6f8-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4e6f8-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4e6f8-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4e6f8-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4e6f8-122">Namespaces</span></span>

<span data-ttu-id="4e6f8-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4e6f8-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4e6f8-125">Membres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4e6f8-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4e6f8-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4e6f8-128">Type :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-128">Type:</span></span>

*   <span data-ttu-id="4e6f8-129">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e6f8-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-130">Properties:</span></span>

|<span data-ttu-id="4e6f8-131">Nom</span><span class="sxs-lookup"><span data-stu-id="4e6f8-131">Name</span></span>| <span data-ttu-id="4e6f8-132">Type</span><span class="sxs-lookup"><span data-stu-id="4e6f8-132">Type</span></span>| <span data-ttu-id="4e6f8-133">Description</span><span class="sxs-lookup"><span data-stu-id="4e6f8-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4e6f8-134">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-134">String</span></span>|<span data-ttu-id="4e6f8-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4e6f8-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-136">String</span></span>|<span data-ttu-id="4e6f8-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e6f8-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4e6f8-138">Requirements</span></span>

|<span data-ttu-id="4e6f8-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4e6f8-139">Requirement</span></span>| <span data-ttu-id="4e6f8-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="4e6f8-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e6f8-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e6f8-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4e6f8-142">1.0</span></span>|
|[<span data-ttu-id="4e6f8-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4e6f8-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e6f8-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4e6f8-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4e6f8-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-145">CoercionType :String</span></span>

<span data-ttu-id="4e6f8-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4e6f8-147">Type :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-147">Type:</span></span>

*   <span data-ttu-id="4e6f8-148">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e6f8-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-149">Properties:</span></span>

|<span data-ttu-id="4e6f8-150">Nom</span><span class="sxs-lookup"><span data-stu-id="4e6f8-150">Name</span></span>| <span data-ttu-id="4e6f8-151">Type</span><span class="sxs-lookup"><span data-stu-id="4e6f8-151">Type</span></span>| <span data-ttu-id="4e6f8-152">Description</span><span class="sxs-lookup"><span data-stu-id="4e6f8-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4e6f8-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-153">String</span></span>|<span data-ttu-id="4e6f8-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4e6f8-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-155">String</span></span>|<span data-ttu-id="4e6f8-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e6f8-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4e6f8-157">Requirements</span></span>

|<span data-ttu-id="4e6f8-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4e6f8-158">Requirement</span></span>| <span data-ttu-id="4e6f8-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="4e6f8-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e6f8-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e6f8-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4e6f8-161">1.0</span></span>|
|[<span data-ttu-id="4e6f8-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4e6f8-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e6f8-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4e6f8-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4e6f8-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-164">EventType :String</span></span>

<span data-ttu-id="4e6f8-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4e6f8-166">Type :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-166">Type:</span></span>

*   <span data-ttu-id="4e6f8-167">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e6f8-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-168">Properties:</span></span>

| <span data-ttu-id="4e6f8-169">Nom</span><span class="sxs-lookup"><span data-stu-id="4e6f8-169">Name</span></span> | <span data-ttu-id="4e6f8-170">Type</span><span class="sxs-lookup"><span data-stu-id="4e6f8-170">Type</span></span> | <span data-ttu-id="4e6f8-171">Description</span><span class="sxs-lookup"><span data-stu-id="4e6f8-171">Description</span></span> | <span data-ttu-id="4e6f8-172">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="4e6f8-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="4e6f8-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-173">String</span></span> | <span data-ttu-id="4e6f8-174">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4e6f8-175">1.7</span><span class="sxs-lookup"><span data-stu-id="4e6f8-175">-17</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4e6f8-176">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-176">String</span></span> | <span data-ttu-id="4e6f8-177">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-177">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4e6f8-178">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4e6f8-178">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="4e6f8-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-179">String</span></span> | <span data-ttu-id="4e6f8-180">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-180">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4e6f8-181">1,5</span><span class="sxs-lookup"><span data-stu-id="4e6f8-181">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="4e6f8-182">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-182">String</span></span> | <span data-ttu-id="4e6f8-183">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-183">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="4e6f8-184">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4e6f8-184">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4e6f8-185">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-185">String</span></span> | <span data-ttu-id="4e6f8-186">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-186">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4e6f8-187">1.7</span><span class="sxs-lookup"><span data-stu-id="4e6f8-187">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4e6f8-188">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-188">String</span></span> | <span data-ttu-id="4e6f8-189">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-189">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4e6f8-190">1.7</span><span class="sxs-lookup"><span data-stu-id="4e6f8-190">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e6f8-191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4e6f8-191">Requirements</span></span>

|<span data-ttu-id="4e6f8-192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4e6f8-192">Requirement</span></span>| <span data-ttu-id="4e6f8-193">Valeur</span><span class="sxs-lookup"><span data-stu-id="4e6f8-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e6f8-194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e6f8-195">1,5</span><span class="sxs-lookup"><span data-stu-id="4e6f8-195">1.5</span></span> |
|[<span data-ttu-id="4e6f8-196">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4e6f8-196">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e6f8-197">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4e6f8-197">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4e6f8-198">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-198">SourceProperty :String</span></span>

<span data-ttu-id="4e6f8-199">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-199">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4e6f8-200">Type :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-200">Type:</span></span>

*   <span data-ttu-id="4e6f8-201">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-201">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e6f8-202">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4e6f8-202">Properties:</span></span>

|<span data-ttu-id="4e6f8-203">Nom</span><span class="sxs-lookup"><span data-stu-id="4e6f8-203">Name</span></span>| <span data-ttu-id="4e6f8-204">Type</span><span class="sxs-lookup"><span data-stu-id="4e6f8-204">Type</span></span>| <span data-ttu-id="4e6f8-205">Description</span><span class="sxs-lookup"><span data-stu-id="4e6f8-205">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4e6f8-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4e6f8-206">String</span></span>|<span data-ttu-id="4e6f8-207">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-207">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4e6f8-208">String</span><span class="sxs-lookup"><span data-stu-id="4e6f8-208">String</span></span>|<span data-ttu-id="4e6f8-209">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4e6f8-209">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e6f8-210">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4e6f8-210">Requirements</span></span>

|<span data-ttu-id="4e6f8-211">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4e6f8-211">Requirement</span></span>| <span data-ttu-id="4e6f8-212">Valeur</span><span class="sxs-lookup"><span data-stu-id="4e6f8-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e6f8-213">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4e6f8-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e6f8-214">1.0</span><span class="sxs-lookup"><span data-stu-id="4e6f8-214">1.0</span></span>|
|[<span data-ttu-id="4e6f8-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4e6f8-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e6f8-216">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4e6f8-216">Compose or read</span></span>|