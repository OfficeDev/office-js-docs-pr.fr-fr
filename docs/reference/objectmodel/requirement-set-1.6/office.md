 

# <a name="office"></a><span data-ttu-id="4ec18-101">Office</span><span class="sxs-lookup"><span data-stu-id="4ec18-101">Office</span></span>

<span data-ttu-id="4ec18-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4ec18-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ec18-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ec18-104">Requirements</span></span>

|<span data-ttu-id="4ec18-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4ec18-105">Requirement</span></span>| <span data-ttu-id="4ec18-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ec18-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ec18-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ec18-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ec18-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4ec18-108">1.0</span></span>|
|[<span data-ttu-id="4ec18-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ec18-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ec18-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ec18-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4ec18-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4ec18-111">Members and methods</span></span>

| <span data-ttu-id="4ec18-112">Membre</span><span class="sxs-lookup"><span data-stu-id="4ec18-112">Member</span></span> | <span data-ttu-id="4ec18-113">Type</span><span class="sxs-lookup"><span data-stu-id="4ec18-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4ec18-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4ec18-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4ec18-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4ec18-115">Member</span></span> |
| [<span data-ttu-id="4ec18-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4ec18-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4ec18-117">Membre</span><span class="sxs-lookup"><span data-stu-id="4ec18-117">Member</span></span> |
| [<span data-ttu-id="4ec18-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4ec18-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4ec18-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4ec18-119">Member</span></span> |
| [<span data-ttu-id="4ec18-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4ec18-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4ec18-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4ec18-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4ec18-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4ec18-122">Namespaces</span></span>

<span data-ttu-id="4ec18-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ec18-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4ec18-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4ec18-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4ec18-125">Membres</span><span class="sxs-lookup"><span data-stu-id="4ec18-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4ec18-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4ec18-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4ec18-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4ec18-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4ec18-128">Type :</span><span class="sxs-lookup"><span data-stu-id="4ec18-128">Type:</span></span>

*   <span data-ttu-id="4ec18-129">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ec18-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ec18-130">Properties:</span></span>

|<span data-ttu-id="4ec18-131">Nom</span><span class="sxs-lookup"><span data-stu-id="4ec18-131">Name</span></span>| <span data-ttu-id="4ec18-132">Type</span><span class="sxs-lookup"><span data-stu-id="4ec18-132">Type</span></span>| <span data-ttu-id="4ec18-133">Description</span><span class="sxs-lookup"><span data-stu-id="4ec18-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4ec18-134">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-134">String</span></span>|<span data-ttu-id="4ec18-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4ec18-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4ec18-136">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-136">String</span></span>|<span data-ttu-id="4ec18-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4ec18-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ec18-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ec18-138">Requirements</span></span>

|<span data-ttu-id="4ec18-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4ec18-139">Requirement</span></span>| <span data-ttu-id="4ec18-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ec18-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ec18-141">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ec18-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ec18-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4ec18-142">1.0</span></span>|
|[<span data-ttu-id="4ec18-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ec18-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ec18-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ec18-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4ec18-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4ec18-145">CoercionType :String</span></span>

<span data-ttu-id="4ec18-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4ec18-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ec18-147">Type :</span><span class="sxs-lookup"><span data-stu-id="4ec18-147">Type:</span></span>

*   <span data-ttu-id="4ec18-148">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ec18-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ec18-149">Properties:</span></span>

|<span data-ttu-id="4ec18-150">Nom</span><span class="sxs-lookup"><span data-stu-id="4ec18-150">Name</span></span>| <span data-ttu-id="4ec18-151">Type</span><span class="sxs-lookup"><span data-stu-id="4ec18-151">Type</span></span>| <span data-ttu-id="4ec18-152">Description</span><span class="sxs-lookup"><span data-stu-id="4ec18-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4ec18-153">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-153">String</span></span>|<span data-ttu-id="4ec18-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4ec18-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4ec18-155">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-155">String</span></span>|<span data-ttu-id="4ec18-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4ec18-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ec18-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ec18-157">Requirements</span></span>

|<span data-ttu-id="4ec18-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4ec18-158">Requirement</span></span>| <span data-ttu-id="4ec18-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ec18-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ec18-160">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ec18-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ec18-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4ec18-161">1.0</span></span>|
|[<span data-ttu-id="4ec18-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ec18-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ec18-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ec18-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4ec18-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4ec18-164">EventType :String</span></span>

<span data-ttu-id="4ec18-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4ec18-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4ec18-166">Type :</span><span class="sxs-lookup"><span data-stu-id="4ec18-166">Type:</span></span>

*   <span data-ttu-id="4ec18-167">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ec18-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ec18-168">Properties:</span></span>

| <span data-ttu-id="4ec18-169">Nom</span><span class="sxs-lookup"><span data-stu-id="4ec18-169">Name</span></span> | <span data-ttu-id="4ec18-170">Type</span><span class="sxs-lookup"><span data-stu-id="4ec18-170">Type</span></span> | <span data-ttu-id="4ec18-171">Description</span><span class="sxs-lookup"><span data-stu-id="4ec18-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="4ec18-172">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-172">String</span></span> | <span data-ttu-id="4ec18-173">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4ec18-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4ec18-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ec18-174">Requirements</span></span>

|<span data-ttu-id="4ec18-175">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4ec18-175">Requirement</span></span>| <span data-ttu-id="4ec18-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ec18-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ec18-177">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ec18-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ec18-178">1,5</span><span class="sxs-lookup"><span data-stu-id="4ec18-178">1.5</span></span> |
|[<span data-ttu-id="4ec18-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ec18-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ec18-180">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ec18-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4ec18-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4ec18-181">SourceProperty :String</span></span>

<span data-ttu-id="4ec18-182">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4ec18-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ec18-183">Type :</span><span class="sxs-lookup"><span data-stu-id="4ec18-183">Type:</span></span>

*   <span data-ttu-id="4ec18-184">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ec18-185">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ec18-185">Properties:</span></span>

|<span data-ttu-id="4ec18-186">Nom</span><span class="sxs-lookup"><span data-stu-id="4ec18-186">Name</span></span>| <span data-ttu-id="4ec18-187">Type</span><span class="sxs-lookup"><span data-stu-id="4ec18-187">Type</span></span>| <span data-ttu-id="4ec18-188">Description</span><span class="sxs-lookup"><span data-stu-id="4ec18-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4ec18-189">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-189">String</span></span>|<span data-ttu-id="4ec18-190">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4ec18-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4ec18-191">Chaîne​</span><span class="sxs-lookup"><span data-stu-id="4ec18-191">String</span></span>|<span data-ttu-id="4ec18-192">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4ec18-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ec18-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ec18-193">Requirements</span></span>

|<span data-ttu-id="4ec18-194">Condition requise</span><span class="sxs-lookup"><span data-stu-id="4ec18-194">Requirement</span></span>| <span data-ttu-id="4ec18-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ec18-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ec18-196">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ec18-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ec18-197">1.0</span><span class="sxs-lookup"><span data-stu-id="4ec18-197">1.0</span></span>|
|[<span data-ttu-id="4ec18-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ec18-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ec18-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ec18-199">Compose or read</span></span>|