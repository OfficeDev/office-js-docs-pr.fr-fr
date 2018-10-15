 

# <a name="office"></a><span data-ttu-id="936af-101">Office</span><span class="sxs-lookup"><span data-stu-id="936af-101">Office</span></span>

<span data-ttu-id="936af-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="936af-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="936af-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="936af-104">Requirements</span></span>

|<span data-ttu-id="936af-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="936af-105">Requirement</span></span>| <span data-ttu-id="936af-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="936af-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="936af-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="936af-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="936af-108">1.0</span><span class="sxs-lookup"><span data-stu-id="936af-108">1.0</span></span>|
|[<span data-ttu-id="936af-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="936af-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="936af-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="936af-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="936af-111">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="936af-111">Namespaces</span></span>

<span data-ttu-id="936af-112">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="936af-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="936af-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="936af-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="936af-114">Membres</span><span class="sxs-lookup"><span data-stu-id="936af-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="936af-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="936af-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="936af-116">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="936af-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="936af-117">Type :</span><span class="sxs-lookup"><span data-stu-id="936af-117">Type:</span></span>

*   <span data-ttu-id="936af-118">String</span><span class="sxs-lookup"><span data-stu-id="936af-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="936af-119">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="936af-119">Properties:</span></span>

|<span data-ttu-id="936af-120">Nom</span><span class="sxs-lookup"><span data-stu-id="936af-120">Name</span></span>| <span data-ttu-id="936af-121">Type</span><span class="sxs-lookup"><span data-stu-id="936af-121">Type</span></span>| <span data-ttu-id="936af-122">Description</span><span class="sxs-lookup"><span data-stu-id="936af-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="936af-123">String</span><span class="sxs-lookup"><span data-stu-id="936af-123">String</span></span>|<span data-ttu-id="936af-124">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="936af-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="936af-125">String</span><span class="sxs-lookup"><span data-stu-id="936af-125">String</span></span>|<span data-ttu-id="936af-126">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="936af-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="936af-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="936af-127">Requirements</span></span>

|<span data-ttu-id="936af-128">Condition requise</span><span class="sxs-lookup"><span data-stu-id="936af-128">Requirement</span></span>| <span data-ttu-id="936af-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="936af-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="936af-130">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="936af-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="936af-131">1.0</span><span class="sxs-lookup"><span data-stu-id="936af-131">1.0</span></span>|
|[<span data-ttu-id="936af-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="936af-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="936af-133">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="936af-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="936af-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="936af-134">CoercionType :String</span></span>

<span data-ttu-id="936af-135">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="936af-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="936af-136">Type :</span><span class="sxs-lookup"><span data-stu-id="936af-136">Type:</span></span>

*   <span data-ttu-id="936af-137">String</span><span class="sxs-lookup"><span data-stu-id="936af-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="936af-138">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="936af-138">Properties:</span></span>

|<span data-ttu-id="936af-139">Nom</span><span class="sxs-lookup"><span data-stu-id="936af-139">Name</span></span>| <span data-ttu-id="936af-140">Type</span><span class="sxs-lookup"><span data-stu-id="936af-140">Type</span></span>| <span data-ttu-id="936af-141">Description</span><span class="sxs-lookup"><span data-stu-id="936af-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="936af-142">String</span><span class="sxs-lookup"><span data-stu-id="936af-142">String</span></span>|<span data-ttu-id="936af-143">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="936af-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="936af-144">String</span><span class="sxs-lookup"><span data-stu-id="936af-144">String</span></span>|<span data-ttu-id="936af-145">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="936af-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="936af-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="936af-146">Requirements</span></span>

|<span data-ttu-id="936af-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="936af-147">Requirement</span></span>| <span data-ttu-id="936af-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="936af-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="936af-149">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="936af-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="936af-150">1.0</span><span class="sxs-lookup"><span data-stu-id="936af-150">1.0</span></span>|
|[<span data-ttu-id="936af-151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="936af-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="936af-152">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="936af-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="936af-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="936af-153">SourceProperty :String</span></span>

<span data-ttu-id="936af-154">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="936af-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="936af-155">Type :</span><span class="sxs-lookup"><span data-stu-id="936af-155">Type:</span></span>

*   <span data-ttu-id="936af-156">String</span><span class="sxs-lookup"><span data-stu-id="936af-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="936af-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="936af-157">Properties:</span></span>

|<span data-ttu-id="936af-158">Nom</span><span class="sxs-lookup"><span data-stu-id="936af-158">Name</span></span>| <span data-ttu-id="936af-159">Type</span><span class="sxs-lookup"><span data-stu-id="936af-159">Type</span></span>| <span data-ttu-id="936af-160">Description</span><span class="sxs-lookup"><span data-stu-id="936af-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="936af-161">String</span><span class="sxs-lookup"><span data-stu-id="936af-161">String</span></span>|<span data-ttu-id="936af-162">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="936af-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="936af-163">String</span><span class="sxs-lookup"><span data-stu-id="936af-163">String</span></span>|<span data-ttu-id="936af-164">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="936af-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="936af-165">Condition requise</span><span class="sxs-lookup"><span data-stu-id="936af-165">Requirements</span></span>

|<span data-ttu-id="936af-166">Condition requise</span><span class="sxs-lookup"><span data-stu-id="936af-166">Requirement</span></span>| <span data-ttu-id="936af-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="936af-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="936af-168">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="936af-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="936af-169">1.0</span><span class="sxs-lookup"><span data-stu-id="936af-169">1.0</span></span>|
|[<span data-ttu-id="936af-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="936af-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="936af-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="936af-171">Compose or read</span></span>|