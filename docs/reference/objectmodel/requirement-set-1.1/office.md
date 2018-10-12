 

# <a name="office"></a><span data-ttu-id="b588a-101">Office</span><span class="sxs-lookup"><span data-stu-id="b588a-101">Office</span></span>

<span data-ttu-id="b588a-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b588a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b588a-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b588a-104">Requirements</span></span>

|<span data-ttu-id="b588a-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="b588a-105">Requirement</span></span>| <span data-ttu-id="b588a-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="b588a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b588a-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b588a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b588a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b588a-108">1.0</span></span>|
|[<span data-ttu-id="b588a-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b588a-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b588a-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b588a-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="b588a-111">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="b588a-111">Namespaces</span></span>

<span data-ttu-id="b588a-112">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b588a-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b588a-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b588a-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b588a-114">Membres</span><span class="sxs-lookup"><span data-stu-id="b588a-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b588a-115">AsyncResultStatus : Chaîne</span><span class="sxs-lookup"><span data-stu-id="b588a-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="b588a-116">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b588a-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b588a-117">Type :</span><span class="sxs-lookup"><span data-stu-id="b588a-117">Type:</span></span>

*   <span data-ttu-id="b588a-118">String</span><span class="sxs-lookup"><span data-stu-id="b588a-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b588a-119">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b588a-119">Properties:</span></span>

|<span data-ttu-id="b588a-120">Nom</span><span class="sxs-lookup"><span data-stu-id="b588a-120">Name</span></span>| <span data-ttu-id="b588a-121">Type</span><span class="sxs-lookup"><span data-stu-id="b588a-121">Type</span></span>| <span data-ttu-id="b588a-122">Description</span><span class="sxs-lookup"><span data-stu-id="b588a-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b588a-123">String</span><span class="sxs-lookup"><span data-stu-id="b588a-123">String</span></span>|<span data-ttu-id="b588a-124">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="b588a-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b588a-125">String</span><span class="sxs-lookup"><span data-stu-id="b588a-125">String</span></span>|<span data-ttu-id="b588a-126">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="b588a-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b588a-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b588a-127">Requirements</span></span>

|<span data-ttu-id="b588a-128">Condition requise</span><span class="sxs-lookup"><span data-stu-id="b588a-128">Requirement</span></span>| <span data-ttu-id="b588a-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="b588a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="b588a-130">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b588a-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b588a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="b588a-131">1.0</span></span>|
|[<span data-ttu-id="b588a-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b588a-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b588a-133">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b588a-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="b588a-134">CoercionType : Chaîne</span><span class="sxs-lookup"><span data-stu-id="b588a-134">CoercionType :String</span></span>

<span data-ttu-id="b588a-135">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b588a-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b588a-136">Type :</span><span class="sxs-lookup"><span data-stu-id="b588a-136">Type:</span></span>

*   <span data-ttu-id="b588a-137">String</span><span class="sxs-lookup"><span data-stu-id="b588a-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b588a-138">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b588a-138">Properties:</span></span>

|<span data-ttu-id="b588a-139">Nom</span><span class="sxs-lookup"><span data-stu-id="b588a-139">Name</span></span>| <span data-ttu-id="b588a-140">Type</span><span class="sxs-lookup"><span data-stu-id="b588a-140">Type</span></span>| <span data-ttu-id="b588a-141">Description</span><span class="sxs-lookup"><span data-stu-id="b588a-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b588a-142">String</span><span class="sxs-lookup"><span data-stu-id="b588a-142">String</span></span>|<span data-ttu-id="b588a-143">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="b588a-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b588a-144">String</span><span class="sxs-lookup"><span data-stu-id="b588a-144">String</span></span>|<span data-ttu-id="b588a-145">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="b588a-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b588a-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b588a-146">Requirements</span></span>

|<span data-ttu-id="b588a-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="b588a-147">Requirement</span></span>| <span data-ttu-id="b588a-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="b588a-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b588a-149">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b588a-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b588a-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b588a-150">1.0</span></span>|
|[<span data-ttu-id="b588a-151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b588a-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b588a-152">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b588a-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="b588a-153">SourceProperty : Chaîne</span><span class="sxs-lookup"><span data-stu-id="b588a-153">SourceProperty :String</span></span>

<span data-ttu-id="b588a-154">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b588a-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b588a-155">Type :</span><span class="sxs-lookup"><span data-stu-id="b588a-155">Type:</span></span>

*   <span data-ttu-id="b588a-156">String</span><span class="sxs-lookup"><span data-stu-id="b588a-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b588a-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b588a-157">Properties:</span></span>

|<span data-ttu-id="b588a-158">Nom</span><span class="sxs-lookup"><span data-stu-id="b588a-158">Name</span></span>| <span data-ttu-id="b588a-159">Type</span><span class="sxs-lookup"><span data-stu-id="b588a-159">Type</span></span>| <span data-ttu-id="b588a-160">Description</span><span class="sxs-lookup"><span data-stu-id="b588a-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b588a-161">String</span><span class="sxs-lookup"><span data-stu-id="b588a-161">String</span></span>|<span data-ttu-id="b588a-162">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="b588a-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b588a-163">String</span><span class="sxs-lookup"><span data-stu-id="b588a-163">String</span></span>|<span data-ttu-id="b588a-164">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="b588a-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b588a-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b588a-165">Requirements</span></span>

|<span data-ttu-id="b588a-166">Condition requise</span><span class="sxs-lookup"><span data-stu-id="b588a-166">Requirement</span></span>| <span data-ttu-id="b588a-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="b588a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b588a-168">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b588a-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b588a-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b588a-169">1.0</span></span>|
|[<span data-ttu-id="b588a-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b588a-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b588a-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b588a-171">Compose or read</span></span>|