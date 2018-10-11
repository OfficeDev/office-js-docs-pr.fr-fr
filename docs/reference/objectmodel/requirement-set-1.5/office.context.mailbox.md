
# <a name="mailbox"></a><span data-ttu-id="1d372-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="1d372-101">mailbox</span></span>

### <span data-ttu-id="1d372-p101">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="1d372-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="1d372-104">Donne accès au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="1d372-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d372-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-105">Requirements</span></span>

|<span data-ttu-id="1d372-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-106">Requirement</span></span>| <span data-ttu-id="1d372-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-108">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-109">1.0</span></span>|
|[<span data-ttu-id="1d372-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="1d372-111">Restricted</span></span>|
|[<span data-ttu-id="1d372-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1d372-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="1d372-114">Members and methods</span></span>

| <span data-ttu-id="1d372-115">Membre</span><span class="sxs-lookup"><span data-stu-id="1d372-115">Member</span></span> | <span data-ttu-id="1d372-116">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1d372-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="1d372-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="1d372-118">Membre</span><span class="sxs-lookup"><span data-stu-id="1d372-118">Member</span></span> |
| [<span data-ttu-id="1d372-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="1d372-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="1d372-120">Membre</span><span class="sxs-lookup"><span data-stu-id="1d372-120">Member</span></span> |
| [<span data-ttu-id="1d372-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1d372-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1d372-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-122">Method</span></span> |
| [<span data-ttu-id="1d372-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="1d372-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="1d372-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-124">Method</span></span> |
| [<span data-ttu-id="1d372-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1d372-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="1d372-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-126">Method</span></span> |
| [<span data-ttu-id="1d372-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="1d372-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="1d372-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-128">Method</span></span> |
| [<span data-ttu-id="1d372-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="1d372-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="1d372-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-130">Method</span></span> |
| [<span data-ttu-id="1d372-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1d372-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="1d372-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-132">Method</span></span> |
| [<span data-ttu-id="1d372-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="1d372-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="1d372-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-134">Method</span></span> |
| [<span data-ttu-id="1d372-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1d372-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="1d372-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-136">Method</span></span> |
| [<span data-ttu-id="1d372-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d372-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="1d372-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-138">Method</span></span> |
| [<span data-ttu-id="1d372-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d372-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="1d372-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-140">Method</span></span> |
| [<span data-ttu-id="1d372-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d372-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="1d372-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-142">Method</span></span> |
| [<span data-ttu-id="1d372-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="1d372-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="1d372-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="1d372-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1d372-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="1d372-145">Namespaces</span></span>

<span data-ttu-id="1d372-146">[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d372-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1d372-147">[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d372-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1d372-148">[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d372-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1d372-149">Membres</span><span class="sxs-lookup"><span data-stu-id="1d372-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1d372-150">ewsUrl : chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-150">ewsUrl :String</span></span>

<span data-ttu-id="1d372-p102">Obtient l’URL du point de terminaison des services web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1d372-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-153">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-153">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-p103">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Vous pouvez par exemple créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1d372-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1d372-156">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `ewsUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="1d372-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="1d372-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1d372-159">Type :</span><span class="sxs-lookup"><span data-stu-id="1d372-159">Type:</span></span>

*   <span data-ttu-id="1d372-160">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d372-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-161">Requirements</span></span>

|<span data-ttu-id="1d372-162">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-162">Requirement</span></span>| <span data-ttu-id="1d372-163">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-164">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-165">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-165">1.0</span></span>|
|[<span data-ttu-id="1d372-166">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-167">ReadItem</span></span>|
|[<span data-ttu-id="1d372-168">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-169">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="1d372-170">restUrl : chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-170">restUrl :String</span></span>

<span data-ttu-id="1d372-171">Obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="1d372-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="1d372-172">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1d372-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="1d372-173">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="1d372-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="1d372-p105">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-176">Les clients Outlook connectés aux installations locales d’Exchange 2016 (ou version ultérieure) avec une URL REST personnalisée configurée renvoient une valeur non valide pour `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="1d372-176">Note: Outlook clients connected to on-premises installations of Exchange 2016 with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d372-177">Type :</span><span class="sxs-lookup"><span data-stu-id="1d372-177">Type:</span></span>

*   <span data-ttu-id="1d372-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d372-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-179">Requirements</span></span>

|<span data-ttu-id="1d372-180">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-180">Requirement</span></span>| <span data-ttu-id="1d372-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-182">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-183">1.5</span><span class="sxs-lookup"><span data-stu-id="1d372-183">1.5</span></span> |
|[<span data-ttu-id="1d372-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-185">ReadItem</span></span>|
|[<span data-ttu-id="1d372-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="1d372-188">Méthodes</span><span class="sxs-lookup"><span data-stu-id="1d372-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1d372-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d372-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1d372-190">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="1d372-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1d372-p106">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé lorsque l’utilisateur sélectionne un nouvel élément. Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’interface utilisateur du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1d372-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-193">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-193">Parameters:</span></span>

| <span data-ttu-id="1d372-194">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-194">Name</span></span> | <span data-ttu-id="1d372-195">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-195">Type</span></span> | <span data-ttu-id="1d372-196">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d372-196">Attributes</span></span> | <span data-ttu-id="1d372-197">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1d372-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1d372-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1d372-199">L’événement qui doit invoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="1d372-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1d372-200">Fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-200">Function</span></span> || <span data-ttu-id="1d372-p107">La fonction permettant de gérer l’événement. Cette fonction doit accepter un seul paramètre, qui est un d’objet textuel. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1d372-204">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-204">Object</span></span> | <span data-ttu-id="1d372-205">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-205">&lt;optional&gt;</span></span> | <span data-ttu-id="1d372-206">Littéral d'objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1d372-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d372-207">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-207">Object</span></span> | <span data-ttu-id="1d372-208">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-208">&lt;optional&gt;</span></span> | <span data-ttu-id="1d372-209">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1d372-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1d372-210">fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-210">function</span></span>| <span data-ttu-id="1d372-211">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-211">&lt;optional&gt;</span></span>|<span data-ttu-id="1d372-212">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d372-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-213">Requirements</span></span>

|<span data-ttu-id="1d372-214">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-214">Requirement</span></span>| <span data-ttu-id="1d372-215">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-216">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-217">1.5</span><span class="sxs-lookup"><span data-stu-id="1d372-217">1.5</span></span> |
|[<span data-ttu-id="1d372-218">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-219">ReadItem</span></span> |
|[<span data-ttu-id="1d372-220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-221">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="1d372-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d372-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1d372-224">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="1d372-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-225">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-p108">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="1d372-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-228">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-228">Parameters:</span></span>

|<span data-ttu-id="1d372-229">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-229">Name</span></span>| <span data-ttu-id="1d372-230">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-230">Type</span></span>| <span data-ttu-id="1d372-231">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d372-232">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-232">String</span></span>|<span data-ttu-id="1d372-233">Un ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="1d372-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="1d372-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1d372-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="1d372-235">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="1d372-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-236">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-236">Requirements</span></span>

|<span data-ttu-id="1d372-237">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-237">Requirement</span></span>| <span data-ttu-id="1d372-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-239">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-240">1.3</span><span class="sxs-lookup"><span data-stu-id="1d372-240">1.3</span></span>|
|[<span data-ttu-id="1d372-241">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-242">Restreint</span><span class="sxs-lookup"><span data-stu-id="1d372-242">Restricted</span></span>|
|[<span data-ttu-id="1d372-243">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-244">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d372-245">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1d372-245">Returns:</span></span>

<span data-ttu-id="1d372-246">Type : Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1d372-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="1d372-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="1d372-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="1d372-249">Obtient un dictionnaire contenant des informations d’heure en heure locale client.</span><span class="sxs-lookup"><span data-stu-id="1d372-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1d372-p109">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur client ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1d372-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1d372-p110">Si l’application de messagerie s'exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="1d372-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-255">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-255">Parameters:</span></span>

|<span data-ttu-id="1d372-256">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-256">Name</span></span>| <span data-ttu-id="1d372-257">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-257">Type</span></span>| <span data-ttu-id="1d372-258">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1d372-259">Date</span><span class="sxs-lookup"><span data-stu-id="1d372-259">Date</span></span>|<span data-ttu-id="1d372-260">Un objet Date</span><span class="sxs-lookup"><span data-stu-id="1d372-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-261">Requirements</span></span>

|<span data-ttu-id="1d372-262">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-262">Requirement</span></span>| <span data-ttu-id="1d372-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-264">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-265">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-265">1.0</span></span>|
|[<span data-ttu-id="1d372-266">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-267">ReadItem</span></span>|
|[<span data-ttu-id="1d372-268">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-269">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d372-270">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1d372-270">Returns:</span></span>

<span data-ttu-id="1d372-271">Type : [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="1d372-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="1d372-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d372-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1d372-273">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="1d372-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-274">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-p111">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS dans un format adapté à REST.</span><span class="sxs-lookup"><span data-stu-id="1d372-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-277">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-277">Parameters:</span></span>

|<span data-ttu-id="1d372-278">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-278">Name</span></span>| <span data-ttu-id="1d372-279">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-279">Type</span></span>| <span data-ttu-id="1d372-280">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d372-281">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-281">String</span></span>|<span data-ttu-id="1d372-282">Un ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="1d372-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="1d372-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1d372-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="1d372-284">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="1d372-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-285">Requirements</span></span>

|<span data-ttu-id="1d372-286">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-286">Requirement</span></span>| <span data-ttu-id="1d372-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-288">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-289">1.3</span><span class="sxs-lookup"><span data-stu-id="1d372-289">1.3</span></span>|
|[<span data-ttu-id="1d372-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-291">Restreint</span><span class="sxs-lookup"><span data-stu-id="1d372-291">Restricted</span></span>|
|[<span data-ttu-id="1d372-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-293">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d372-294">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1d372-294">Returns:</span></span>

<span data-ttu-id="1d372-295">Type : Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1d372-296">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1d372-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1d372-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1d372-298">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="1d372-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1d372-299">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="1d372-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-300">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-300">Parameters:</span></span>

|<span data-ttu-id="1d372-301">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-301">Name</span></span>| <span data-ttu-id="1d372-302">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-302">Type</span></span>| <span data-ttu-id="1d372-303">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1d372-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1d372-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="1d372-305">Valeur en heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="1d372-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-306">Requirements</span></span>

|<span data-ttu-id="1d372-307">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-307">Requirement</span></span>| <span data-ttu-id="1d372-308">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-309">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-310">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-310">1.0</span></span>|
|[<span data-ttu-id="1d372-311">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-312">ReadItem</span></span>|
|[<span data-ttu-id="1d372-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-314">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d372-315">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1d372-315">Returns:</span></span>

<span data-ttu-id="1d372-316">Un objet Date avec l’heure exprimée en UTC.</span><span class="sxs-lookup"><span data-stu-id="1d372-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="1d372-317">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d372-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d372-318">Date</span><span class="sxs-lookup"><span data-stu-id="1d372-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="1d372-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1d372-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1d372-320">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="1d372-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-321">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-322">La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les équipements mobiles.</span><span class="sxs-lookup"><span data-stu-id="1d372-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1d372-p112">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="1d372-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1d372-325">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1d372-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1d372-326">Si l’identificateur d’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur n'est retourné.</span><span class="sxs-lookup"><span data-stu-id="1d372-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-327">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-327">Parameters:</span></span>

|<span data-ttu-id="1d372-328">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-328">Name</span></span>| <span data-ttu-id="1d372-329">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-329">Type</span></span>| <span data-ttu-id="1d372-330">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d372-331">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-331">String</span></span>|<span data-ttu-id="1d372-332">L'identificateur EWS (services web Exchange) pour un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="1d372-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-333">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-333">Requirements</span></span>

|<span data-ttu-id="1d372-334">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-334">Requirement</span></span>| <span data-ttu-id="1d372-335">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-336">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-337">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-337">1.0</span></span>|
|[<span data-ttu-id="1d372-338">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-339">ReadItem</span></span>|
|[<span data-ttu-id="1d372-340">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-341">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-342">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="1d372-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1d372-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1d372-344">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="1d372-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-345">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-346">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre sur le bureau, ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="1d372-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1d372-347">Dans Outlook Web App, cette méthode n'ouvre le formulaire indiqué que si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1d372-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1d372-348">Si l’identificateur d’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.</span><span class="sxs-lookup"><span data-stu-id="1d372-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1d372-p113">N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1d372-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-351">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-351">Parameters:</span></span>

|<span data-ttu-id="1d372-352">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-352">Name</span></span>| <span data-ttu-id="1d372-353">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-353">Type</span></span>| <span data-ttu-id="1d372-354">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d372-355">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-355">String</span></span>|<span data-ttu-id="1d372-356">Identificateur EWS (services web Exchange) pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="1d372-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-357">Requirements</span></span>

|<span data-ttu-id="1d372-358">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-358">Requirement</span></span>| <span data-ttu-id="1d372-359">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-360">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-361">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-361">1.0</span></span>|
|[<span data-ttu-id="1d372-362">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-363">ReadItem</span></span>|
|[<span data-ttu-id="1d372-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-365">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-366">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1d372-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1d372-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1d372-368">Affiche un formulaire pour créer un rendez-vous de calendrier.</span><span class="sxs-lookup"><span data-stu-id="1d372-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-369">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1d372-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d372-p114">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont fournis, les champs du formulaire de rendez-vous sont automatiquement remplis avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="1d372-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1d372-p115">Dans Outlook Web App et OWA for Devices, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire inclura les participants et un bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="1d372-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1d372-p116">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans les paramètres `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="1d372-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1d372-377">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="1d372-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-378">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-378">Parameters:</span></span>

|<span data-ttu-id="1d372-379">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-379">Name</span></span>| <span data-ttu-id="1d372-380">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-380">Type</span></span>| <span data-ttu-id="1d372-381">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1d372-382">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-382">Object</span></span> | <span data-ttu-id="1d372-383">Un dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1d372-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1d372-384">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d372-p117">Un tableau de chaînes contenant les adresses de messagerie ou un tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1d372-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1d372-387">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d372-p118">Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1d372-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1d372-390">Date</span><span class="sxs-lookup"><span data-stu-id="1d372-390">Date</span></span> | <span data-ttu-id="1d372-391">Un objet `Date` indiquant la date et l’heure du début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1d372-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1d372-392">Date</span><span class="sxs-lookup"><span data-stu-id="1d372-392">Date</span></span> | <span data-ttu-id="1d372-393">Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1d372-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1d372-394">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-394">String</span></span> | <span data-ttu-id="1d372-p119">Une chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1d372-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1d372-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1d372-p120">Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1d372-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1d372-400">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-400">String</span></span> | <span data-ttu-id="1d372-p121">Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1d372-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1d372-403">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-403">String</span></span> | <span data-ttu-id="1d372-p122">Le corps du rendez-vous. Le contenu du corps est limité à une taille maximale de 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1d372-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d372-406">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-406">Requirements</span></span>

|<span data-ttu-id="1d372-407">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-407">Requirement</span></span>| <span data-ttu-id="1d372-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-409">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-410">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-410">1.0</span></span>|
|[<span data-ttu-id="1d372-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-412">ReadItem</span></span>|
|[<span data-ttu-id="1d372-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-415">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="1d372-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1d372-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="1d372-417">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d372-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="1d372-p123">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="1d372-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-420">Les compléments doivent, dans la mesure du possible, utiliser les API REST plutôt que les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d372-420">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="1d372-421">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="1d372-421">**REST Tokens**</span></span>

<span data-ttu-id="1d372-p124">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut seulement accéder à l’élément actif et à ses pièces jointes en lecture seule, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="1d372-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="1d372-425">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="1d372-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="1d372-426">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="1d372-426">**EWS Tokens**</span></span>

<span data-ttu-id="1d372-p125">Lorsque un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut seulement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1d372-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="1d372-429">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="1d372-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-430">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-430">Parameters:</span></span>

|<span data-ttu-id="1d372-431">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-431">Name</span></span>| <span data-ttu-id="1d372-432">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-432">Type</span></span>| <span data-ttu-id="1d372-433">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d372-433">Attributes</span></span>| <span data-ttu-id="1d372-434">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="1d372-435">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-435">Object</span></span> | <span data-ttu-id="1d372-436">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-436">&lt;optional&gt;</span></span> | <span data-ttu-id="1d372-437">Littéral d'objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1d372-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="1d372-438">Booléen</span><span class="sxs-lookup"><span data-stu-id="1d372-438">Boolean</span></span> |  <span data-ttu-id="1d372-439">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-439">&lt;optional&gt;</span></span> | <span data-ttu-id="1d372-p126">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d372-442">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-442">Object</span></span> |  <span data-ttu-id="1d372-443">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-443">&lt;optional&gt;</span></span> | <span data-ttu-id="1d372-444">Toute donnée d'état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1d372-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="1d372-445">fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-445">function</span></span>||<span data-ttu-id="1d372-p127">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-448">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-448">Requirements</span></span>

|<span data-ttu-id="1d372-449">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-449">Requirement</span></span>| <span data-ttu-id="1d372-450">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-451">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-452">1.5</span><span class="sxs-lookup"><span data-stu-id="1d372-452">1.5</span></span> |
|[<span data-ttu-id="1d372-453">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-454">ReadItem</span></span>|
|[<span data-ttu-id="1d372-455">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-456">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-457">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1d372-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d372-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d372-459">Obtient une chaîne qui contient un jeton utilisé pour obtenir une pièce jointe ou un élément à partir d’un Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1d372-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1d372-p128">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="1d372-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1d372-p129">Vous pouvez transmettre le jeton et un identificateur de pièce jointe ou un identificateur d'élément à un système tiers. Le système tiers utilise le jeton comme jeton d'autorisation au porteur pour appeler l'opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services Web Exchange, pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1d372-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1d372-465">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste, pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="1d372-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="1d372-p130">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-468">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-468">Parameters:</span></span>

|<span data-ttu-id="1d372-469">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-469">Name</span></span>| <span data-ttu-id="1d372-470">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-470">Type</span></span>| <span data-ttu-id="1d372-471">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d372-471">Attributes</span></span>| <span data-ttu-id="1d372-472">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1d372-473">fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-473">function</span></span>||<span data-ttu-id="1d372-p131">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d372-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1d372-476">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-476">Object</span></span>| <span data-ttu-id="1d372-477">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-477">&lt;optional&gt;</span></span>|<span data-ttu-id="1d372-478">Toute donnée d'état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1d372-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-479">Requirements</span></span>

|<span data-ttu-id="1d372-480">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-480">Requirement</span></span>| <span data-ttu-id="1d372-481">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-482">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-483">1.3</span><span class="sxs-lookup"><span data-stu-id="1d372-483">1.3</span></span>|
|[<span data-ttu-id="1d372-484">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-485">ReadItem</span></span>|
|[<span data-ttu-id="1d372-486">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-487">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-488">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1d372-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d372-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d372-490">Obtient un jeton identifiant l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="1d372-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1d372-491">La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="1d372-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-492">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-492">Parameters:</span></span>

|<span data-ttu-id="1d372-493">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-493">Name</span></span>| <span data-ttu-id="1d372-494">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-494">Type</span></span>| <span data-ttu-id="1d372-495">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d372-495">Attributes</span></span>| <span data-ttu-id="1d372-496">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1d372-497">fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-497">function</span></span>||<span data-ttu-id="1d372-498">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d372-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d372-499">Le jeton est fourni sous la forme d'une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d372-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1d372-500">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-500">Object</span></span>| <span data-ttu-id="1d372-501">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-501">&lt;optional&gt;</span></span>|<span data-ttu-id="1d372-502">Toute donnée d'état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1d372-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-503">Requirements</span></span>

|<span data-ttu-id="1d372-504">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-504">Requirement</span></span>| <span data-ttu-id="1d372-505">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-506">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-507">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-507">1.0</span></span>|
|[<span data-ttu-id="1d372-508">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d372-509">ReadItem</span></span>|
|[<span data-ttu-id="1d372-510">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-511">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-512">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1d372-513">makeEwsRequestAsync (données, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d372-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1d372-514">Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1d372-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-515">Cette méthode n'est pas prise en charge dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="1d372-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1d372-516">Dans Outlook pour iOS ou Outlook pour Android</span><span class="sxs-lookup"><span data-stu-id="1d372-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="1d372-517">Lorsque le complément est chargé dans une boîte aux lettres Gmail</span><span class="sxs-lookup"><span data-stu-id="1d372-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1d372-518">Dans ces cas, les compléments doivent plutôt [utiliser des API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1d372-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1d372-519">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange, de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="1d372-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="1d372-520">Pour une liste des opérations EWS prises en charge, voir [Appeler des services web depuis un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="1d372-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1d372-521">Avec la méthode `makeEwsRequestAsync`, vous ne pouvez pas demander des éléments associés à un dossier.</span><span class="sxs-lookup"><span data-stu-id="1d372-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1d372-522">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="1d372-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1d372-p133">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la rubrique [Spécifier les autorisations d'accès de complément de messagerie à la boîte aux lettres de l'utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="1d372-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1d372-525">L’administrateur du serveur doit définir `OAuthAuthentication` sur true dans l'annuaire EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="1d372-525">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1d372-526">Différences entre versions</span><span class="sxs-lookup"><span data-stu-id="1d372-526">Version differences</span></span>

<span data-ttu-id="1d372-527">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook antérieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="1d372-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1d372-p134">Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="1d372-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d372-531">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1d372-531">Parameters:</span></span>

|<span data-ttu-id="1d372-532">Nom</span><span class="sxs-lookup"><span data-stu-id="1d372-532">Name</span></span>| <span data-ttu-id="1d372-533">Type</span><span class="sxs-lookup"><span data-stu-id="1d372-533">Type</span></span>| <span data-ttu-id="1d372-534">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d372-534">Attributes</span></span>| <span data-ttu-id="1d372-535">Description</span><span class="sxs-lookup"><span data-stu-id="1d372-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1d372-536">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1d372-536">String</span></span>||<span data-ttu-id="1d372-537">La demande EWS.</span><span class="sxs-lookup"><span data-stu-id="1d372-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1d372-538">fonction</span><span class="sxs-lookup"><span data-stu-id="1d372-538">function</span></span>||<span data-ttu-id="1d372-539">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d372-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d372-540">Le résultat XML de l’appel EWS est fourni comme une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d372-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="1d372-541">Si le résultat dépasse une taille de 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="1d372-541">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="1d372-542">Objet</span><span class="sxs-lookup"><span data-stu-id="1d372-542">Object</span></span>| <span data-ttu-id="1d372-543">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1d372-543">&lt;optional&gt;</span></span>|<span data-ttu-id="1d372-544">Toute donnée d'état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1d372-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d372-545">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1d372-545">Requirements</span></span>

|<span data-ttu-id="1d372-546">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1d372-546">Requirement</span></span>| <span data-ttu-id="1d372-547">Valeur</span><span class="sxs-lookup"><span data-stu-id="1d372-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d372-548">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1d372-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d372-549">1.0</span><span class="sxs-lookup"><span data-stu-id="1d372-549">1.0</span></span>|
|[<span data-ttu-id="1d372-550">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1d372-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d372-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1d372-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1d372-552">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1d372-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d372-553">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1d372-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d372-554">Exemple</span><span class="sxs-lookup"><span data-stu-id="1d372-554">Example</span></span>

<span data-ttu-id="1d372-555">L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` et obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="1d372-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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