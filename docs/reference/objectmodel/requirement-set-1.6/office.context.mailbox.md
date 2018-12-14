
# <a name="mailbox"></a><span data-ttu-id="c876d-101">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-101">mailbox</span></span>

### <span data-ttu-id="c876d-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="c876d-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="c876d-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="c876d-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c876d-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-105">Requirements</span></span>

|<span data-ttu-id="c876d-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-106">Requirement</span></span>| <span data-ttu-id="c876d-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-109">1.0</span></span>|
|[<span data-ttu-id="c876d-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c876d-111">Restricted</span></span>|
|[<span data-ttu-id="c876d-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c876d-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c876d-114">Members and methods</span></span>

| <span data-ttu-id="c876d-115">Membre</span><span class="sxs-lookup"><span data-stu-id="c876d-115">Member</span></span> | <span data-ttu-id="c876d-116">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c876d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c876d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c876d-118">Membre</span><span class="sxs-lookup"><span data-stu-id="c876d-118">Member</span></span> |
| [<span data-ttu-id="c876d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="c876d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c876d-120">Membre</span><span class="sxs-lookup"><span data-stu-id="c876d-120">Member</span></span> |
| [<span data-ttu-id="c876d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c876d-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-122">Method</span></span> |
| [<span data-ttu-id="c876d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c876d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c876d-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-124">Method</span></span> |
| [<span data-ttu-id="c876d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c876d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="c876d-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-126">Method</span></span> |
| [<span data-ttu-id="c876d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c876d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c876d-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-128">Method</span></span> |
| [<span data-ttu-id="c876d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c876d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c876d-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-130">Method</span></span> |
| [<span data-ttu-id="c876d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c876d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c876d-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-132">Method</span></span> |
| [<span data-ttu-id="c876d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c876d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c876d-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-134">Method</span></span> |
| [<span data-ttu-id="c876d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c876d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c876d-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-136">Method</span></span> |
| [<span data-ttu-id="c876d-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="c876d-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="c876d-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-138">Method</span></span> |
| [<span data-ttu-id="c876d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c876d-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-140">Method</span></span> |
| [<span data-ttu-id="c876d-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c876d-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-142">Method</span></span> |
| [<span data-ttu-id="c876d-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c876d-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-144">Method</span></span> |
| [<span data-ttu-id="c876d-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c876d-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-146">Method</span></span> |
| [<span data-ttu-id="c876d-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c876d-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c876d-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="c876d-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c876d-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c876d-149">Namespaces</span></span>

<span data-ttu-id="c876d-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c876d-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c876d-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c876d-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c876d-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c876d-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c876d-153">Membres</span><span class="sxs-lookup"><span data-stu-id="c876d-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c876d-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="c876d-154">ewsUrl :String</span></span>

<span data-ttu-id="c876d-p102">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c876d-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-157">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-p103">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c876d-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c876d-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="c876d-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c876d-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c876d-163">Type :</span><span class="sxs-lookup"><span data-stu-id="c876d-163">Type:</span></span>

*   <span data-ttu-id="c876d-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c876d-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c876d-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-165">Requirements</span></span>

|<span data-ttu-id="c876d-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-166">Requirement</span></span>| <span data-ttu-id="c876d-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-169">1.0</span></span>|
|[<span data-ttu-id="c876d-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-171">ReadItem</span></span>|
|[<span data-ttu-id="c876d-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="c876d-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="c876d-174">restUrl :String</span></span>

<span data-ttu-id="c876d-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="c876d-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c876d-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c876d-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="c876d-177">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="c876d-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="c876d-p105">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c876d-180">Type :</span><span class="sxs-lookup"><span data-stu-id="c876d-180">Type:</span></span>

*   <span data-ttu-id="c876d-181">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c876d-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c876d-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-182">Requirements</span></span>

|<span data-ttu-id="c876d-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-183">Requirement</span></span>| <span data-ttu-id="c876d-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-186">1,5</span><span class="sxs-lookup"><span data-stu-id="c876d-186">1.5</span></span> |
|[<span data-ttu-id="c876d-187">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-188">ReadItem</span></span>|
|[<span data-ttu-id="c876d-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-190">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="c876d-191">Méthodes</span><span class="sxs-lookup"><span data-stu-id="c876d-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c876d-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c876d-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c876d-193">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c876d-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c876d-194">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="c876d-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="c876d-195">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c876d-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-196">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-196">Parameters:</span></span>

| <span data-ttu-id="c876d-197">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-197">Name</span></span> | <span data-ttu-id="c876d-198">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-198">Type</span></span> | <span data-ttu-id="c876d-199">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-199">Attributes</span></span> | <span data-ttu-id="c876d-200">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c876d-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c876d-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c876d-202">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c876d-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c876d-203">Fonction</span><span class="sxs-lookup"><span data-stu-id="c876d-203">Function</span></span> || <span data-ttu-id="c876d-p107">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c876d-207">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-207">Object</span></span> | <span data-ttu-id="c876d-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-208">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-209">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c876d-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c876d-210">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-210">Object</span></span> | <span data-ttu-id="c876d-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-211">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-212">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c876d-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c876d-213">fonction</span><span class="sxs-lookup"><span data-stu-id="c876d-213">function</span></span>| <span data-ttu-id="c876d-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-214">&lt;optional&gt;</span></span>|<span data-ttu-id="c876d-215">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c876d-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-216">Requirements</span></span>

|<span data-ttu-id="c876d-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-217">Requirement</span></span>| <span data-ttu-id="c876d-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-220">1,5</span><span class="sxs-lookup"><span data-stu-id="c876d-220">1.5</span></span> |
|[<span data-ttu-id="c876d-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-222">ReadItem</span></span> |
|[<span data-ttu-id="c876d-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-224">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-224">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-225">Example</span></span>

```js
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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c876d-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c876d-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c876d-227">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="c876d-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-228">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-p108">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="c876d-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-231">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-231">Parameters:</span></span>

|<span data-ttu-id="c876d-232">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-232">Name</span></span>| <span data-ttu-id="c876d-233">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-233">Type</span></span>| <span data-ttu-id="c876d-234">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c876d-235">String</span><span class="sxs-lookup"><span data-stu-id="c876d-235">String</span></span>|<span data-ttu-id="c876d-236">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="c876d-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c876d-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c876d-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="c876d-238">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="c876d-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-239">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-239">Requirements</span></span>

|<span data-ttu-id="c876d-240">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-240">Requirement</span></span>| <span data-ttu-id="c876d-241">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-242">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-243">1.3</span><span class="sxs-lookup"><span data-stu-id="c876d-243">1.3</span></span>|
|[<span data-ttu-id="c876d-244">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-244">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-245">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c876d-245">Restricted</span></span>|
|[<span data-ttu-id="c876d-246">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-246">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-247">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-247">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c876d-248">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c876d-248">Returns:</span></span>

<span data-ttu-id="c876d-249">Type : String</span><span class="sxs-lookup"><span data-stu-id="c876d-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c876d-250">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-250">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="c876d-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="c876d-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="c876d-252">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="c876d-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c876d-p109">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c876d-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c876d-p110">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="c876d-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-258">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-258">Parameters:</span></span>

|<span data-ttu-id="c876d-259">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-259">Name</span></span>| <span data-ttu-id="c876d-260">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-260">Type</span></span>| <span data-ttu-id="c876d-261">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c876d-262">Date</span><span class="sxs-lookup"><span data-stu-id="c876d-262">Date</span></span>|<span data-ttu-id="c876d-263">Objet Date</span><span class="sxs-lookup"><span data-stu-id="c876d-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-264">Requirements</span></span>

|<span data-ttu-id="c876d-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-265">Requirement</span></span>| <span data-ttu-id="c876d-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-268">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-268">1.0</span></span>|
|[<span data-ttu-id="c876d-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-270">ReadItem</span></span>|
|[<span data-ttu-id="c876d-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-272">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-272">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c876d-273">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c876d-273">Returns:</span></span>

<span data-ttu-id="c876d-274">Type : [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="c876d-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c876d-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c876d-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c876d-276">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="c876d-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-277">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-p111">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="c876d-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-280">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-280">Parameters:</span></span>

|<span data-ttu-id="c876d-281">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-281">Name</span></span>| <span data-ttu-id="c876d-282">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-282">Type</span></span>| <span data-ttu-id="c876d-283">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c876d-284">String</span><span class="sxs-lookup"><span data-stu-id="c876d-284">String</span></span>|<span data-ttu-id="c876d-285">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="c876d-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c876d-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c876d-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="c876d-287">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="c876d-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-288">Requirements</span></span>

|<span data-ttu-id="c876d-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-289">Requirement</span></span>| <span data-ttu-id="c876d-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-292">1.3</span><span class="sxs-lookup"><span data-stu-id="c876d-292">1.3</span></span>|
|[<span data-ttu-id="c876d-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-294">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c876d-294">Restricted</span></span>|
|[<span data-ttu-id="c876d-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-296">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-296">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c876d-297">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c876d-297">Returns:</span></span>

<span data-ttu-id="c876d-298">Type : String</span><span class="sxs-lookup"><span data-stu-id="c876d-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c876d-299">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-299">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c876d-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c876d-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c876d-301">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="c876d-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c876d-302">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="c876d-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-303">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-303">Parameters:</span></span>

|<span data-ttu-id="c876d-304">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-304">Name</span></span>| <span data-ttu-id="c876d-305">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-305">Type</span></span>| <span data-ttu-id="c876d-306">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c876d-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c876d-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="c876d-308">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="c876d-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-309">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-309">Requirements</span></span>

|<span data-ttu-id="c876d-310">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-310">Requirement</span></span>| <span data-ttu-id="c876d-311">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-312">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-313">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-313">1.0</span></span>|
|[<span data-ttu-id="c876d-314">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-315">ReadItem</span></span>|
|[<span data-ttu-id="c876d-316">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-317">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-317">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c876d-318">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c876d-318">Returns:</span></span>

<span data-ttu-id="c876d-319">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="c876d-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="c876d-320">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="c876d-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c876d-321">Date</span><span class="sxs-lookup"><span data-stu-id="c876d-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="c876d-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c876d-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c876d-323">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="c876d-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-324">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-325">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="c876d-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c876d-p112">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="c876d-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c876d-328">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c876d-329">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="c876d-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-330">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-330">Parameters:</span></span>

|<span data-ttu-id="c876d-331">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-331">Name</span></span>| <span data-ttu-id="c876d-332">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-332">Type</span></span>| <span data-ttu-id="c876d-333">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c876d-334">String</span><span class="sxs-lookup"><span data-stu-id="c876d-334">String</span></span>|<span data-ttu-id="c876d-335">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="c876d-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-336">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-336">Requirements</span></span>

|<span data-ttu-id="c876d-337">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-337">Requirement</span></span>| <span data-ttu-id="c876d-338">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-339">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-340">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-340">1.0</span></span>|
|[<span data-ttu-id="c876d-341">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-341">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-342">ReadItem</span></span>|
|[<span data-ttu-id="c876d-343">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-343">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-344">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-344">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-345">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="c876d-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c876d-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c876d-347">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="c876d-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-348">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-349">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="c876d-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c876d-350">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c876d-351">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="c876d-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c876d-p113">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c876d-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-354">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-354">Parameters:</span></span>

|<span data-ttu-id="c876d-355">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-355">Name</span></span>| <span data-ttu-id="c876d-356">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-356">Type</span></span>| <span data-ttu-id="c876d-357">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c876d-358">String</span><span class="sxs-lookup"><span data-stu-id="c876d-358">String</span></span>|<span data-ttu-id="c876d-359">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="c876d-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-360">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-360">Requirements</span></span>

|<span data-ttu-id="c876d-361">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-361">Requirement</span></span>| <span data-ttu-id="c876d-362">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-364">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-364">1.0</span></span>|
|[<span data-ttu-id="c876d-365">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-366">ReadItem</span></span>|
|[<span data-ttu-id="c876d-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-368">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-368">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-369">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c876d-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c876d-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c876d-371">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="c876d-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-372">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c876d-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c876d-p114">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="c876d-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c876d-p115">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="c876d-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c876d-p116">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="c876d-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c876d-380">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="c876d-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-381">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-381">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-382">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="c876d-382">All parameters are optional.</span></span>

|<span data-ttu-id="c876d-383">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-383">Name</span></span>| <span data-ttu-id="c876d-384">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-384">Type</span></span>| <span data-ttu-id="c876d-385">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c876d-386">Object</span><span class="sxs-lookup"><span data-stu-id="c876d-386">Object</span></span> | <span data-ttu-id="c876d-387">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c876d-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c876d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c876d-p117">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c876d-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c876d-p118">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c876d-394">Date</span><span class="sxs-lookup"><span data-stu-id="c876d-394">Date</span></span> | <span data-ttu-id="c876d-395">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c876d-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c876d-396">Date</span><span class="sxs-lookup"><span data-stu-id="c876d-396">Date</span></span> | <span data-ttu-id="c876d-397">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c876d-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c876d-398">String</span><span class="sxs-lookup"><span data-stu-id="c876d-398">String</span></span> | <span data-ttu-id="c876d-p119">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c876d-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c876d-p120">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c876d-404">String</span><span class="sxs-lookup"><span data-stu-id="c876d-404">String</span></span> | <span data-ttu-id="c876d-p121">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c876d-407">String</span><span class="sxs-lookup"><span data-stu-id="c876d-407">String</span></span> | <span data-ttu-id="c876d-p122">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c876d-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c876d-410">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-410">Requirements</span></span>

|<span data-ttu-id="c876d-411">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-411">Requirement</span></span>| <span data-ttu-id="c876d-412">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-413">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-414">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-414">1.0</span></span>|
|[<span data-ttu-id="c876d-415">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-415">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-416">ReadItem</span></span>|
|[<span data-ttu-id="c876d-417">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-417">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-418">Lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-419">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="c876d-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c876d-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="c876d-421">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="c876d-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="c876d-422">La méthode `displayNewMessageForm` ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="c876d-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="c876d-423">Si des paramètres sont spécifiés, les champs du formulaire de message sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="c876d-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c876d-424">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="c876d-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-425">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-425">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-426">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="c876d-426">All parameters are optional.</span></span>

|<span data-ttu-id="c876d-427">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-427">Name</span></span>| <span data-ttu-id="c876d-428">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-428">Type</span></span>| <span data-ttu-id="c876d-429">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c876d-430">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-430">Object</span></span> | <span data-ttu-id="c876d-431">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="c876d-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="c876d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c876d-433">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne À.</span><span class="sxs-lookup"><span data-stu-id="c876d-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="c876d-434">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="c876d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c876d-436">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cc.</span><span class="sxs-lookup"><span data-stu-id="c876d-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="c876d-437">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="c876d-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c876d-439">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cci.</span><span class="sxs-lookup"><span data-stu-id="c876d-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="c876d-440">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c876d-441">String</span><span class="sxs-lookup"><span data-stu-id="c876d-441">String</span></span> | <span data-ttu-id="c876d-442">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="c876d-442">A string containing the subject of the message.</span></span> <span data-ttu-id="c876d-443">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="c876d-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="c876d-444">String</span><span class="sxs-lookup"><span data-stu-id="c876d-444">String</span></span> | <span data-ttu-id="c876d-445">Corps du message HTML.</span><span class="sxs-lookup"><span data-stu-id="c876d-445">The HTML body of the message.</span></span> <span data-ttu-id="c876d-446">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c876d-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="c876d-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c876d-448">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="c876d-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="c876d-449">String</span><span class="sxs-lookup"><span data-stu-id="c876d-449">String</span></span> | <span data-ttu-id="c876d-p129">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="c876d-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="c876d-452">String</span><span class="sxs-lookup"><span data-stu-id="c876d-452">String</span></span> | <span data-ttu-id="c876d-453">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c876d-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="c876d-454">String</span><span class="sxs-lookup"><span data-stu-id="c876d-454">String</span></span> | <span data-ttu-id="c876d-p130">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="c876d-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="c876d-457">Boolean</span><span class="sxs-lookup"><span data-stu-id="c876d-457">Boolean</span></span> | <span data-ttu-id="c876d-p131">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c876d-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="c876d-460">String</span><span class="sxs-lookup"><span data-stu-id="c876d-460">String</span></span> | <span data-ttu-id="c876d-461">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="c876d-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="c876d-462">ID d’élément EWS du courrier électronique existant à joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="c876d-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="c876d-463">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c876d-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="c876d-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-464">Requirements</span></span>

|<span data-ttu-id="c876d-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-465">Requirement</span></span>| <span data-ttu-id="c876d-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-468">1.6</span><span class="sxs-lookup"><span data-stu-id="c876d-468">1.6</span></span> |
|[<span data-ttu-id="c876d-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-470">ReadItem</span></span>|
|[<span data-ttu-id="c876d-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-473">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c876d-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c876d-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c876d-475">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="c876d-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c876d-p133">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="c876d-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-478">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="c876d-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="c876d-479">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="c876d-479">**REST Tokens**</span></span>

<span data-ttu-id="c876d-p134">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="c876d-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c876d-483">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="c876d-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c876d-484">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="c876d-484">**EWS Tokens**</span></span>

<span data-ttu-id="c876d-p135">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c876d-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c876d-487">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="c876d-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-488">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-488">Parameters:</span></span>

|<span data-ttu-id="c876d-489">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-489">Name</span></span>| <span data-ttu-id="c876d-490">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-490">Type</span></span>| <span data-ttu-id="c876d-491">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-491">Attributes</span></span>| <span data-ttu-id="c876d-492">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c876d-493">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-493">Object</span></span> | <span data-ttu-id="c876d-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-494">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-495">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c876d-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c876d-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="c876d-496">Boolean</span></span> |  <span data-ttu-id="c876d-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-497">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-p136">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c876d-500">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-500">Object</span></span> |  <span data-ttu-id="c876d-501">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-501">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-502">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c876d-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c876d-503">fonction</span><span class="sxs-lookup"><span data-stu-id="c876d-503">function</span></span>||<span data-ttu-id="c876d-p137">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-506">Requirements</span></span>

|<span data-ttu-id="c876d-507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-507">Requirement</span></span>| <span data-ttu-id="c876d-508">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-510">1,5</span><span class="sxs-lookup"><span data-stu-id="c876d-510">1.5</span></span> |
|[<span data-ttu-id="c876d-511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-512">ReadItem</span></span>|
|[<span data-ttu-id="c876d-513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-514">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c876d-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c876d-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c876d-517">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="c876d-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c876d-p138">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="c876d-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c876d-p139">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c876d-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c876d-523">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="c876d-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="c876d-p140">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-526">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-526">Parameters:</span></span>

|<span data-ttu-id="c876d-527">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-527">Name</span></span>| <span data-ttu-id="c876d-528">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-528">Type</span></span>| <span data-ttu-id="c876d-529">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-529">Attributes</span></span>| <span data-ttu-id="c876d-530">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c876d-531">function</span><span class="sxs-lookup"><span data-stu-id="c876d-531">function</span></span>||<span data-ttu-id="c876d-p141">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c876d-534">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-534">Object</span></span>| <span data-ttu-id="c876d-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-535">&lt;optional&gt;</span></span>|<span data-ttu-id="c876d-536">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c876d-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-537">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-537">Requirements</span></span>

|<span data-ttu-id="c876d-538">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-538">Requirement</span></span>| <span data-ttu-id="c876d-539">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-540">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-541">1.3</span><span class="sxs-lookup"><span data-stu-id="c876d-541">1.3</span></span>|
|[<span data-ttu-id="c876d-542">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-543">ReadItem</span></span>|
|[<span data-ttu-id="c876d-544">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-545">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-546">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-546">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c876d-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c876d-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c876d-548">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="c876d-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c876d-549">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="c876d-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-550">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-550">Parameters:</span></span>

|<span data-ttu-id="c876d-551">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-551">Name</span></span>| <span data-ttu-id="c876d-552">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-552">Type</span></span>| <span data-ttu-id="c876d-553">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-553">Attributes</span></span>| <span data-ttu-id="c876d-554">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c876d-555">function</span><span class="sxs-lookup"><span data-stu-id="c876d-555">function</span></span>||<span data-ttu-id="c876d-556">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c876d-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c876d-557">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c876d-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c876d-558">Object</span><span class="sxs-lookup"><span data-stu-id="c876d-558">Object</span></span>| <span data-ttu-id="c876d-559">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-559">&lt;optional&gt;</span></span>|<span data-ttu-id="c876d-560">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c876d-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-561">Requirements</span></span>

|<span data-ttu-id="c876d-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-562">Requirement</span></span>| <span data-ttu-id="c876d-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-565">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-565">1.0</span></span>|
|[<span data-ttu-id="c876d-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-567">ReadItem</span></span>|
|[<span data-ttu-id="c876d-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-569">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-570">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c876d-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c876d-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c876d-572">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c876d-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-573">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="c876d-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c876d-574">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="c876d-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="c876d-575">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="c876d-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c876d-576">Dans ces cas de figure, les compléments doivent [utiliser les API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c876d-576">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c876d-577">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="c876d-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="c876d-578">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="c876d-578">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c876d-579">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c876d-580">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="c876d-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c876d-p143">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="c876d-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c876d-583">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="c876d-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c876d-584">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="c876d-584">Version differences</span></span>

<span data-ttu-id="c876d-585">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="c876d-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c876d-p144">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="c876d-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-589">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-589">Parameters:</span></span>

|<span data-ttu-id="c876d-590">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-590">Name</span></span>| <span data-ttu-id="c876d-591">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-591">Type</span></span>| <span data-ttu-id="c876d-592">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-592">Attributes</span></span>| <span data-ttu-id="c876d-593">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c876d-594">String</span><span class="sxs-lookup"><span data-stu-id="c876d-594">String</span></span>||<span data-ttu-id="c876d-595">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="c876d-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c876d-596">function</span><span class="sxs-lookup"><span data-stu-id="c876d-596">function</span></span>||<span data-ttu-id="c876d-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c876d-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c876d-598">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c876d-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="c876d-599">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="c876d-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c876d-600">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-600">Object</span></span>| <span data-ttu-id="c876d-601">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-601">&lt;optional&gt;</span></span>|<span data-ttu-id="c876d-602">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c876d-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-603">Requirements</span></span>

|<span data-ttu-id="c876d-604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-604">Requirement</span></span>| <span data-ttu-id="c876d-605">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-607">1.0</span><span class="sxs-lookup"><span data-stu-id="c876d-607">1.0</span></span>|
|[<span data-ttu-id="c876d-608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c876d-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c876d-610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-611">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c876d-612">Exemple</span><span class="sxs-lookup"><span data-stu-id="c876d-612">Example</span></span>

<span data-ttu-id="c876d-613">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="c876d-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c876d-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c876d-614">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c876d-615">Retire un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c876d-615">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="c876d-616">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c876d-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c876d-617">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c876d-617">Parameters:</span></span>

| <span data-ttu-id="c876d-618">Nom</span><span class="sxs-lookup"><span data-stu-id="c876d-618">Name</span></span> | <span data-ttu-id="c876d-619">Type</span><span class="sxs-lookup"><span data-stu-id="c876d-619">Type</span></span> | <span data-ttu-id="c876d-620">Attributs</span><span class="sxs-lookup"><span data-stu-id="c876d-620">Attributes</span></span> | <span data-ttu-id="c876d-621">Description</span><span class="sxs-lookup"><span data-stu-id="c876d-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c876d-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c876d-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c876d-623">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c876d-623">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c876d-624">Fonction</span><span class="sxs-lookup"><span data-stu-id="c876d-624">Function</span></span> || <span data-ttu-id="c876d-p146">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c876d-p146">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c876d-628">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-628">Object</span></span> | <span data-ttu-id="c876d-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-629">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-630">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c876d-630">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c876d-631">Objet</span><span class="sxs-lookup"><span data-stu-id="c876d-631">Object</span></span> | <span data-ttu-id="c876d-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-632">&lt;optional&gt;</span></span> | <span data-ttu-id="c876d-633">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c876d-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c876d-634">fonction</span><span class="sxs-lookup"><span data-stu-id="c876d-634">function</span></span>| <span data-ttu-id="c876d-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c876d-635">&lt;optional&gt;</span></span>|<span data-ttu-id="c876d-636">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c876d-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c876d-637">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c876d-637">Requirements</span></span>

|<span data-ttu-id="c876d-638">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c876d-638">Requirement</span></span>| <span data-ttu-id="c876d-639">Valeur</span><span class="sxs-lookup"><span data-stu-id="c876d-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="c876d-640">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c876d-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c876d-641">1,5</span><span class="sxs-lookup"><span data-stu-id="c876d-641">1.5</span></span> |
|[<span data-ttu-id="c876d-642">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c876d-642">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c876d-643">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c876d-643">ReadItem</span></span> |
|[<span data-ttu-id="c876d-644">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c876d-644">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c876d-645">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c876d-645">Compose or read</span></span>|