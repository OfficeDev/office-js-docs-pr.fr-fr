
# <a name="mailbox"></a><span data-ttu-id="1e77b-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="1e77b-101">mailbox</span></span>

### <span data-ttu-id="1e77b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="1e77b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="1e77b-104">Donne accès au modèle d’objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="1e77b-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e77b-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-105">Requirements</span></span>

|<span data-ttu-id="1e77b-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-106">Requirement</span></span>| <span data-ttu-id="1e77b-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-108">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-109">1.0</span></span>|
|[<span data-ttu-id="1e77b-110">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="1e77b-111">Restricted</span></span>|
|[<span data-ttu-id="1e77b-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1e77b-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="1e77b-114">Members and methods</span></span>

| <span data-ttu-id="1e77b-115">Membre</span><span class="sxs-lookup"><span data-stu-id="1e77b-115">Member</span></span> | <span data-ttu-id="1e77b-116">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1e77b-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="1e77b-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="1e77b-118">Membre</span><span class="sxs-lookup"><span data-stu-id="1e77b-118">Member</span></span> |
| [<span data-ttu-id="1e77b-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="1e77b-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="1e77b-120">Membre</span><span class="sxs-lookup"><span data-stu-id="1e77b-120">Member</span></span> |
| [<span data-ttu-id="1e77b-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1e77b-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1e77b-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-122">Method</span></span> |
| [<span data-ttu-id="1e77b-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="1e77b-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="1e77b-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-124">Method</span></span> |
| [<span data-ttu-id="1e77b-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1e77b-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="1e77b-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-126">Method</span></span> |
| [<span data-ttu-id="1e77b-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="1e77b-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="1e77b-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-128">Method</span></span> |
| [<span data-ttu-id="1e77b-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="1e77b-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="1e77b-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-130">Method</span></span> |
| [<span data-ttu-id="1e77b-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1e77b-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="1e77b-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-132">Method</span></span> |
| [<span data-ttu-id="1e77b-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="1e77b-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="1e77b-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-134">Method</span></span> |
| [<span data-ttu-id="1e77b-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1e77b-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="1e77b-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-136">Method</span></span> |
| [<span data-ttu-id="1e77b-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="1e77b-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="1e77b-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-138">Method</span></span> |
| [<span data-ttu-id="1e77b-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1e77b-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="1e77b-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-140">Method</span></span> |
| [<span data-ttu-id="1e77b-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1e77b-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="1e77b-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-142">Method</span></span> |
| [<span data-ttu-id="1e77b-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1e77b-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="1e77b-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-144">Method</span></span> |
| [<span data-ttu-id="1e77b-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="1e77b-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="1e77b-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="1e77b-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1e77b-147">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="1e77b-147">Namespaces</span></span>

<span data-ttu-id="1e77b-148">[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e77b-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1e77b-149">[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e77b-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1e77b-150">[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e77b-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1e77b-151">Membres</span><span class="sxs-lookup"><span data-stu-id="1e77b-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1e77b-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="1e77b-152">ewsUrl :String</span></span>

<span data-ttu-id="1e77b-p102">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-155">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-p103">La valeur `ewsUrl` peut être utilisée par un service distant pour effectuer des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir les pièces jointes de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1e77b-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1e77b-158">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `ewsUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="1e77b-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="1e77b-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1e77b-161">Type :</span><span class="sxs-lookup"><span data-stu-id="1e77b-161">Type:</span></span>

*   <span data-ttu-id="1e77b-162">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e77b-163">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-163">Requirements</span></span>

|<span data-ttu-id="1e77b-164">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-164">Requirement</span></span>| <span data-ttu-id="1e77b-165">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-166">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-167">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-167">1.0</span></span>|
|[<span data-ttu-id="1e77b-168">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-169">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="1e77b-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="1e77b-172">restUrl :String</span></span>

<span data-ttu-id="1e77b-173">Obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="1e77b-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="1e77b-174">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1e77b-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="1e77b-175">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="1e77b-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="1e77b-p105">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1e77b-178">Type :</span><span class="sxs-lookup"><span data-stu-id="1e77b-178">Type:</span></span>

*   <span data-ttu-id="1e77b-179">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e77b-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-180">Requirements</span></span>

|<span data-ttu-id="1e77b-181">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-181">Requirement</span></span>| <span data-ttu-id="1e77b-182">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-183">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-184">1.5</span><span class="sxs-lookup"><span data-stu-id="1e77b-184">1.5</span></span> |
|[<span data-ttu-id="1e77b-185">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-186">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-187">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-188">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="1e77b-189">Méthodes</span><span class="sxs-lookup"><span data-stu-id="1e77b-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1e77b-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e77b-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1e77b-191">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="1e77b-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1e77b-p106">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé lorsque l’utilisateur sélectionne un nouvel élément. Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’interface utilisateur du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-194">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-194">Parameters:</span></span>

| <span data-ttu-id="1e77b-195">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-195">Name</span></span> | <span data-ttu-id="1e77b-196">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-196">Type</span></span> | <span data-ttu-id="1e77b-197">Attributs</span><span class="sxs-lookup"><span data-stu-id="1e77b-197">Attributes</span></span> | <span data-ttu-id="1e77b-198">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1e77b-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1e77b-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1e77b-200">L’événement qui doit invoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="1e77b-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1e77b-201">Fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-201">Function</span></span> || <span data-ttu-id="1e77b-p107">La fonction permettant de gérer l’événement. Cette fonction doit accepter un seul paramètre, qui est un d’objet textuel. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1e77b-205">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-205">Object</span></span> | <span data-ttu-id="1e77b-206">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-206">&lt;optional&gt;</span></span> | <span data-ttu-id="1e77b-207">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1e77b-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1e77b-208">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-208">Object</span></span> | <span data-ttu-id="1e77b-209">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-209">&lt;optional&gt;</span></span> | <span data-ttu-id="1e77b-210">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="1e77b-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1e77b-211">fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-211">function</span></span>| <span data-ttu-id="1e77b-212">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-212">&lt;optional&gt;</span></span>|<span data-ttu-id="1e77b-213">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1e77b-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-214">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-214">Requirements</span></span>

|<span data-ttu-id="1e77b-215">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-215">Requirement</span></span>| <span data-ttu-id="1e77b-216">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-217">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-218">1.5</span><span class="sxs-lookup"><span data-stu-id="1e77b-218">1.5</span></span> |
|[<span data-ttu-id="1e77b-219">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-220">ReadItem</span></span> |
|[<span data-ttu-id="1e77b-221">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-222">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-223">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-223">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="1e77b-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1e77b-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1e77b-225">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="1e77b-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-226">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-226">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-p108">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)) utilisent un format différent de celui employé par les services Web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-229">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-229">Parameters:</span></span>

|<span data-ttu-id="1e77b-230">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-230">Name</span></span>| <span data-ttu-id="1e77b-231">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-231">Type</span></span>| <span data-ttu-id="1e77b-232">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1e77b-233">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-233">String</span></span>|<span data-ttu-id="1e77b-234">Un ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="1e77b-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="1e77b-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1e77b-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="1e77b-236">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="1e77b-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-237">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-237">Requirements</span></span>

|<span data-ttu-id="1e77b-238">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-238">Requirement</span></span>| <span data-ttu-id="1e77b-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-240">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-241">1.3</span><span class="sxs-lookup"><span data-stu-id="1e77b-241">1.3</span></span>|
|[<span data-ttu-id="1e77b-242">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-243">Restreint</span><span class="sxs-lookup"><span data-stu-id="1e77b-243">Restricted</span></span>|
|[<span data-ttu-id="1e77b-244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-245">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e77b-246">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1e77b-246">Returns:</span></span>

<span data-ttu-id="1e77b-247">Type : String</span><span class="sxs-lookup"><span data-stu-id="1e77b-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1e77b-248">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="1e77b-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="1e77b-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="1e77b-250">Obtient un dictionnaire contenant des informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="1e77b-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1e77b-p109">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur client, Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire auquel l’utilisateur s'attend.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1e77b-p110">Si l’application de messagerie s’exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-256">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-256">Parameters:</span></span>

|<span data-ttu-id="1e77b-257">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-257">Name</span></span>| <span data-ttu-id="1e77b-258">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-258">Type</span></span>| <span data-ttu-id="1e77b-259">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1e77b-260">Date</span><span class="sxs-lookup"><span data-stu-id="1e77b-260">Date</span></span>|<span data-ttu-id="1e77b-261">Un objet Date</span><span class="sxs-lookup"><span data-stu-id="1e77b-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-262">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-262">Requirements</span></span>

|<span data-ttu-id="1e77b-263">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-263">Requirement</span></span>| <span data-ttu-id="1e77b-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-265">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-266">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-266">1.0</span></span>|
|[<span data-ttu-id="1e77b-267">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-268">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-270">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e77b-271">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1e77b-271">Returns:</span></span>

<span data-ttu-id="1e77b-272">Type : [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="1e77b-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="1e77b-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1e77b-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1e77b-274">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="1e77b-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-275">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-275">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-p111">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS dans un format adapté à REST.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-278">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-278">Parameters:</span></span>

|<span data-ttu-id="1e77b-279">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-279">Name</span></span>| <span data-ttu-id="1e77b-280">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-280">Type</span></span>| <span data-ttu-id="1e77b-281">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1e77b-282">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-282">String</span></span>|<span data-ttu-id="1e77b-283">Un ID d’élément mis en forme pour les services Web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="1e77b-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="1e77b-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1e77b-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="1e77b-285">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="1e77b-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-286">Requirements</span></span>

|<span data-ttu-id="1e77b-287">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-287">Requirement</span></span>| <span data-ttu-id="1e77b-288">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-289">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-290">1.3</span><span class="sxs-lookup"><span data-stu-id="1e77b-290">1.3</span></span>|
|[<span data-ttu-id="1e77b-291">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-292">Restreint</span><span class="sxs-lookup"><span data-stu-id="1e77b-292">Restricted</span></span>|
|[<span data-ttu-id="1e77b-293">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-294">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e77b-295">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1e77b-295">Returns:</span></span>

<span data-ttu-id="1e77b-296">Type : String</span><span class="sxs-lookup"><span data-stu-id="1e77b-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1e77b-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1e77b-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1e77b-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1e77b-299">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="1e77b-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1e77b-300">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="1e77b-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-301">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-301">Parameters:</span></span>

|<span data-ttu-id="1e77b-302">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-302">Name</span></span>| <span data-ttu-id="1e77b-303">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-303">Type</span></span>| <span data-ttu-id="1e77b-304">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1e77b-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1e77b-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="1e77b-306">Valeur en heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="1e77b-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-307">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-307">Requirements</span></span>

|<span data-ttu-id="1e77b-308">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-308">Requirement</span></span>| <span data-ttu-id="1e77b-309">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-310">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-311">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-311">1.0</span></span>|
|[<span data-ttu-id="1e77b-312">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-313">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-314">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-315">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e77b-316">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="1e77b-316">Returns:</span></span>

<span data-ttu-id="1e77b-317">Un objet Date avec l’heure exprimée en UTC.</span><span class="sxs-lookup"><span data-stu-id="1e77b-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="1e77b-318">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1e77b-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1e77b-319">Date</span><span class="sxs-lookup"><span data-stu-id="1e77b-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="1e77b-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1e77b-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1e77b-321">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="1e77b-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-322">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-322">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-323">La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les équipements mobiles.</span><span class="sxs-lookup"><span data-stu-id="1e77b-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1e77b-p112">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1e77b-326">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1e77b-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1e77b-327">Si l’identificateur d’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur n'est retourné.</span><span class="sxs-lookup"><span data-stu-id="1e77b-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-328">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-328">Parameters:</span></span>

|<span data-ttu-id="1e77b-329">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-329">Name</span></span>| <span data-ttu-id="1e77b-330">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-330">Type</span></span>| <span data-ttu-id="1e77b-331">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1e77b-332">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-332">String</span></span>|<span data-ttu-id="1e77b-333">L’identificateur EWS (services Web Exchange) pour un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="1e77b-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-334">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-334">Requirements</span></span>

|<span data-ttu-id="1e77b-335">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-335">Requirement</span></span>| <span data-ttu-id="1e77b-336">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-337">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-338">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-338">1.0</span></span>|
|[<span data-ttu-id="1e77b-339">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-340">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-341">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-342">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-343">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="1e77b-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1e77b-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1e77b-345">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="1e77b-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-346">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-346">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-347">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre sur le bureau, ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="1e77b-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1e77b-348">Dans Outlook Web App, cette méthode n’ouvre le formulaire indiqué que si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1e77b-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1e77b-349">Si l’identificateur d’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.</span><span class="sxs-lookup"><span data-stu-id="1e77b-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1e77b-p113">N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-352">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-352">Parameters:</span></span>

|<span data-ttu-id="1e77b-353">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-353">Name</span></span>| <span data-ttu-id="1e77b-354">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-354">Type</span></span>| <span data-ttu-id="1e77b-355">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1e77b-356">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-356">String</span></span>|<span data-ttu-id="1e77b-357">Identificateur EWS (services Web Exchange) pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="1e77b-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-358">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-358">Requirements</span></span>

|<span data-ttu-id="1e77b-359">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-359">Requirement</span></span>| <span data-ttu-id="1e77b-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-361">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-362">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-362">1.0</span></span>|
|[<span data-ttu-id="1e77b-363">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-364">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-366">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1e77b-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1e77b-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1e77b-369">Affiche un formulaire pour créer un rendez-vous de calendrier.</span><span class="sxs-lookup"><span data-stu-id="1e77b-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-370">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="1e77b-370">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e77b-p114">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont fournis, les champs du formulaire de rendez-vous sont automatiquement remplis avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1e77b-p115">Dans Outlook Web App et OWA for Devices, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire inclura les participants et un bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1e77b-p116">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans les paramètres `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1e77b-378">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="1e77b-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-379">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-380">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="1e77b-380">Note: All parameters are optional.</span></span>

|<span data-ttu-id="1e77b-381">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-381">Name</span></span>| <span data-ttu-id="1e77b-382">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-382">Type</span></span>| <span data-ttu-id="1e77b-383">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1e77b-384">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-384">Object</span></span> | <span data-ttu-id="1e77b-385">Un dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1e77b-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1e77b-386">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1e77b-p117">Un tableau de chaînes contenant les adresses de messagerie ou un tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1e77b-389">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1e77b-p118">Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1e77b-392">Date</span><span class="sxs-lookup"><span data-stu-id="1e77b-392">Date</span></span> | <span data-ttu-id="1e77b-393">Un objet `Date` indiquant la date et l’heure du début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1e77b-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1e77b-394">Date</span><span class="sxs-lookup"><span data-stu-id="1e77b-394">Date</span></span> | <span data-ttu-id="1e77b-395">Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1e77b-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1e77b-396">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-396">String</span></span> | <span data-ttu-id="1e77b-p119">Une chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1e77b-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1e77b-p120">Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1e77b-402">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-402">String</span></span> | <span data-ttu-id="1e77b-p121">Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1e77b-405">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-405">String</span></span> | <span data-ttu-id="1e77b-p122">Le corps du rendez-vous. Le contenu du corps est limité à une taille maximale de 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1e77b-408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-408">Requirements</span></span>

|<span data-ttu-id="1e77b-409">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-409">Requirement</span></span>| <span data-ttu-id="1e77b-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-411">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-412">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-412">1.0</span></span>|
|[<span data-ttu-id="1e77b-413">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-414">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-417">Example</span></span>

```
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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="1e77b-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1e77b-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="1e77b-419">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="1e77b-420">La méthode `displayNewMessageForm` ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="1e77b-421">Si des paramètres sont spécifiés, les champs du formulaire de message sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="1e77b-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1e77b-422">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="1e77b-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-423">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-424">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="1e77b-424">Note: All parameters are optional.</span></span>

|<span data-ttu-id="1e77b-425">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-425">Name</span></span>| <span data-ttu-id="1e77b-426">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-426">Type</span></span>| <span data-ttu-id="1e77b-427">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1e77b-428">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-428">Object</span></span> | <span data-ttu-id="1e77b-429">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="1e77b-430">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1e77b-431">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne À.</span><span class="sxs-lookup"><span data-stu-id="1e77b-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="1e77b-432">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="1e77b-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="1e77b-433">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1e77b-434">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cc.</span><span class="sxs-lookup"><span data-stu-id="1e77b-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="1e77b-435">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="1e77b-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="1e77b-436">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1e77b-437">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cci.</span><span class="sxs-lookup"><span data-stu-id="1e77b-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="1e77b-438">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="1e77b-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1e77b-439">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-439">String</span></span> | <span data-ttu-id="1e77b-440">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-440">A string containing the subject of the message.</span></span> <span data-ttu-id="1e77b-441">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="1e77b-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="1e77b-442">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-442">String</span></span> | <span data-ttu-id="1e77b-443">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-443">The HTML body of the message.</span></span> <span data-ttu-id="1e77b-444">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="1e77b-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="1e77b-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1e77b-446">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="1e77b-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="1e77b-447">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-447">String</span></span> | <span data-ttu-id="1e77b-p129">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="1e77b-450">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-450">String</span></span> | <span data-ttu-id="1e77b-451">Chaîne qui contient le nom de la pièce jointe, d'une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="1e77b-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="1e77b-452">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-452">String</span></span> | <span data-ttu-id="1e77b-p130">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="1e77b-455">Boolean</span><span class="sxs-lookup"><span data-stu-id="1e77b-455">Boolean</span></span> | <span data-ttu-id="1e77b-p131">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="1e77b-458">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-458">String</span></span> | <span data-ttu-id="1e77b-459">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="1e77b-460">L’id d’élément EWS du courrier électronique existant à joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="1e77b-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="1e77b-461">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="1e77b-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="1e77b-462">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-462">Requirements</span></span>

|<span data-ttu-id="1e77b-463">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-463">Requirement</span></span>| <span data-ttu-id="1e77b-464">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-465">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-466">1.6</span><span class="sxs-lookup"><span data-stu-id="1e77b-466">-16</span></span> |
|[<span data-ttu-id="1e77b-467">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-468">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-469">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-470">Lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-471">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-471">Example</span></span>

```
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="1e77b-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1e77b-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="1e77b-473">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1e77b-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="1e77b-p133">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-476">Les compléments doivent, dans la mesure du possible, utiliser les API REST plutôt que les services Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1e77b-476">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="1e77b-477">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="1e77b-477">**REST Tokens**</span></span>

<span data-ttu-id="1e77b-p134">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services Web Exchange. Le jeton peut seulement accéder à l’élément actif et à ses pièces jointes en lecture seule, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="1e77b-481">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="1e77b-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="1e77b-482">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="1e77b-482">**EWS Tokens**</span></span>

<span data-ttu-id="1e77b-p135">Lorsque un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut seulement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="1e77b-485">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="1e77b-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-486">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-486">Parameters:</span></span>

|<span data-ttu-id="1e77b-487">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-487">Name</span></span>| <span data-ttu-id="1e77b-488">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-488">Type</span></span>| <span data-ttu-id="1e77b-489">Attributs</span><span class="sxs-lookup"><span data-stu-id="1e77b-489">Attributes</span></span>| <span data-ttu-id="1e77b-490">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="1e77b-491">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-491">Object</span></span> | <span data-ttu-id="1e77b-492">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-492">&lt;optional&gt;</span></span> | <span data-ttu-id="1e77b-493">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="1e77b-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="1e77b-494">Boolean</span><span class="sxs-lookup"><span data-stu-id="1e77b-494">Boolean</span></span> |  <span data-ttu-id="1e77b-495">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-495">&lt;optional&gt;</span></span> | <span data-ttu-id="1e77b-p136">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services Web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1e77b-498">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-498">Object</span></span> |  <span data-ttu-id="1e77b-499">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-499">&lt;optional&gt;</span></span> | <span data-ttu-id="1e77b-500">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1e77b-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="1e77b-501">fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-501">function</span></span>||<span data-ttu-id="1e77b-p137">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-504">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-504">Requirements</span></span>

|<span data-ttu-id="1e77b-505">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-505">Requirement</span></span>| <span data-ttu-id="1e77b-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-507">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-508">1.5</span><span class="sxs-lookup"><span data-stu-id="1e77b-508">1.5</span></span> |
|[<span data-ttu-id="1e77b-509">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-510">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-512">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-513">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1e77b-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1e77b-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1e77b-515">Obtient une chaîne qui contient un jeton utilisé pour obtenir une pièce jointe ou un élément à partir d’un Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1e77b-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1e77b-p138">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1e77b-p139">Vous pouvez transmettre le jeton et un identificateur de pièce jointe ou un identificateur d'élément à un système tiers. Le système tiers utilise le jeton comme jeton d'autorisation au porteur pour appeler l'opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services Web Exchange, pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1e77b-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1e77b-521">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste, pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="1e77b-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="1e77b-p140">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-524">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-524">Parameters:</span></span>

|<span data-ttu-id="1e77b-525">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-525">Name</span></span>| <span data-ttu-id="1e77b-526">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-526">Type</span></span>| <span data-ttu-id="1e77b-527">Attributs</span><span class="sxs-lookup"><span data-stu-id="1e77b-527">Attributes</span></span>| <span data-ttu-id="1e77b-528">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1e77b-529">fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-529">function</span></span>||<span data-ttu-id="1e77b-p141">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1e77b-532">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-532">Object</span></span>| <span data-ttu-id="1e77b-533">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-533">&lt;optional&gt;</span></span>|<span data-ttu-id="1e77b-534">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1e77b-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-535">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-535">Requirements</span></span>

|<span data-ttu-id="1e77b-536">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-536">Requirement</span></span>| <span data-ttu-id="1e77b-537">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-538">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-539">1.3</span><span class="sxs-lookup"><span data-stu-id="1e77b-539">1.3</span></span>|
|[<span data-ttu-id="1e77b-540">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-541">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-542">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-543">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-544">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1e77b-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1e77b-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1e77b-546">Obtient un jeton identifiant l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="1e77b-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1e77b-547">La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="1e77b-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-548">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-548">Parameters:</span></span>

|<span data-ttu-id="1e77b-549">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-549">Name</span></span>| <span data-ttu-id="1e77b-550">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-550">Type</span></span>| <span data-ttu-id="1e77b-551">Attributs</span><span class="sxs-lookup"><span data-stu-id="1e77b-551">Attributes</span></span>| <span data-ttu-id="1e77b-552">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1e77b-553">fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-553">function</span></span>||<span data-ttu-id="1e77b-554">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1e77b-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1e77b-555">Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1e77b-556">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-556">Object</span></span>| <span data-ttu-id="1e77b-557">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-557">&lt;optional&gt;</span></span>|<span data-ttu-id="1e77b-558">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1e77b-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-559">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-559">Requirements</span></span>

|<span data-ttu-id="1e77b-560">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-560">Requirement</span></span>| <span data-ttu-id="1e77b-561">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-562">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-563">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-563">1.0</span></span>|
|[<span data-ttu-id="1e77b-564">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e77b-565">ReadItem</span></span>|
|[<span data-ttu-id="1e77b-566">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-567">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-568">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1e77b-569">makeEwsRequestAsync (données, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1e77b-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1e77b-570">Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1e77b-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-571">Cette méthode n’est pas prise en charge dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="1e77b-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1e77b-572">Dans Outlook pour iOS ou Outlook pour Android</span><span class="sxs-lookup"><span data-stu-id="1e77b-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="1e77b-573">Lorsque le complément est chargé dans une boîte aux lettres Gmail</span><span class="sxs-lookup"><span data-stu-id="1e77b-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1e77b-574">Dans ces cas, les compléments doivent plutôt [utiliser des API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1e77b-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1e77b-575">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange, de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="1e77b-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="1e77b-576">Pour une liste des opérations EWS prises en charge, voir [Appeler des services Web depuis un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="1e77b-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1e77b-577">Avec la méthode `makeEwsRequestAsync`, vous ne pouvez pas demander des éléments associés à un dossier.</span><span class="sxs-lookup"><span data-stu-id="1e77b-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1e77b-578">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="1e77b-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1e77b-p143">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la rubrique [Spécifier les autorisations d'accès de complément de messagerie à la boîte aux lettres de l'utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="1e77b-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1e77b-581">L’administrateur du serveur doit définir `OAuthAuthentication` sur true dans l'annuaire EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="1e77b-581">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1e77b-582">Différences entre versions</span><span class="sxs-lookup"><span data-stu-id="1e77b-582">Version differences</span></span>

<span data-ttu-id="1e77b-583">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook antérieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1e77b-p144">Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le Web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le Web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="1e77b-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e77b-587">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="1e77b-587">Parameters:</span></span>

|<span data-ttu-id="1e77b-588">Nom</span><span class="sxs-lookup"><span data-stu-id="1e77b-588">Name</span></span>| <span data-ttu-id="1e77b-589">Type</span><span class="sxs-lookup"><span data-stu-id="1e77b-589">Type</span></span>| <span data-ttu-id="1e77b-590">Attributs</span><span class="sxs-lookup"><span data-stu-id="1e77b-590">Attributes</span></span>| <span data-ttu-id="1e77b-591">Description</span><span class="sxs-lookup"><span data-stu-id="1e77b-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1e77b-592">String</span><span class="sxs-lookup"><span data-stu-id="1e77b-592">String</span></span>||<span data-ttu-id="1e77b-593">La demande EWS.</span><span class="sxs-lookup"><span data-stu-id="1e77b-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1e77b-594">fonction</span><span class="sxs-lookup"><span data-stu-id="1e77b-594">function</span></span>||<span data-ttu-id="1e77b-595">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1e77b-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1e77b-596">Le résultat XML de l’appel EWS est fourni comme une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1e77b-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="1e77b-597">Si le résultat dépasse une taille de 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="1e77b-597">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="1e77b-598">Objet</span><span class="sxs-lookup"><span data-stu-id="1e77b-598">Object</span></span>| <span data-ttu-id="1e77b-599">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="1e77b-599">&lt;optional&gt;</span></span>|<span data-ttu-id="1e77b-600">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1e77b-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e77b-601">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1e77b-601">Requirements</span></span>

|<span data-ttu-id="1e77b-602">Condition requise</span><span class="sxs-lookup"><span data-stu-id="1e77b-602">Requirement</span></span>| <span data-ttu-id="1e77b-603">Valeur</span><span class="sxs-lookup"><span data-stu-id="1e77b-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e77b-604">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1e77b-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e77b-605">1.0</span><span class="sxs-lookup"><span data-stu-id="1e77b-605">1.0</span></span>|
|[<span data-ttu-id="1e77b-606">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1e77b-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e77b-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1e77b-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1e77b-608">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1e77b-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e77b-609">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1e77b-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e77b-610">Exemple</span><span class="sxs-lookup"><span data-stu-id="1e77b-610">Example</span></span>

<span data-ttu-id="1e77b-611">L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` et obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="1e77b-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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