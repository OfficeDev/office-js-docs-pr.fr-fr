
# <a name="mailbox"></a><span data-ttu-id="d288b-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="d288b-101">mailbox</span></span>

### <span data-ttu-id="d288b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="d288b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="d288b-104">Donne accès au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="d288b-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d288b-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-105">Requirements</span></span>

|<span data-ttu-id="d288b-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-106">Requirement</span></span>| <span data-ttu-id="d288b-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-108">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-109">1.0</span></span>|
|[<span data-ttu-id="d288b-110">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="d288b-111">Restricted</span></span>|
|[<span data-ttu-id="d288b-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d288b-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d288b-114">Members and methods</span></span>

| <span data-ttu-id="d288b-115">Membre</span><span class="sxs-lookup"><span data-stu-id="d288b-115">Member</span></span> | <span data-ttu-id="d288b-116">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d288b-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d288b-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d288b-118">Membre</span><span class="sxs-lookup"><span data-stu-id="d288b-118">Member</span></span> |
| [<span data-ttu-id="d288b-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d288b-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d288b-120">Membre</span><span class="sxs-lookup"><span data-stu-id="d288b-120">Member</span></span> |
| [<span data-ttu-id="d288b-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d288b-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d288b-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-122">Method</span></span> |
| [<span data-ttu-id="d288b-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d288b-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d288b-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-124">Method</span></span> |
| [<span data-ttu-id="d288b-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d288b-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | <span data-ttu-id="d288b-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-126">Method</span></span> |
| [<span data-ttu-id="d288b-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d288b-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d288b-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-128">Method</span></span> |
| [<span data-ttu-id="d288b-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d288b-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d288b-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-130">Method</span></span> |
| [<span data-ttu-id="d288b-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d288b-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d288b-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-132">Method</span></span> |
| [<span data-ttu-id="d288b-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d288b-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d288b-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-134">Method</span></span> |
| [<span data-ttu-id="d288b-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d288b-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d288b-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-136">Method</span></span> |
| [<span data-ttu-id="d288b-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="d288b-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d288b-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-138">Method</span></span> |
| [<span data-ttu-id="d288b-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d288b-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d288b-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-140">Method</span></span> |
| [<span data-ttu-id="d288b-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d288b-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d288b-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-142">Method</span></span> |
| [<span data-ttu-id="d288b-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d288b-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d288b-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-144">Method</span></span> |
| [<span data-ttu-id="d288b-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d288b-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d288b-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="d288b-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d288b-147">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d288b-147">Namespaces</span></span>

<span data-ttu-id="d288b-148">[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d288b-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d288b-149">[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d288b-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d288b-150">[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d288b-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d288b-151">Membres</span><span class="sxs-lookup"><span data-stu-id="d288b-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d288b-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="d288b-152">ewsUrl :String</span></span>

<span data-ttu-id="d288b-p102">Obtient l’URL du point de terminaison des services Web Exchange  (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d288b-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-155">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-p103">La valeur `ewsUrl` peut être utilisée par un service distant pour effectuer des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir les pièces jointes de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d288b-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d288b-158">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `ewsUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="d288b-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d288b-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d288b-161">Type :</span><span class="sxs-lookup"><span data-stu-id="d288b-161">Type:</span></span>

*   <span data-ttu-id="d288b-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d288b-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d288b-163">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-163">Requirements</span></span>

|<span data-ttu-id="d288b-164">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-164">Requirement</span></span>| <span data-ttu-id="d288b-165">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-166">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-167">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-167">1.0</span></span>|
|[<span data-ttu-id="d288b-168">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-169">ReadItem</span></span>|
|[<span data-ttu-id="d288b-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="d288b-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="d288b-172">restUrl :String</span></span>

<span data-ttu-id="d288b-173">Obtient l’URL du point de terminaison REST de ce compte de courrier.</span><span class="sxs-lookup"><span data-stu-id="d288b-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d288b-174">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d288b-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d288b-175">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="d288b-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d288b-p105">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d288b-178">Type :</span><span class="sxs-lookup"><span data-stu-id="d288b-178">Type:</span></span>

*   <span data-ttu-id="d288b-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d288b-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d288b-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-180">Requirements</span></span>

|<span data-ttu-id="d288b-181">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-181">Requirement</span></span>| <span data-ttu-id="d288b-182">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-183">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-184">1.5</span><span class="sxs-lookup"><span data-stu-id="d288b-184">1.5</span></span> |
|[<span data-ttu-id="d288b-185">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-186">ReadItem</span></span>|
|[<span data-ttu-id="d288b-187">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-188">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="d288b-189">Méthodes</span><span class="sxs-lookup"><span data-stu-id="d288b-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d288b-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d288b-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d288b-191">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="d288b-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d288b-192">Pour le moment, les types d’événements pris en charge sont `Office.EventType.ItemChanged` et `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="d288b-192">Currently the supported event types are `Office.EventType.ItemChanged`, , and `Office.EventType.OfficeThemeChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-193">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-193">Parameters:</span></span>

| <span data-ttu-id="d288b-194">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-194">Name</span></span> | <span data-ttu-id="d288b-195">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-195">Type</span></span> | <span data-ttu-id="d288b-196">Attributs</span><span class="sxs-lookup"><span data-stu-id="d288b-196">Attributes</span></span> | <span data-ttu-id="d288b-197">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d288b-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d288b-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d288b-199">L’événement qui doit invoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="d288b-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d288b-200">Fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-200">Function</span></span> || <span data-ttu-id="d288b-p106">La fonction permettant de gérer l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d'objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d288b-204">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-204">Object</span></span> | <span data-ttu-id="d288b-205">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-205">&lt;optional&gt;</span></span> | <span data-ttu-id="d288b-206">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d288b-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d288b-207">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-207">Object</span></span> | <span data-ttu-id="d288b-208">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-208">&lt;optional&gt;</span></span> | <span data-ttu-id="d288b-209">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d288b-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d288b-210">fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-210">function</span></span>| <span data-ttu-id="d288b-211">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-211">&lt;optional&gt;</span></span>|<span data-ttu-id="d288b-212">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d288b-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-213">Requirements</span></span>

|<span data-ttu-id="d288b-214">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-214">Requirement</span></span>| <span data-ttu-id="d288b-215">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-216">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-217">1.5</span><span class="sxs-lookup"><span data-stu-id="d288b-217">1.5</span></span> |
|[<span data-ttu-id="d288b-218">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-219">ReadItem</span></span> |
|[<span data-ttu-id="d288b-220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-221">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d288b-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d288b-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d288b-224">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="d288b-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-225">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-p107">Les ID d’éléments extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)) utilisent un format différent de celui employé par les services Web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="d288b-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-228">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-228">Parameters:</span></span>

|<span data-ttu-id="d288b-229">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-229">Name</span></span>| <span data-ttu-id="d288b-230">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-230">Type</span></span>| <span data-ttu-id="d288b-231">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d288b-232">String</span><span class="sxs-lookup"><span data-stu-id="d288b-232">String</span></span>|<span data-ttu-id="d288b-233">Un ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="d288b-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d288b-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d288b-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d288b-235">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="d288b-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-236">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-236">Requirements</span></span>

|<span data-ttu-id="d288b-237">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-237">Requirement</span></span>| <span data-ttu-id="d288b-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-239">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-240">1.3</span><span class="sxs-lookup"><span data-stu-id="d288b-240">1.3</span></span>|
|[<span data-ttu-id="d288b-241">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-242">Restreint</span><span class="sxs-lookup"><span data-stu-id="d288b-242">Restricted</span></span>|
|[<span data-ttu-id="d288b-243">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-244">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d288b-245">Retourne :</span><span class="sxs-lookup"><span data-stu-id="d288b-245">Returns:</span></span>

<span data-ttu-id="d288b-246">Type : String</span><span class="sxs-lookup"><span data-stu-id="d288b-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d288b-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="d288b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d288b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="d288b-249">Obtient un dictionnaire contenant des informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="d288b-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d288b-p108">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur client, Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire auquel l’utilisateur s'attend.</span><span class="sxs-lookup"><span data-stu-id="d288b-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d288b-p109">Si l’application de messagerie s'exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="d288b-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-255">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-255">Parameters:</span></span>

|<span data-ttu-id="d288b-256">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-256">Name</span></span>| <span data-ttu-id="d288b-257">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-257">Type</span></span>| <span data-ttu-id="d288b-258">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d288b-259">Date</span><span class="sxs-lookup"><span data-stu-id="d288b-259">Date</span></span>|<span data-ttu-id="d288b-260">Un objet Date</span><span class="sxs-lookup"><span data-stu-id="d288b-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-261">Requirements</span></span>

|<span data-ttu-id="d288b-262">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-262">Requirement</span></span>| <span data-ttu-id="d288b-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-264">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-265">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-265">1.0</span></span>|
|[<span data-ttu-id="d288b-266">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-267">ReadItem</span></span>|
|[<span data-ttu-id="d288b-268">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-269">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d288b-270">Retourne :</span><span class="sxs-lookup"><span data-stu-id="d288b-270">Returns:</span></span>

<span data-ttu-id="d288b-271">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d288b-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d288b-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d288b-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d288b-273">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="d288b-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-274">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-p110">Les ID d’éléments récupérés via EWS ou via la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS dans un format adapté à REST.</span><span class="sxs-lookup"><span data-stu-id="d288b-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-277">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-277">Parameters:</span></span>

|<span data-ttu-id="d288b-278">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-278">Name</span></span>| <span data-ttu-id="d288b-279">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-279">Type</span></span>| <span data-ttu-id="d288b-280">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d288b-281">String</span><span class="sxs-lookup"><span data-stu-id="d288b-281">String</span></span>|<span data-ttu-id="d288b-282">Un ID d’élément mis en forme pour les services Web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="d288b-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d288b-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d288b-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d288b-284">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="d288b-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-285">Requirements</span></span>

|<span data-ttu-id="d288b-286">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-286">Requirement</span></span>| <span data-ttu-id="d288b-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-288">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-289">1.3</span><span class="sxs-lookup"><span data-stu-id="d288b-289">1.3</span></span>|
|[<span data-ttu-id="d288b-290">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-291">Restreint</span><span class="sxs-lookup"><span data-stu-id="d288b-291">Restricted</span></span>|
|[<span data-ttu-id="d288b-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-293">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d288b-294">Retourne :</span><span class="sxs-lookup"><span data-stu-id="d288b-294">Returns:</span></span>

<span data-ttu-id="d288b-295">Type : String</span><span class="sxs-lookup"><span data-stu-id="d288b-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d288b-296">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d288b-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d288b-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d288b-298">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="d288b-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d288b-299">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="d288b-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-300">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-300">Parameters:</span></span>

|<span data-ttu-id="d288b-301">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-301">Name</span></span>| <span data-ttu-id="d288b-302">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-302">Type</span></span>| <span data-ttu-id="d288b-303">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d288b-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d288b-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="d288b-305">Valeur en heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="d288b-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-306">Requirements</span></span>

|<span data-ttu-id="d288b-307">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-307">Requirement</span></span>| <span data-ttu-id="d288b-308">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-309">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-310">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-310">1.0</span></span>|
|[<span data-ttu-id="d288b-311">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-312">ReadItem</span></span>|
|[<span data-ttu-id="d288b-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-314">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d288b-315">Retourne :</span><span class="sxs-lookup"><span data-stu-id="d288b-315">Returns:</span></span>

<span data-ttu-id="d288b-316">Un objet Date avec l’heure exprimée en UTC.</span><span class="sxs-lookup"><span data-stu-id="d288b-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="d288b-317">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="d288b-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d288b-318">Date</span><span class="sxs-lookup"><span data-stu-id="d288b-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="d288b-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d288b-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d288b-320">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="d288b-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-321">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-322">La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les équipements mobiles.</span><span class="sxs-lookup"><span data-stu-id="d288b-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d288b-p111">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="d288b-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d288b-325">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d288b-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d288b-326">Si l’identificateur d’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur n'est retourné.</span><span class="sxs-lookup"><span data-stu-id="d288b-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-327">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-327">Parameters:</span></span>

|<span data-ttu-id="d288b-328">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-328">Name</span></span>| <span data-ttu-id="d288b-329">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-329">Type</span></span>| <span data-ttu-id="d288b-330">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d288b-331">String</span><span class="sxs-lookup"><span data-stu-id="d288b-331">String</span></span>|<span data-ttu-id="d288b-332">L'identificateur EWS (services web Exchange) pour un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="d288b-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-333">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-333">Requirements</span></span>

|<span data-ttu-id="d288b-334">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-334">Requirement</span></span>| <span data-ttu-id="d288b-335">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-336">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-337">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-337">1.0</span></span>|
|[<span data-ttu-id="d288b-338">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-339">ReadItem</span></span>|
|[<span data-ttu-id="d288b-340">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-341">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-342">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="d288b-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d288b-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d288b-344">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="d288b-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-345">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-346">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre sur le bureau, ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="d288b-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d288b-347">Dans Outlook Web App, cette méthode n’ouvre le formulaire indiqué que si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d288b-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d288b-348">Si l’identificateur d’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.</span><span class="sxs-lookup"><span data-stu-id="d288b-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d288b-p112">N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d288b-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-351">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-351">Parameters:</span></span>

|<span data-ttu-id="d288b-352">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-352">Name</span></span>| <span data-ttu-id="d288b-353">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-353">Type</span></span>| <span data-ttu-id="d288b-354">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d288b-355">String</span><span class="sxs-lookup"><span data-stu-id="d288b-355">String</span></span>|<span data-ttu-id="d288b-356">Identificateur EWS (services web Exchange) pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="d288b-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-357">Requirements</span></span>

|<span data-ttu-id="d288b-358">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-358">Requirement</span></span>| <span data-ttu-id="d288b-359">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-360">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-361">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-361">1.0</span></span>|
|[<span data-ttu-id="d288b-362">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-363">ReadItem</span></span>|
|[<span data-ttu-id="d288b-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-365">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-366">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d288b-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d288b-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d288b-368">Affiche un formulaire pour créer un rendez-vous de calendrier.</span><span class="sxs-lookup"><span data-stu-id="d288b-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-369">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d288b-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d288b-p113">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont fournis, les champs du formulaire de rendez-vous sont automatiquement remplis avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="d288b-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d288b-p114">Dans Outlook Web App et OWA pour les appareils, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire inclura les participants et un bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="d288b-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d288b-p115">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans les paramètres `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="d288b-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d288b-377">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="d288b-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-378">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-379">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="d288b-379">Note: All parameters are optional.</span></span>

|<span data-ttu-id="d288b-380">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-380">Name</span></span>| <span data-ttu-id="d288b-381">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-381">Type</span></span>| <span data-ttu-id="d288b-382">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d288b-383">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-383">Object</span></span> | <span data-ttu-id="d288b-384">Un dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d288b-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d288b-385">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d288b-p116">Un tableau de chaînes contenant les adresses de messagerie ou un tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="d288b-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d288b-388">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d288b-p117">Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="d288b-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d288b-391">Date</span><span class="sxs-lookup"><span data-stu-id="d288b-391">Date</span></span> | <span data-ttu-id="d288b-392">Un objet `Date` indiquant la date et l’heure du début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d288b-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d288b-393">Date</span><span class="sxs-lookup"><span data-stu-id="d288b-393">Date</span></span> | <span data-ttu-id="d288b-394">Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d288b-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d288b-395">String</span><span class="sxs-lookup"><span data-stu-id="d288b-395">String</span></span> | <span data-ttu-id="d288b-p118">Une chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d288b-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d288b-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d288b-p119">Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="d288b-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d288b-401">String</span><span class="sxs-lookup"><span data-stu-id="d288b-401">String</span></span> | <span data-ttu-id="d288b-p120">Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d288b-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d288b-404">String</span><span class="sxs-lookup"><span data-stu-id="d288b-404">String</span></span> | <span data-ttu-id="d288b-p121">Le corps du rendez-vous. Le contenu du corps est limité à une taille maximale de 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d288b-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d288b-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-407">Requirements</span></span>

|<span data-ttu-id="d288b-408">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-408">Requirement</span></span>| <span data-ttu-id="d288b-409">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-410">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-411">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-411">1.0</span></span>|
|[<span data-ttu-id="d288b-412">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-413">ReadItem</span></span>|
|[<span data-ttu-id="d288b-414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-415">Lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-416">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d288b-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d288b-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d288b-418">Affiche un formulaire permettant de créer un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="d288b-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d288b-419">La méthode `displayNewMessageForm` ouvre un formulaire qui permet à l’utilisateur de créer un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="d288b-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="d288b-420">Si des paramètres sont spécifiés, les champs du formulaire de message sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="d288b-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d288b-421">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="d288b-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-422">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-423">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="d288b-423">Note: All parameters are optional.</span></span>

|<span data-ttu-id="d288b-424">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-424">Name</span></span>| <span data-ttu-id="d288b-425">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-425">Type</span></span>| <span data-ttu-id="d288b-426">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d288b-427">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-427">Object</span></span> | <span data-ttu-id="d288b-428">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="d288b-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d288b-429">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d288b-430">Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne À.</span><span class="sxs-lookup"><span data-stu-id="d288b-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="d288b-431">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="d288b-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d288b-432">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d288b-433">Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cc.</span><span class="sxs-lookup"><span data-stu-id="d288b-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="d288b-434">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="d288b-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d288b-435">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d288b-436">Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cci.</span><span class="sxs-lookup"><span data-stu-id="d288b-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="d288b-437">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="d288b-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d288b-438">String</span><span class="sxs-lookup"><span data-stu-id="d288b-438">String</span></span> | <span data-ttu-id="d288b-439">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="d288b-439">A string containing the subject of the message.</span></span> <span data-ttu-id="d288b-440">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="d288b-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d288b-441">String</span><span class="sxs-lookup"><span data-stu-id="d288b-441">String</span></span> | <span data-ttu-id="d288b-442">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="d288b-442">The HTML body of the message.</span></span> <span data-ttu-id="d288b-443">Le contenu du corps du message est limitée à une taille maximum de 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d288b-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d288b-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d288b-445">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="d288b-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d288b-446">String</span><span class="sxs-lookup"><span data-stu-id="d288b-446">String</span></span> | <span data-ttu-id="d288b-p128">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="d288b-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d288b-449">String</span><span class="sxs-lookup"><span data-stu-id="d288b-449">String</span></span> | <span data-ttu-id="d288b-450">Chaîne qui contient le nom de la pièce jointe, d’une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d288b-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d288b-451">String</span><span class="sxs-lookup"><span data-stu-id="d288b-451">String</span></span> | <span data-ttu-id="d288b-p129">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="d288b-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d288b-454">Boolean</span><span class="sxs-lookup"><span data-stu-id="d288b-454">Boolean</span></span> | <span data-ttu-id="d288b-p130">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incluse dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d288b-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d288b-457">String</span><span class="sxs-lookup"><span data-stu-id="d288b-457">String</span></span> | <span data-ttu-id="d288b-458">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="d288b-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="d288b-459">L’id d’élément EWS du courrier électronique existant à joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="d288b-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="d288b-460">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d288b-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d288b-461">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-461">Requirements</span></span>

|<span data-ttu-id="d288b-462">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-462">Requirement</span></span>| <span data-ttu-id="d288b-463">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-464">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-465">1.6</span><span class="sxs-lookup"><span data-stu-id="d288b-465">-16</span></span> |
|[<span data-ttu-id="d288b-466">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-467">ReadItem</span></span>|
|[<span data-ttu-id="d288b-468">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-469">Lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-470">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-470">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d288b-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d288b-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d288b-472">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="d288b-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d288b-p132">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l’Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="d288b-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-475">Les compléments doivent, dans la mesure du possible, utiliser les API REST plutôt que les services Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="d288b-475">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="d288b-476">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="d288b-476">**REST Tokens**</span></span>

<span data-ttu-id="d288b-p133">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services Web Exchange. Le jeton peut seulement accéder à l’élément actif et à ses pièces jointes en lecture seule, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="d288b-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d288b-480">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="d288b-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d288b-481">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="d288b-481">**EWS Tokens**</span></span>

<span data-ttu-id="d288b-p134">Lorsque un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton à une étendue limitée à l’accès à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d288b-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d288b-484">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="d288b-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-485">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-485">Parameters:</span></span>

|<span data-ttu-id="d288b-486">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-486">Name</span></span>| <span data-ttu-id="d288b-487">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-487">Type</span></span>| <span data-ttu-id="d288b-488">Attributs</span><span class="sxs-lookup"><span data-stu-id="d288b-488">Attributes</span></span>| <span data-ttu-id="d288b-489">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d288b-490">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-490">Object</span></span> | <span data-ttu-id="d288b-491">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-491">&lt;optional&gt;</span></span> | <span data-ttu-id="d288b-492">Littéral d'objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d288b-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d288b-493">Booléen</span><span class="sxs-lookup"><span data-stu-id="d288b-493">Boolean</span></span> |  <span data-ttu-id="d288b-494">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-494">&lt;optional&gt;</span></span> | <span data-ttu-id="d288b-p135">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services Web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d288b-497">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-497">Object</span></span> |  <span data-ttu-id="d288b-498">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-498">&lt;optional&gt;</span></span> | <span data-ttu-id="d288b-499">Toute donnée d'état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d288b-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d288b-500">fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-500">function</span></span>||<span data-ttu-id="d288b-p136">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-503">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-503">Requirements</span></span>

|<span data-ttu-id="d288b-504">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-504">Requirement</span></span>| <span data-ttu-id="d288b-505">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-506">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-507">1.5</span><span class="sxs-lookup"><span data-stu-id="d288b-507">1.5</span></span> |
|[<span data-ttu-id="d288b-508">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-509">ReadItem</span></span>|
|[<span data-ttu-id="d288b-510">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-511">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-512">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-512">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d288b-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d288b-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d288b-514">Obtient une chaîne qui contient un jeton utilisé pour obtenir une pièce jointe ou un élément à partir d’un Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d288b-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d288b-p137">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir du Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="d288b-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d288b-p138">Vous pouvez transmettre le jeton et un identificateur de pièce jointe ou un identificateur d’élément à un système tiers. Le système tiers utilise le jeton comme jeton d’autorisation au porteur pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services Web Exchange, pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d288b-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d288b-520">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste, pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="d288b-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="d288b-p139">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-523">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-523">Parameters:</span></span>

|<span data-ttu-id="d288b-524">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-524">Name</span></span>| <span data-ttu-id="d288b-525">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-525">Type</span></span>| <span data-ttu-id="d288b-526">Attributs</span><span class="sxs-lookup"><span data-stu-id="d288b-526">Attributes</span></span>| <span data-ttu-id="d288b-527">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d288b-528">fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-528">function</span></span>||<span data-ttu-id="d288b-p140">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d288b-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d288b-531">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-531">Object</span></span>| <span data-ttu-id="d288b-532">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-532">&lt;optional&gt;</span></span>|<span data-ttu-id="d288b-533">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d288b-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-534">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-534">Requirements</span></span>

|<span data-ttu-id="d288b-535">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-535">Requirement</span></span>| <span data-ttu-id="d288b-536">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-537">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-538">1.3</span><span class="sxs-lookup"><span data-stu-id="d288b-538">1.3</span></span>|
|[<span data-ttu-id="d288b-539">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-540">ReadItem</span></span>|
|[<span data-ttu-id="d288b-541">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-542">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-543">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d288b-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d288b-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d288b-545">Obtient un jeton identifiant l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="d288b-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d288b-546">La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d288b-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-547">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-547">Parameters:</span></span>

|<span data-ttu-id="d288b-548">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-548">Name</span></span>| <span data-ttu-id="d288b-549">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-549">Type</span></span>| <span data-ttu-id="d288b-550">Attributs</span><span class="sxs-lookup"><span data-stu-id="d288b-550">Attributes</span></span>| <span data-ttu-id="d288b-551">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d288b-552">fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-552">function</span></span>||<span data-ttu-id="d288b-553">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d288b-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d288b-554">Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d288b-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d288b-555">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-555">Object</span></span>| <span data-ttu-id="d288b-556">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-556">&lt;optional&gt;</span></span>|<span data-ttu-id="d288b-557">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d288b-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-558">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-558">Requirements</span></span>

|<span data-ttu-id="d288b-559">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-559">Requirement</span></span>| <span data-ttu-id="d288b-560">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-561">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-562">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-562">1.0</span></span>|
|[<span data-ttu-id="d288b-563">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d288b-564">ReadItem</span></span>|
|[<span data-ttu-id="d288b-565">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-566">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-567">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d288b-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d288b-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d288b-569">Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d288b-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-570">Cette méthode n’est pas prise en charge dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="d288b-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d288b-571">Dans Outlook pour iOS ou Outlook pour Android</span><span class="sxs-lookup"><span data-stu-id="d288b-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="d288b-572">Lorsque le complément est chargé dans une boîte aux lettres Gmail</span><span class="sxs-lookup"><span data-stu-id="d288b-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d288b-573">Dans ces cas, les compléments doivent plutôt [utiliser des API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d288b-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d288b-574">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange, de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="d288b-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d288b-575">Pour une liste des opérations EWS prises en charge, voir [Appeler des services web depuis un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="d288b-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d288b-576">Vous ne pouvez pas demander les éléments associés au dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="d288b-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d288b-577">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d288b-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d288b-p142">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la page [Indiquer des autorisations pour l'accès du complément de messagerie à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d288b-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d288b-580">L’administrateur du serveur doit définir `OAuthAuthentication` à true dans le dossier EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="d288b-580">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d288b-581">Différences entre versions</span><span class="sxs-lookup"><span data-stu-id="d288b-581">Version differences</span></span>

<span data-ttu-id="d288b-582">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook antérieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d288b-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d288b-p143">Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le Web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le Web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d288b-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d288b-586">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="d288b-586">Parameters:</span></span>

|<span data-ttu-id="d288b-587">Nom</span><span class="sxs-lookup"><span data-stu-id="d288b-587">Name</span></span>| <span data-ttu-id="d288b-588">Type</span><span class="sxs-lookup"><span data-stu-id="d288b-588">Type</span></span>| <span data-ttu-id="d288b-589">Attributs</span><span class="sxs-lookup"><span data-stu-id="d288b-589">Attributes</span></span>| <span data-ttu-id="d288b-590">Description</span><span class="sxs-lookup"><span data-stu-id="d288b-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d288b-591">String</span><span class="sxs-lookup"><span data-stu-id="d288b-591">String</span></span>||<span data-ttu-id="d288b-592">La demande EWS.</span><span class="sxs-lookup"><span data-stu-id="d288b-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d288b-593">fonction</span><span class="sxs-lookup"><span data-stu-id="d288b-593">function</span></span>||<span data-ttu-id="d288b-594">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d288b-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d288b-595">Le résultat XML de l’appel EWS est fourni comme une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d288b-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d288b-596">Si le résultat dépasse une taille de 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="d288b-596">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="d288b-597">Objet</span><span class="sxs-lookup"><span data-stu-id="d288b-597">Object</span></span>| <span data-ttu-id="d288b-598">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d288b-598">&lt;optional&gt;</span></span>|<span data-ttu-id="d288b-599">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d288b-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d288b-600">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d288b-600">Requirements</span></span>

|<span data-ttu-id="d288b-601">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d288b-601">Requirement</span></span>| <span data-ttu-id="d288b-602">Valeur</span><span class="sxs-lookup"><span data-stu-id="d288b-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="d288b-603">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d288b-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d288b-604">1.0</span><span class="sxs-lookup"><span data-stu-id="d288b-604">1.0</span></span>|
|[<span data-ttu-id="d288b-605">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d288b-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d288b-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d288b-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d288b-607">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d288b-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d288b-608">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d288b-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d288b-609">Exemple</span><span class="sxs-lookup"><span data-stu-id="d288b-609">Example</span></span>

<span data-ttu-id="d288b-610">L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` et obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d288b-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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