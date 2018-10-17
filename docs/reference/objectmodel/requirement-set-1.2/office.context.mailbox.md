
# <a name="mailbox"></a><span data-ttu-id="2f409-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="2f409-101">mailbox</span></span>

### <span data-ttu-id="2f409-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="2f409-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="2f409-104">Donne accès au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="2f409-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f409-105">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-105">Requirements</span></span>

|<span data-ttu-id="2f409-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-106">Requirement</span></span>| <span data-ttu-id="2f409-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-108">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-109">1.0</span></span>|
|[<span data-ttu-id="2f409-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2f409-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="2f409-111">Restricted</span></span>|
|[<span data-ttu-id="2f409-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="2f409-114">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="2f409-114">Namespaces</span></span>

<span data-ttu-id="2f409-115">[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2f409-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="2f409-116">[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2f409-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="2f409-117">[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2f409-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="2f409-118">Membres</span><span class="sxs-lookup"><span data-stu-id="2f409-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="2f409-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="2f409-119">ewsUrl :String</span></span>

<span data-ttu-id="2f409-p102">Obtient l’URL du point de terminaison des services Web Exchange  (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2f409-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-122">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2f409-122">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2f409-p103">La valeur `ewsUrl` peut être utilisée par un service distant pour effectuer des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir les pièces jointes de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2f409-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="2f409-125">Type :</span><span class="sxs-lookup"><span data-stu-id="2f409-125">Type:</span></span>

*   <span data-ttu-id="2f409-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f409-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f409-127">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-127">Requirements</span></span>

|<span data-ttu-id="2f409-128">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-128">Requirement</span></span>| <span data-ttu-id="2f409-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-130">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-131">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-131">1.0</span></span>|
|[<span data-ttu-id="2f409-132">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-133">ReadItem</span></span>|
|[<span data-ttu-id="2f409-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-135">Lire</span><span class="sxs-lookup"><span data-stu-id="2f409-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2f409-136">Méthodes</span><span class="sxs-lookup"><span data-stu-id="2f409-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime"></a><span data-ttu-id="2f409-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="2f409-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span></span>

<span data-ttu-id="2f409-138">Obtient un dictionnaire contenant des informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="2f409-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="2f409-p104">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur client, Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire auquel l’utilisateur s'attend.</span><span class="sxs-lookup"><span data-stu-id="2f409-p104">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="2f409-p105">Si l’application de messagerie s'exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="2f409-p105">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-144">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-144">Parameters:</span></span>

|<span data-ttu-id="2f409-145">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-145">Name</span></span>| <span data-ttu-id="2f409-146">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-146">Type</span></span>| <span data-ttu-id="2f409-147">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="2f409-148">Date</span><span class="sxs-lookup"><span data-stu-id="2f409-148">Date</span></span>|<span data-ttu-id="2f409-149">Un objet Date</span><span class="sxs-lookup"><span data-stu-id="2f409-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-150">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-150">Requirements</span></span>

|<span data-ttu-id="2f409-151">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f409-151">Requirement</span></span>| <span data-ttu-id="2f409-152">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-153">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-154">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-154">1.0</span></span>|
|[<span data-ttu-id="2f409-155">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-156">ReadItem</span></span>|
|[<span data-ttu-id="2f409-157">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-158">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-158">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2f409-159">Retourne :</span><span class="sxs-lookup"><span data-stu-id="2f409-159">Returns:</span></span>

<span data-ttu-id="2f409-160">Type : [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="2f409-160">Type: [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="2f409-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="2f409-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="2f409-162">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="2f409-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="2f409-163">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="2f409-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-164">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-164">Parameters:</span></span>

|<span data-ttu-id="2f409-165">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-165">Name</span></span>| <span data-ttu-id="2f409-166">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-166">Type</span></span>| <span data-ttu-id="2f409-167">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="2f409-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2f409-168">LocalClientTime</span></span>](/javascript/api/outlook_1_2/office.LocalClientTime)|<span data-ttu-id="2f409-169">Valeur en heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="2f409-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-170">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f409-170">Requirements</span></span>

|<span data-ttu-id="2f409-171">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-171">Requirement</span></span>| <span data-ttu-id="2f409-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-173">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-174">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-174">1.0</span></span>|
|[<span data-ttu-id="2f409-175">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-176">ReadItem</span></span>|
|[<span data-ttu-id="2f409-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-178">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-178">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2f409-179">Retourne :</span><span class="sxs-lookup"><span data-stu-id="2f409-179">Returns:</span></span>

<span data-ttu-id="2f409-180">Un objet Date avec l’heure exprimée en UTC.</span><span class="sxs-lookup"><span data-stu-id="2f409-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="2f409-181">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2f409-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2f409-182">Date</span><span class="sxs-lookup"><span data-stu-id="2f409-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="2f409-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2f409-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="2f409-184">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="2f409-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-185">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2f409-185">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2f409-186">La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les équipements mobiles.</span><span class="sxs-lookup"><span data-stu-id="2f409-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2f409-p106">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="2f409-p106">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="2f409-189">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2f409-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="2f409-190">Si l’identificateur d’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur n'est retourné.</span><span class="sxs-lookup"><span data-stu-id="2f409-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-191">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-191">Parameters:</span></span>

|<span data-ttu-id="2f409-192">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-192">Name</span></span>| <span data-ttu-id="2f409-193">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-193">Type</span></span>| <span data-ttu-id="2f409-194">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2f409-195">String</span><span class="sxs-lookup"><span data-stu-id="2f409-195">String</span></span>|<span data-ttu-id="2f409-196">L'identificateur EWS (services web Exchange) pour un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="2f409-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-197">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-197">Requirements</span></span>

|<span data-ttu-id="2f409-198">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-198">Requirement</span></span>| <span data-ttu-id="2f409-199">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-200">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-201">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-201">1.0</span></span>|
|[<span data-ttu-id="2f409-202">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-203">ReadItem</span></span>|
|[<span data-ttu-id="2f409-204">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-205">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-205">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-206">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="2f409-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2f409-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="2f409-208">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="2f409-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-209">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2f409-209">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2f409-210">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre sur le bureau, ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="2f409-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2f409-211">Dans Outlook Web App, cette méthode n’ouvre le formulaire indiqué que si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2f409-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="2f409-212">Si l’identificateur d’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.</span><span class="sxs-lookup"><span data-stu-id="2f409-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="2f409-p107">N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2f409-p107">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-215">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-215">Parameters:</span></span>

|<span data-ttu-id="2f409-216">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-216">Name</span></span>| <span data-ttu-id="2f409-217">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-217">Type</span></span>| <span data-ttu-id="2f409-218">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2f409-219">String</span><span class="sxs-lookup"><span data-stu-id="2f409-219">String</span></span>|<span data-ttu-id="2f409-220">Identificateur EWS (services web Exchange) pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="2f409-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-221">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-221">Requirements</span></span>

|<span data-ttu-id="2f409-222">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-222">Requirement</span></span>| <span data-ttu-id="2f409-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-224">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-225">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-225">1.0</span></span>|
|[<span data-ttu-id="2f409-226">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-227">ReadItem</span></span>|
|[<span data-ttu-id="2f409-228">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-229">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-229">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-230">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-230">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="2f409-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2f409-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="2f409-232">Affiche un formulaire pour créer un rendez-vous de calendrier.</span><span class="sxs-lookup"><span data-stu-id="2f409-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-233">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="2f409-233">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2f409-p108">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont fournis, les champs du formulaire de rendez-vous sont automatiquement remplis avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="2f409-p108">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2f409-p109">Dans Outlook Web App et OWA pour les appareils, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire inclura les participants et un bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="2f409-p109">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="2f409-p110">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans les paramètres `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="2f409-p110">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="2f409-241">Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="2f409-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-242">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-242">Parameters:</span></span>

|<span data-ttu-id="2f409-243">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-243">Name</span></span>| <span data-ttu-id="2f409-244">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-244">Type</span></span>| <span data-ttu-id="2f409-245">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2f409-246">Objet</span><span class="sxs-lookup"><span data-stu-id="2f409-246">Object</span></span> | <span data-ttu-id="2f409-247">Un dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2f409-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="2f409-248">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2f409-p111">Un tableau de chaînes contenant les adresses de messagerie ou un tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="2f409-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="2f409-251">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2f409-p112">Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="2f409-p112">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="2f409-254">Date</span><span class="sxs-lookup"><span data-stu-id="2f409-254">Date</span></span> | <span data-ttu-id="2f409-255">Un objet `Date` indiquant la date et l’heure du début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2f409-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="2f409-256">Date</span><span class="sxs-lookup"><span data-stu-id="2f409-256">Date</span></span> | <span data-ttu-id="2f409-257">Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2f409-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="2f409-258">String</span><span class="sxs-lookup"><span data-stu-id="2f409-258">String</span></span> | <span data-ttu-id="2f409-p113">Une chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2f409-p113">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="2f409-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="2f409-p114">Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées.</span><span class="sxs-lookup"><span data-stu-id="2f409-p114">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2f409-264">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f409-264">String</span></span> | <span data-ttu-id="2f409-p115">Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée à un maximum de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2f409-p115">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="2f409-267">String</span><span class="sxs-lookup"><span data-stu-id="2f409-267">String</span></span> | <span data-ttu-id="2f409-p116">Le corps du rendez-vous. Le contenu du corps est limité à une taille maximale de 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2f409-p116">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2f409-270">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-270">Requirements</span></span>

|<span data-ttu-id="2f409-271">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-271">Requirement</span></span>| <span data-ttu-id="2f409-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-273">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-274">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-274">1.0</span></span>|
|[<span data-ttu-id="2f409-275">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-276">ReadItem</span></span>|
|[<span data-ttu-id="2f409-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-278">Lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="2f409-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2f409-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2f409-281">Obtient une chaîne qui contient un jeton utilisé pour obtenir une pièce jointe ou un élément à partir d’un Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2f409-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="2f409-p117">La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir du Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="2f409-p117">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="2f409-p118">Vous pouvez passer le jeton et un identificateur de pièce jointe ou un identificateur d’élément à un système de tierce partie. Celui-ci utilise le jeton en tant que jeton d’autorisation du porteur pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des Exchange Web Services (EWS) pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2f409-p118">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2f409-287">Votre application doit disposer de l’autorisation **ReadItem** indiquée dans son manifeste pour appeler la méthode `getCallbackTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2f409-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-288">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-288">Parameters:</span></span>

|<span data-ttu-id="2f409-289">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-289">Name</span></span>| <span data-ttu-id="2f409-290">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-290">Type</span></span>| <span data-ttu-id="2f409-291">Attributs</span><span class="sxs-lookup"><span data-stu-id="2f409-291">Attributes</span></span>| <span data-ttu-id="2f409-292">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2f409-293">fonction</span><span class="sxs-lookup"><span data-stu-id="2f409-293">function</span></span>||<span data-ttu-id="2f409-294">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2f409-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2f409-295">Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2f409-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2f409-296">Objet</span><span class="sxs-lookup"><span data-stu-id="2f409-296">Object</span></span>| <span data-ttu-id="2f409-297">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-297">&lt;optional&gt;</span></span>|<span data-ttu-id="2f409-298">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2f409-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-299">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-299">Requirements</span></span>

|<span data-ttu-id="2f409-300">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-300">Requirement</span></span>| <span data-ttu-id="2f409-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-302">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-303">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-303">1.0</span></span>|
|[<span data-ttu-id="2f409-304">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-305">ReadItem</span></span>|
|[<span data-ttu-id="2f409-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-307">Lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-308">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="2f409-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2f409-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2f409-310">Obtient un jeton identifiant l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="2f409-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="2f409-311">La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="2f409-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-312">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-312">Parameters:</span></span>

|<span data-ttu-id="2f409-313">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-313">Name</span></span>| <span data-ttu-id="2f409-314">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-314">Type</span></span>| <span data-ttu-id="2f409-315">Attributs</span><span class="sxs-lookup"><span data-stu-id="2f409-315">Attributes</span></span>| <span data-ttu-id="2f409-316">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2f409-317">fonction</span><span class="sxs-lookup"><span data-stu-id="2f409-317">function</span></span>||<span data-ttu-id="2f409-318">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2f409-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2f409-319">Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2f409-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2f409-320">Objet</span><span class="sxs-lookup"><span data-stu-id="2f409-320">Object</span></span>| <span data-ttu-id="2f409-321">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-321">&lt;optional&gt;</span></span>|<span data-ttu-id="2f409-322">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2f409-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-323">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f409-323">Requirements</span></span>

|<span data-ttu-id="2f409-324">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2f409-324">Requirement</span></span>| <span data-ttu-id="2f409-325">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-326">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-327">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-327">1.0</span></span>|
|[<span data-ttu-id="2f409-328">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2f409-329">ReadItem</span></span>|
|[<span data-ttu-id="2f409-330">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-331">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-331">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-332">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-332">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="2f409-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2f409-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="2f409-334">Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2f409-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-335">Cette méthode n’est pas prise en charge dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="2f409-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="2f409-336">Dans Outlook pour iOS ou Outlook pour Android</span><span class="sxs-lookup"><span data-stu-id="2f409-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="2f409-337">Lorsque le complément est chargé dans une boîte aux lettres Gmail</span><span class="sxs-lookup"><span data-stu-id="2f409-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="2f409-338">Dans ces cas, les compléments doivent [utiliser l’API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) à la place pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2f409-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="2f409-p119">Le `makeEwsRequestAsync` méthode envoie une demande EWS au nom de la macro complémentaire à Exchange. Pour obtenir la liste des opérations EWS prises en charge, voir [appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="2f409-p119">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="2f409-341">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="2f409-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="2f409-342">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="2f409-342">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="2f409-p120">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la page [Indiquer des autorisations pour l'accès du complément de messagerie à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="2f409-p120">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="2f409-345">L’administrateur du serveur doit définir `OAuthAuthentication` à true dans le dossier EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="2f409-345">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="2f409-346">Différences entre versions</span><span class="sxs-lookup"><span data-stu-id="2f409-346">Version differences</span></span>

<span data-ttu-id="2f409-347">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook antérieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="2f409-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="2f409-p121">Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le Web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le Web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="2f409-p121">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2f409-351">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="2f409-351">Parameters:</span></span>

|<span data-ttu-id="2f409-352">Nom</span><span class="sxs-lookup"><span data-stu-id="2f409-352">Name</span></span>| <span data-ttu-id="2f409-353">Type</span><span class="sxs-lookup"><span data-stu-id="2f409-353">Type</span></span>| <span data-ttu-id="2f409-354">Attributs</span><span class="sxs-lookup"><span data-stu-id="2f409-354">Attributes</span></span>| <span data-ttu-id="2f409-355">Description</span><span class="sxs-lookup"><span data-stu-id="2f409-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2f409-356">String</span><span class="sxs-lookup"><span data-stu-id="2f409-356">String</span></span>||<span data-ttu-id="2f409-357">La demande EWS.</span><span class="sxs-lookup"><span data-stu-id="2f409-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="2f409-358">fonction</span><span class="sxs-lookup"><span data-stu-id="2f409-358">function</span></span>||<span data-ttu-id="2f409-359">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2f409-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2f409-p122">Le résultat de l’appel EWS XML est fourni sous forme de chaîne dans la `asyncResult.value` propriété. Si le résultat dépasse 1 Mo taille, un message d’erreur est renvoyé à la place.</span><span class="sxs-lookup"><span data-stu-id="2f409-p122">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="2f409-362">Objet</span><span class="sxs-lookup"><span data-stu-id="2f409-362">Object</span></span>| <span data-ttu-id="2f409-363">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="2f409-363">&lt;optional&gt;</span></span>|<span data-ttu-id="2f409-364">Toute donnée d’état qui est passée à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2f409-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f409-365">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="2f409-365">Requirements</span></span>

|<span data-ttu-id="2f409-366">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f409-366">Requirement</span></span>| <span data-ttu-id="2f409-367">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f409-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f409-368">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f409-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2f409-369">1.0</span><span class="sxs-lookup"><span data-stu-id="2f409-369">1.0</span></span>|
|[<span data-ttu-id="2f409-370">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="2f409-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2f409-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2f409-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="2f409-372">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f409-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2f409-373">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f409-373">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f409-374">Exemple</span><span class="sxs-lookup"><span data-stu-id="2f409-374">Example</span></span>

<span data-ttu-id="2f409-375">L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` et obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2f409-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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