
# <a name="userprofile"></a><span data-ttu-id="e8ee9-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="e8ee9-101">userProfile</span></span>

### <span data-ttu-id="e8ee9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="e8ee9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee9-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e8ee9-104">Requirements</span></span>

|<span data-ttu-id="e8ee9-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e8ee9-105">Requirement</span></span>| <span data-ttu-id="e8ee9-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee9-107">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee9-108">1.0</span></span>|
|[<span data-ttu-id="e8ee9-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e8ee9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee9-110">ReadItem</span></span>|
|[<span data-ttu-id="e8ee9-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e8ee9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee9-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e8ee9-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e8ee9-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e8ee9-113">Members and methods</span></span>

| <span data-ttu-id="e8ee9-114">Membre</span><span class="sxs-lookup"><span data-stu-id="e8ee9-114">Member</span></span> | <span data-ttu-id="e8ee9-115">Type</span><span class="sxs-lookup"><span data-stu-id="e8ee9-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="e8ee9-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="e8ee9-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="e8ee9-117">Membre</span><span class="sxs-lookup"><span data-stu-id="e8ee9-117">Member</span></span> |
| [<span data-ttu-id="e8ee9-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e8ee9-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e8ee9-119">Membre</span><span class="sxs-lookup"><span data-stu-id="e8ee9-119">Member</span></span> |
| [<span data-ttu-id="e8ee9-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e8ee9-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e8ee9-121">Membre</span><span class="sxs-lookup"><span data-stu-id="e8ee9-121">Member</span></span> |
| [<span data-ttu-id="e8ee9-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e8ee9-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e8ee9-123">Membre</span><span class="sxs-lookup"><span data-stu-id="e8ee9-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e8ee9-124">Membres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e8ee9-125">accountType : chaîne</span><span class="sxs-lookup"><span data-stu-id="e8ee9-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e8ee9-126">Ce membre est uniquement pris en charge dans Outlook 2016 pour Mac, build 16.9.1212 et versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="e8ee9-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e8ee9-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="e8ee9-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-129">Value</span></span> | <span data-ttu-id="e8ee9-130">Description</span><span class="sxs-lookup"><span data-stu-id="e8ee9-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e8ee9-131">La boîte aux lettres est sur un serveur Exchange localement.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e8ee9-132">La boîte aux lettres est associé à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e8ee9-133">La boîte aux lettres est associé avec un compte Office 365 professionnel ou scolaire.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e8ee9-134">La boîte aux lettres est associé à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e8ee9-135">Type :</span><span class="sxs-lookup"><span data-stu-id="e8ee9-135">Type:</span></span>

*   <span data-ttu-id="e8ee9-136">String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee9-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e8ee9-137">Requirements</span></span>

|<span data-ttu-id="e8ee9-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e8ee9-138">Requirement</span></span>| <span data-ttu-id="e8ee9-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee9-140">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee9-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e8ee9-141">-16</span></span> |
|[<span data-ttu-id="e8ee9-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e8ee9-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee9-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee9-143">ReadItem</span></span>|
|[<span data-ttu-id="e8ee9-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e8ee9-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee9-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e8ee9-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee9-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="e8ee9-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="e8ee9-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-147">displayName :String</span></span>

<span data-ttu-id="e8ee9-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee9-149">Type :</span><span class="sxs-lookup"><span data-stu-id="e8ee9-149">Type:</span></span>

*   <span data-ttu-id="e8ee9-150">String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee9-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e8ee9-151">Requirements</span></span>

|<span data-ttu-id="e8ee9-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e8ee9-152">Requirement</span></span>| <span data-ttu-id="e8ee9-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee9-154">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee9-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee9-155">1.0</span></span>|
|[<span data-ttu-id="e8ee9-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e8ee9-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee9-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee9-157">ReadItem</span></span>|
|[<span data-ttu-id="e8ee9-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e8ee9-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee9-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e8ee9-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee9-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="e8ee9-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e8ee9-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-161">emailAddress :String</span></span>

<span data-ttu-id="e8ee9-162">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee9-163">Type :</span><span class="sxs-lookup"><span data-stu-id="e8ee9-163">Type:</span></span>

*   <span data-ttu-id="e8ee9-164">String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee9-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e8ee9-165">Requirements</span></span>

|<span data-ttu-id="e8ee9-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e8ee9-166">Requirement</span></span>| <span data-ttu-id="e8ee9-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee9-168">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee9-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee9-169">1.0</span></span>|
|[<span data-ttu-id="e8ee9-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e8ee9-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee9-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee9-171">ReadItem</span></span>|
|[<span data-ttu-id="e8ee9-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e8ee9-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee9-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e8ee9-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee9-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="e8ee9-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e8ee9-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-175">timeZone :String</span></span>

<span data-ttu-id="e8ee9-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e8ee9-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee9-177">Type :</span><span class="sxs-lookup"><span data-stu-id="e8ee9-177">Type:</span></span>

*   <span data-ttu-id="e8ee9-178">String</span><span class="sxs-lookup"><span data-stu-id="e8ee9-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee9-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e8ee9-179">Requirements</span></span>

|<span data-ttu-id="e8ee9-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e8ee9-180">Requirement</span></span>| <span data-ttu-id="e8ee9-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="e8ee9-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee9-182">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e8ee9-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee9-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee9-183">1.0</span></span>|
|[<span data-ttu-id="e8ee9-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e8ee9-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee9-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee9-185">ReadItem</span></span>|
|[<span data-ttu-id="e8ee9-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e8ee9-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee9-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e8ee9-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee9-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="e8ee9-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```