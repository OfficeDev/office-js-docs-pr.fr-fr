
# <a name="userprofile"></a><span data-ttu-id="445cc-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="445cc-101">userProfile</span></span>

### <span data-ttu-id="445cc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="445cc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="445cc-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="445cc-104">Requirements</span></span>

|<span data-ttu-id="445cc-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="445cc-105">Requirement</span></span>| <span data-ttu-id="445cc-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="445cc-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="445cc-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="445cc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="445cc-108">1.0</span></span>|
|[<span data-ttu-id="445cc-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="445cc-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="445cc-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="445cc-110">ReadItem</span></span>|
|[<span data-ttu-id="445cc-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="445cc-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="445cc-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="445cc-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="445cc-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="445cc-113">Members and methods</span></span>

| <span data-ttu-id="445cc-114">Membre</span><span class="sxs-lookup"><span data-stu-id="445cc-114">Member</span></span> | <span data-ttu-id="445cc-115">Type</span><span class="sxs-lookup"><span data-stu-id="445cc-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="445cc-116">accountType</span><span class="sxs-lookup"><span data-stu-id="445cc-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="445cc-117">Member</span><span class="sxs-lookup"><span data-stu-id="445cc-117">Member</span></span> |
| [<span data-ttu-id="445cc-118">displayName</span><span class="sxs-lookup"><span data-stu-id="445cc-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="445cc-119">Membre</span><span class="sxs-lookup"><span data-stu-id="445cc-119">Member</span></span> |
| [<span data-ttu-id="445cc-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="445cc-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="445cc-121">Membre</span><span class="sxs-lookup"><span data-stu-id="445cc-121">Member</span></span> |
| [<span data-ttu-id="445cc-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="445cc-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="445cc-123">Membre</span><span class="sxs-lookup"><span data-stu-id="445cc-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="445cc-124">Members</span><span class="sxs-lookup"><span data-stu-id="445cc-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="445cc-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="445cc-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="445cc-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="445cc-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="445cc-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="445cc-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="445cc-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="445cc-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="445cc-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-129">Value</span></span> | <span data-ttu-id="445cc-130">Description</span><span class="sxs-lookup"><span data-stu-id="445cc-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="445cc-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="445cc-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="445cc-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="445cc-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="445cc-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="445cc-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="445cc-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="445cc-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="445cc-135">Type :</span><span class="sxs-lookup"><span data-stu-id="445cc-135">Type:</span></span>

*   <span data-ttu-id="445cc-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="445cc-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="445cc-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="445cc-137">Requirements</span></span>

|<span data-ttu-id="445cc-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="445cc-138">Requirement</span></span>| <span data-ttu-id="445cc-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="445cc-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="445cc-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="445cc-141">1.6</span><span class="sxs-lookup"><span data-stu-id="445cc-141">-16</span></span> |
|[<span data-ttu-id="445cc-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="445cc-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="445cc-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="445cc-143">ReadItem</span></span>|
|[<span data-ttu-id="445cc-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="445cc-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="445cc-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="445cc-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="445cc-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="445cc-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="445cc-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="445cc-147">displayName :String</span></span>

<span data-ttu-id="445cc-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="445cc-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="445cc-149">Type :</span><span class="sxs-lookup"><span data-stu-id="445cc-149">Type:</span></span>

*   <span data-ttu-id="445cc-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="445cc-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="445cc-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="445cc-151">Requirements</span></span>

|<span data-ttu-id="445cc-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="445cc-152">Requirement</span></span>| <span data-ttu-id="445cc-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="445cc-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="445cc-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="445cc-155">1.0</span><span class="sxs-lookup"><span data-stu-id="445cc-155">1.0</span></span>|
|[<span data-ttu-id="445cc-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="445cc-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="445cc-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="445cc-157">ReadItem</span></span>|
|[<span data-ttu-id="445cc-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="445cc-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="445cc-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="445cc-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="445cc-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="445cc-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="445cc-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="445cc-161">emailAddress :String</span></span>

<span data-ttu-id="445cc-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="445cc-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="445cc-163">Type :</span><span class="sxs-lookup"><span data-stu-id="445cc-163">Type:</span></span>

*   <span data-ttu-id="445cc-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="445cc-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="445cc-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="445cc-165">Requirements</span></span>

|<span data-ttu-id="445cc-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="445cc-166">Requirement</span></span>| <span data-ttu-id="445cc-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="445cc-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="445cc-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="445cc-169">1.0</span><span class="sxs-lookup"><span data-stu-id="445cc-169">1.0</span></span>|
|[<span data-ttu-id="445cc-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="445cc-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="445cc-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="445cc-171">ReadItem</span></span>|
|[<span data-ttu-id="445cc-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="445cc-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="445cc-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="445cc-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="445cc-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="445cc-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="445cc-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="445cc-175">timeZone :String</span></span>

<span data-ttu-id="445cc-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="445cc-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="445cc-177">Type :</span><span class="sxs-lookup"><span data-stu-id="445cc-177">Type:</span></span>

*   <span data-ttu-id="445cc-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="445cc-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="445cc-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="445cc-179">Requirements</span></span>

|<span data-ttu-id="445cc-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="445cc-180">Requirement</span></span>| <span data-ttu-id="445cc-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="445cc-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="445cc-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="445cc-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="445cc-183">1.0</span><span class="sxs-lookup"><span data-stu-id="445cc-183">1.0</span></span>|
|[<span data-ttu-id="445cc-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="445cc-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="445cc-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="445cc-185">ReadItem</span></span>|
|[<span data-ttu-id="445cc-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="445cc-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="445cc-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="445cc-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="445cc-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="445cc-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```