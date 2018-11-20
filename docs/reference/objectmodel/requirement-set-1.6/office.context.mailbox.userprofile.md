
# <a name="userprofile"></a><span data-ttu-id="1300e-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="1300e-101">userProfile</span></span>

### <span data-ttu-id="1300e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="1300e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1300e-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1300e-104">Requirements</span></span>

|<span data-ttu-id="1300e-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1300e-105">Requirement</span></span>| <span data-ttu-id="1300e-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1300e-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1300e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1300e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1300e-108">1.0</span></span>|
|[<span data-ttu-id="1300e-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1300e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1300e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1300e-110">ReadItem</span></span>|
|[<span data-ttu-id="1300e-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1300e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1300e-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1300e-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1300e-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="1300e-113">Members and methods</span></span>

| <span data-ttu-id="1300e-114">Membre</span><span class="sxs-lookup"><span data-stu-id="1300e-114">Member</span></span> | <span data-ttu-id="1300e-115">Type</span><span class="sxs-lookup"><span data-stu-id="1300e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1300e-116">accountType</span><span class="sxs-lookup"><span data-stu-id="1300e-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="1300e-117">Member</span><span class="sxs-lookup"><span data-stu-id="1300e-117">Member</span></span> |
| [<span data-ttu-id="1300e-118">displayName</span><span class="sxs-lookup"><span data-stu-id="1300e-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="1300e-119">Membre</span><span class="sxs-lookup"><span data-stu-id="1300e-119">Member</span></span> |
| [<span data-ttu-id="1300e-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1300e-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="1300e-121">Membre</span><span class="sxs-lookup"><span data-stu-id="1300e-121">Member</span></span> |
| [<span data-ttu-id="1300e-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="1300e-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="1300e-123">Membre</span><span class="sxs-lookup"><span data-stu-id="1300e-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1300e-124">Members</span><span class="sxs-lookup"><span data-stu-id="1300e-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="1300e-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="1300e-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="1300e-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="1300e-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="1300e-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1300e-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="1300e-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="1300e-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="1300e-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-129">Value</span></span> | <span data-ttu-id="1300e-130">Description</span><span class="sxs-lookup"><span data-stu-id="1300e-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="1300e-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="1300e-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="1300e-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="1300e-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="1300e-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="1300e-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="1300e-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="1300e-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="1300e-135">Type :</span><span class="sxs-lookup"><span data-stu-id="1300e-135">Type:</span></span>

*   <span data-ttu-id="1300e-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1300e-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1300e-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1300e-137">Requirements</span></span>

|<span data-ttu-id="1300e-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1300e-138">Requirement</span></span>| <span data-ttu-id="1300e-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="1300e-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1300e-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1300e-141">1.6</span><span class="sxs-lookup"><span data-stu-id="1300e-141">-16</span></span> |
|[<span data-ttu-id="1300e-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1300e-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1300e-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1300e-143">ReadItem</span></span>|
|[<span data-ttu-id="1300e-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1300e-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1300e-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1300e-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1300e-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="1300e-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="1300e-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="1300e-147">displayName :String</span></span>

<span data-ttu-id="1300e-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1300e-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1300e-149">Type :</span><span class="sxs-lookup"><span data-stu-id="1300e-149">Type:</span></span>

*   <span data-ttu-id="1300e-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1300e-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1300e-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1300e-151">Requirements</span></span>

|<span data-ttu-id="1300e-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1300e-152">Requirement</span></span>| <span data-ttu-id="1300e-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="1300e-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1300e-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1300e-155">1.0</span><span class="sxs-lookup"><span data-stu-id="1300e-155">1.0</span></span>|
|[<span data-ttu-id="1300e-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1300e-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1300e-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1300e-157">ReadItem</span></span>|
|[<span data-ttu-id="1300e-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1300e-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1300e-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1300e-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1300e-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="1300e-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="1300e-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="1300e-161">emailAddress :String</span></span>

<span data-ttu-id="1300e-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1300e-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1300e-163">Type :</span><span class="sxs-lookup"><span data-stu-id="1300e-163">Type:</span></span>

*   <span data-ttu-id="1300e-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1300e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1300e-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1300e-165">Requirements</span></span>

|<span data-ttu-id="1300e-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1300e-166">Requirement</span></span>| <span data-ttu-id="1300e-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1300e-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1300e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1300e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="1300e-169">1.0</span></span>|
|[<span data-ttu-id="1300e-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1300e-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1300e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1300e-171">ReadItem</span></span>|
|[<span data-ttu-id="1300e-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1300e-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1300e-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1300e-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1300e-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="1300e-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="1300e-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="1300e-175">timeZone :String</span></span>

<span data-ttu-id="1300e-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1300e-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1300e-177">Type :</span><span class="sxs-lookup"><span data-stu-id="1300e-177">Type:</span></span>

*   <span data-ttu-id="1300e-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1300e-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1300e-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1300e-179">Requirements</span></span>

|<span data-ttu-id="1300e-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1300e-180">Requirement</span></span>| <span data-ttu-id="1300e-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="1300e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="1300e-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1300e-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1300e-183">1.0</span><span class="sxs-lookup"><span data-stu-id="1300e-183">1.0</span></span>|
|[<span data-ttu-id="1300e-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1300e-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1300e-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1300e-185">ReadItem</span></span>|
|[<span data-ttu-id="1300e-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1300e-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1300e-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="1300e-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1300e-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="1300e-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```