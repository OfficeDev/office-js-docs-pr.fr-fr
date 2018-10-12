
# <a name="userprofile"></a><span data-ttu-id="f8c46-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="f8c46-101">userProfile</span></span>

### <span data-ttu-id="f8c46-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="f8c46-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8c46-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8c46-104">Requirements</span></span>

|<span data-ttu-id="f8c46-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f8c46-105">Requirement</span></span>| <span data-ttu-id="f8c46-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8c46-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8c46-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8c46-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f8c46-108">1.0</span></span>|
|[<span data-ttu-id="f8c46-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f8c46-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8c46-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8c46-110">ReadItem</span></span>|
|[<span data-ttu-id="f8c46-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8c46-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8c46-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8c46-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f8c46-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="f8c46-113">Members and methods</span></span>

| <span data-ttu-id="f8c46-114">Membre</span><span class="sxs-lookup"><span data-stu-id="f8c46-114">Member</span></span> | <span data-ttu-id="f8c46-115">Type</span><span class="sxs-lookup"><span data-stu-id="f8c46-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="f8c46-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="f8c46-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="f8c46-117">Membre</span><span class="sxs-lookup"><span data-stu-id="f8c46-117">Member</span></span> |
| [<span data-ttu-id="f8c46-118">displayName</span><span class="sxs-lookup"><span data-stu-id="f8c46-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="f8c46-119">Membre</span><span class="sxs-lookup"><span data-stu-id="f8c46-119">Member</span></span> |
| [<span data-ttu-id="f8c46-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f8c46-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f8c46-121">Membre</span><span class="sxs-lookup"><span data-stu-id="f8c46-121">Member</span></span> |
| [<span data-ttu-id="f8c46-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="f8c46-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f8c46-123">Membre</span><span class="sxs-lookup"><span data-stu-id="f8c46-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f8c46-124">Membres</span><span class="sxs-lookup"><span data-stu-id="f8c46-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="f8c46-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="f8c46-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="f8c46-126">Ce membre est actuellement uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="f8c46-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="f8c46-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="f8c46-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="f8c46-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="f8c46-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="f8c46-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-129">Value</span></span> | <span data-ttu-id="f8c46-130">Description</span><span class="sxs-lookup"><span data-stu-id="f8c46-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="f8c46-131">La boîte aux lettres est sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="f8c46-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="f8c46-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="f8c46-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="f8c46-133">La boîte aux lettres est associée avec un compte Office 365 professionnel ou scolaire.</span><span class="sxs-lookup"><span data-stu-id="f8c46-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="f8c46-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="f8c46-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="f8c46-135">Type :</span><span class="sxs-lookup"><span data-stu-id="f8c46-135">Type:</span></span>

*   <span data-ttu-id="f8c46-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8c46-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8c46-137">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8c46-137">Requirements</span></span>

|<span data-ttu-id="f8c46-138">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f8c46-138">Requirement</span></span>| <span data-ttu-id="f8c46-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8c46-140">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8c46-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8c46-141">1.6</span><span class="sxs-lookup"><span data-stu-id="f8c46-141">-16</span></span> |
|[<span data-ttu-id="f8c46-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f8c46-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8c46-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8c46-143">ReadItem</span></span>|
|[<span data-ttu-id="f8c46-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8c46-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8c46-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8c46-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8c46-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="f8c46-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="f8c46-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="f8c46-147">displayName :String</span></span>

<span data-ttu-id="f8c46-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f8c46-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f8c46-149">Type :</span><span class="sxs-lookup"><span data-stu-id="f8c46-149">Type:</span></span>

*   <span data-ttu-id="f8c46-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8c46-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8c46-151">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8c46-151">Requirements</span></span>

|<span data-ttu-id="f8c46-152">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f8c46-152">Requirement</span></span>| <span data-ttu-id="f8c46-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8c46-154">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8c46-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8c46-155">1.0</span><span class="sxs-lookup"><span data-stu-id="f8c46-155">1.0</span></span>|
|[<span data-ttu-id="f8c46-156">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f8c46-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8c46-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8c46-157">ReadItem</span></span>|
|[<span data-ttu-id="f8c46-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8c46-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8c46-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8c46-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8c46-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="f8c46-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="f8c46-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="f8c46-161">emailAddress :String</span></span>

<span data-ttu-id="f8c46-162">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f8c46-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f8c46-163">Type :</span><span class="sxs-lookup"><span data-stu-id="f8c46-163">Type:</span></span>

*   <span data-ttu-id="f8c46-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8c46-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8c46-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8c46-165">Requirements</span></span>

|<span data-ttu-id="f8c46-166">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f8c46-166">Requirement</span></span>| <span data-ttu-id="f8c46-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8c46-168">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8c46-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8c46-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f8c46-169">1.0</span></span>|
|[<span data-ttu-id="f8c46-170">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f8c46-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8c46-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8c46-171">ReadItem</span></span>|
|[<span data-ttu-id="f8c46-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8c46-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8c46-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8c46-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8c46-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="f8c46-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="f8c46-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="f8c46-175">timeZone :String</span></span>

<span data-ttu-id="f8c46-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f8c46-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f8c46-177">Type :</span><span class="sxs-lookup"><span data-stu-id="f8c46-177">Type:</span></span>

*   <span data-ttu-id="f8c46-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8c46-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8c46-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8c46-179">Requirements</span></span>

|<span data-ttu-id="f8c46-180">Condition requise</span><span class="sxs-lookup"><span data-stu-id="f8c46-180">Requirement</span></span>| <span data-ttu-id="f8c46-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8c46-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8c46-182">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8c46-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8c46-183">1.0</span><span class="sxs-lookup"><span data-stu-id="f8c46-183">1.0</span></span>|
|[<span data-ttu-id="f8c46-184">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="f8c46-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8c46-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8c46-185">ReadItem</span></span>|
|[<span data-ttu-id="f8c46-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8c46-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8c46-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8c46-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8c46-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="f8c46-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```