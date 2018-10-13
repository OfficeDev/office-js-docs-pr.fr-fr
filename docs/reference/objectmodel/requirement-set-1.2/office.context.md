
# <a name="context"></a><span data-ttu-id="aac71-101">context</span><span class="sxs-lookup"><span data-stu-id="aac71-101">context</span></span>

### <span data-ttu-id="aac71-p101">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="aac71-p101">[Office](Office.md). context</span></span>

<span data-ttu-id="aac71-p102">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context dans l’interface API partagée](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="aac71-p102">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="aac71-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="aac71-106">Requirements</span></span>

|<span data-ttu-id="aac71-107">Condition requise</span><span class="sxs-lookup"><span data-stu-id="aac71-107">Requirement</span></span>| <span data-ttu-id="aac71-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="aac71-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="aac71-109">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="aac71-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aac71-110">1.0</span><span class="sxs-lookup"><span data-stu-id="aac71-110">1.0</span></span>|
|[<span data-ttu-id="aac71-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="aac71-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aac71-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="aac71-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="aac71-113">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="aac71-113">Namespaces</span></span>

<span data-ttu-id="aac71-114">[mailbox](office.context.mailbox.md) : permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="aac71-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="aac71-115">Membres</span><span class="sxs-lookup"><span data-stu-id="aac71-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="aac71-116">displayLanguage : Chaîne</span><span class="sxs-lookup"><span data-stu-id="aac71-116">displayLanguage :String</span></span>

<span data-ttu-id="aac71-117">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="aac71-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="aac71-118">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="aac71-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="aac71-119">Type :</span><span class="sxs-lookup"><span data-stu-id="aac71-119">Type:</span></span>

*   <span data-ttu-id="aac71-120">Chaîne</span><span class="sxs-lookup"><span data-stu-id="aac71-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aac71-121">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="aac71-121">Requirements</span></span>

|<span data-ttu-id="aac71-122">Condition requise</span><span class="sxs-lookup"><span data-stu-id="aac71-122">Requirement</span></span>| <span data-ttu-id="aac71-123">Valeur</span><span class="sxs-lookup"><span data-stu-id="aac71-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="aac71-124">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="aac71-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aac71-125">1.0</span><span class="sxs-lookup"><span data-stu-id="aac71-125">1.0</span></span>|
|[<span data-ttu-id="aac71-126">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="aac71-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aac71-127">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="aac71-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aac71-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="aac71-128">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="aac71-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="aac71-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="aac71-130">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie, enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="aac71-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="aac71-131">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="aac71-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="aac71-132">Type :</span><span class="sxs-lookup"><span data-stu-id="aac71-132">Type:</span></span>

*   [<span data-ttu-id="aac71-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="aac71-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="aac71-134">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="aac71-134">Requirements</span></span>

|<span data-ttu-id="aac71-135">Condition requise</span><span class="sxs-lookup"><span data-stu-id="aac71-135">Requirement</span></span>| <span data-ttu-id="aac71-136">Valeur</span><span class="sxs-lookup"><span data-stu-id="aac71-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="aac71-137">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="aac71-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aac71-138">1.0</span><span class="sxs-lookup"><span data-stu-id="aac71-138">1.0</span></span>|
|[<span data-ttu-id="aac71-139">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="aac71-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aac71-140">Restreint</span><span class="sxs-lookup"><span data-stu-id="aac71-140">Restricted</span></span>|
|[<span data-ttu-id="aac71-141">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="aac71-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aac71-142">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="aac71-142">Compose or read</span></span>|