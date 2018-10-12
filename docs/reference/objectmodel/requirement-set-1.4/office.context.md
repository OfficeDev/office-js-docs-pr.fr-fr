
# <a name="context"></a><span data-ttu-id="c46ac-101">context</span><span class="sxs-lookup"><span data-stu-id="c46ac-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="c46ac-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="c46ac-102">[Office](Office.md).context</span></span>

<span data-ttu-id="c46ac-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context dans l’interface API partagée](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="c46ac-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c46ac-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c46ac-105">Requirements</span></span>

|<span data-ttu-id="c46ac-106">Condition</span><span class="sxs-lookup"><span data-stu-id="c46ac-106">Requirement</span></span>| <span data-ttu-id="c46ac-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c46ac-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c46ac-108">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c46ac-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c46ac-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c46ac-109">1.0</span></span>|
|[<span data-ttu-id="c46ac-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c46ac-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c46ac-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c46ac-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="c46ac-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c46ac-112">Namespaces</span></span>

<span data-ttu-id="c46ac-113">[mailbox](office.context.mailbox.md) : Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="c46ac-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="c46ac-114">Membres</span><span class="sxs-lookup"><span data-stu-id="c46ac-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="c46ac-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="c46ac-115">displayLanguage :String</span></span>

<span data-ttu-id="c46ac-116">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c46ac-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c46ac-117">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c46ac-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c46ac-118">Type :</span><span class="sxs-lookup"><span data-stu-id="c46ac-118">Type:</span></span>

*   <span data-ttu-id="c46ac-119">String</span><span class="sxs-lookup"><span data-stu-id="c46ac-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c46ac-120">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="c46ac-120">Requirements</span></span>

|<span data-ttu-id="c46ac-121">Condition</span><span class="sxs-lookup"><span data-stu-id="c46ac-121">Requirement</span></span>| <span data-ttu-id="c46ac-122">Valeur</span><span class="sxs-lookup"><span data-stu-id="c46ac-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="c46ac-123">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c46ac-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c46ac-124">1.0</span><span class="sxs-lookup"><span data-stu-id="c46ac-124">1.0</span></span>|
|[<span data-ttu-id="c46ac-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c46ac-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c46ac-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c46ac-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c46ac-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="c46ac-127">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="c46ac-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="c46ac-128">officeTheme :Object</span></span>

<span data-ttu-id="c46ac-129">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="c46ac-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c46ac-130">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c46ac-130">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c46ac-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c46ac-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c46ac-133">Type :</span><span class="sxs-lookup"><span data-stu-id="c46ac-133">Type:</span></span>

*   <span data-ttu-id="c46ac-134">Objet</span><span class="sxs-lookup"><span data-stu-id="c46ac-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="c46ac-135">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c46ac-135">Properties:</span></span>

|<span data-ttu-id="c46ac-136">Nom</span><span class="sxs-lookup"><span data-stu-id="c46ac-136">Name</span></span>| <span data-ttu-id="c46ac-137">Type</span><span class="sxs-lookup"><span data-stu-id="c46ac-137">Type</span></span>| <span data-ttu-id="c46ac-138">Description</span><span class="sxs-lookup"><span data-stu-id="c46ac-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c46ac-139">String</span><span class="sxs-lookup"><span data-stu-id="c46ac-139">String</span></span>|<span data-ttu-id="c46ac-140">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c46ac-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c46ac-141">String</span><span class="sxs-lookup"><span data-stu-id="c46ac-141">String</span></span>|<span data-ttu-id="c46ac-142">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c46ac-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c46ac-143">String</span><span class="sxs-lookup"><span data-stu-id="c46ac-143">String</span></span>|<span data-ttu-id="c46ac-144">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c46ac-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c46ac-145">String</span><span class="sxs-lookup"><span data-stu-id="c46ac-145">String</span></span>|<span data-ttu-id="c46ac-146">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c46ac-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c46ac-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c46ac-147">Requirements</span></span>

|<span data-ttu-id="c46ac-148">Condition</span><span class="sxs-lookup"><span data-stu-id="c46ac-148">Requirement</span></span>| <span data-ttu-id="c46ac-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="c46ac-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="c46ac-150">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c46ac-150">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c46ac-151">1.3</span><span class="sxs-lookup"><span data-stu-id="c46ac-151">1.3</span></span>|
|[<span data-ttu-id="c46ac-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c46ac-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c46ac-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c46ac-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c46ac-154">Exemple</span><span class="sxs-lookup"><span data-stu-id="c46ac-154">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="c46ac-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="c46ac-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="c46ac-156">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c46ac-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c46ac-157">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="c46ac-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c46ac-158">Type :</span><span class="sxs-lookup"><span data-stu-id="c46ac-158">Type:</span></span>

*   [<span data-ttu-id="c46ac-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c46ac-159">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c46ac-160">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c46ac-160">Requirements</span></span>

|<span data-ttu-id="c46ac-161">Condition</span><span class="sxs-lookup"><span data-stu-id="c46ac-161">Requirement</span></span>| <span data-ttu-id="c46ac-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="c46ac-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="c46ac-163">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c46ac-163">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c46ac-164">1.0</span><span class="sxs-lookup"><span data-stu-id="c46ac-164">1.0</span></span>|
|[<span data-ttu-id="c46ac-165">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c46ac-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c46ac-166">Restreint</span><span class="sxs-lookup"><span data-stu-id="c46ac-166">Restricted</span></span>|
|[<span data-ttu-id="c46ac-167">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c46ac-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c46ac-168">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c46ac-168">Compose or read</span></span>|