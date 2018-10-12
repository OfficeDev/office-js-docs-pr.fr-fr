# <a name="host-element"></a><span data-ttu-id="5626c-101">Élément Host</span><span class="sxs-lookup"><span data-stu-id="5626c-101">Host element</span></span>

<span data-ttu-id="5626c-102">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="5626c-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="5626c-103">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="5626c-103">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="5626c-104">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="5626c-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="5626c-105">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="5626c-105">Basic manifest</span></span>

<span data-ttu-id="5626c-106">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="5626c-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="5626c-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="5626c-107">Attributes</span></span>

| <span data-ttu-id="5626c-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="5626c-108">Attribute</span></span>     | <span data-ttu-id="5626c-109">Type</span><span class="sxs-lookup"><span data-stu-id="5626c-109">Type</span></span>   | <span data-ttu-id="5626c-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5626c-110">Required</span></span> | <span data-ttu-id="5626c-111">Description</span><span class="sxs-lookup"><span data-stu-id="5626c-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="5626c-112">Name</span><span class="sxs-lookup"><span data-stu-id="5626c-112">Name</span></span>](#name) | <span data-ttu-id="5626c-113">string</span><span class="sxs-lookup"><span data-stu-id="5626c-113">string</span></span> | <span data-ttu-id="5626c-114">obligatoire</span><span class="sxs-lookup"><span data-stu-id="5626c-114">required</span></span> | <span data-ttu-id="5626c-115">Nom du type d’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="5626c-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="5626c-116">Name</span><span class="sxs-lookup"><span data-stu-id="5626c-116">Name</span></span>
<span data-ttu-id="5626c-p102">Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="5626c-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="5626c-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="5626c-119">`Document` (Word)</span></span>
- <span data-ttu-id="5626c-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="5626c-120">`Database` (Access)</span></span>
- <span data-ttu-id="5626c-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="5626c-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="5626c-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="5626c-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="5626c-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="5626c-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="5626c-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="5626c-124">`Project` (Project)</span></span>
- <span data-ttu-id="5626c-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="5626c-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="5626c-126">Exemple</span><span class="sxs-lookup"><span data-stu-id="5626c-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="5626c-127">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5626c-127">VersionOverrides node</span></span>
<span data-ttu-id="5626c-128">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="5626c-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="5626c-129">Attributs</span><span class="sxs-lookup"><span data-stu-id="5626c-129">Attributes</span></span>

|  <span data-ttu-id="5626c-130">Attribut</span><span class="sxs-lookup"><span data-stu-id="5626c-130">Attribute</span></span>  |  <span data-ttu-id="5626c-131">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5626c-131">Required</span></span>  |  <span data-ttu-id="5626c-132">Description</span><span class="sxs-lookup"><span data-stu-id="5626c-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5626c-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="5626c-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="5626c-134">Oui</span><span class="sxs-lookup"><span data-stu-id="5626c-134">Yes</span></span>  | <span data-ttu-id="5626c-135">Décrit l’hôte d’Office dans lequel ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="5626c-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="5626c-136">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="5626c-136">Child elements</span></span>

|  <span data-ttu-id="5626c-137">Élément</span><span class="sxs-lookup"><span data-stu-id="5626c-137">Element</span></span> |  <span data-ttu-id="5626c-138">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5626c-138">Required</span></span>  |  <span data-ttu-id="5626c-139">Description</span><span class="sxs-lookup"><span data-stu-id="5626c-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5626c-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="5626c-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="5626c-141">Oui</span><span class="sxs-lookup"><span data-stu-id="5626c-141">Yes</span></span>   |  <span data-ttu-id="5626c-142">Définit les paramètres pour le facteur de forme bureau.</span><span class="sxs-lookup"><span data-stu-id="5626c-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="5626c-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="5626c-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="5626c-144">Non</span><span class="sxs-lookup"><span data-stu-id="5626c-144">No</span></span>   |  <span data-ttu-id="5626c-p103">Définit les paramètres pour le facteur de forme mobile. **Remarque :** cet élément est uniquement pris en charge dans Outlook pour iOS.</span><span class="sxs-lookup"><span data-stu-id="5626c-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="5626c-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="5626c-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="5626c-148">Non</span><span class="sxs-lookup"><span data-stu-id="5626c-148">No</span></span>   |  <span data-ttu-id="5626c-149">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="5626c-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="5626c-150">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5626c-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="5626c-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="5626c-151">xsi:type</span></span>

<span data-ttu-id="5626c-152">Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’applique également les paramètres contenus.</span><span class="sxs-lookup"><span data-stu-id="5626c-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="5626c-153">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="5626c-153">The value must be one of the following:</span></span>

- <span data-ttu-id="5626c-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="5626c-154">`Document` (Word)</span></span>
- <span data-ttu-id="5626c-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="5626c-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="5626c-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="5626c-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="5626c-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="5626c-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="5626c-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="5626c-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="5626c-159">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="5626c-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
