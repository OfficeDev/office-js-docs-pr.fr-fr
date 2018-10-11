# <a name="desktopformfactor-element"></a><span data-ttu-id="3bd72-101">Élément DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="3bd72-101">DesktopFormFactor element</span></span>

<span data-ttu-id="3bd72-p101">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme pour bureau inclut Office pour Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="3bd72-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="3bd72-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="3bd72-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3bd72-107">L'élément SupportsSharedFolders est disponible uniquement dans le complément Outlook Preview Requirement Set sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="3bd72-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="3bd72-108">Les compléments qui utilisent cet élément ne sont pas autorisés dans l’Office Store ou le Déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="3bd72-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3bd72-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="3bd72-109">Child elements</span></span>

| <span data-ttu-id="3bd72-110">Élément</span><span class="sxs-lookup"><span data-stu-id="3bd72-110">Element</span></span>                               | <span data-ttu-id="3bd72-111">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3bd72-111">Required</span></span> | <span data-ttu-id="3bd72-112">Description</span><span class="sxs-lookup"><span data-stu-id="3bd72-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="3bd72-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="3bd72-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="3bd72-114">Oui</span><span class="sxs-lookup"><span data-stu-id="3bd72-114">Yes</span></span>      | <span data-ttu-id="3bd72-115">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="3bd72-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="3bd72-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="3bd72-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="3bd72-117">Oui</span><span class="sxs-lookup"><span data-stu-id="3bd72-117">Yes</span></span>      | <span data-ttu-id="3bd72-118">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3bd72-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="3bd72-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="3bd72-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="3bd72-120">Non</span><span class="sxs-lookup"><span data-stu-id="3bd72-120">No</span></span>       | <span data-ttu-id="3bd72-121">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="3bd72-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="3bd72-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="3bd72-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="3bd72-123">Non</span><span class="sxs-lookup"><span data-stu-id="3bd72-123">No</span></span>       | <span data-ttu-id="3bd72-124">Définit si le complément Outlook est disponible dans les scénarios délégués et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="3bd72-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="3bd72-125">Aperçu de l'ensemble d'exigences.</span><span class="sxs-lookup"><span data-stu-id="3bd72-125">Outlook add-in API Preview requirement set</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="3bd72-126">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="3bd72-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
