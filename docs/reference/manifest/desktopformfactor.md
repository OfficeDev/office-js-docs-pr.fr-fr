# <a name="desktopformfactor-element"></a><span data-ttu-id="4e8bb-101">Élément DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="4e8bb-101">DesktopFormFactor element</span></span>

<span data-ttu-id="4e8bb-p101">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme pour bureau inclut Office pour Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="4e8bb-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="4e8bb-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="4e8bb-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="4e8bb-107">Child elements</span></span>

| <span data-ttu-id="4e8bb-108">Élément</span><span class="sxs-lookup"><span data-stu-id="4e8bb-108">Element</span></span>                               | <span data-ttu-id="4e8bb-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="4e8bb-109">Required</span></span> | <span data-ttu-id="4e8bb-110">Description</span><span class="sxs-lookup"><span data-stu-id="4e8bb-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="4e8bb-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="4e8bb-111">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="4e8bb-112">Oui</span><span class="sxs-lookup"><span data-stu-id="4e8bb-112">Yes</span></span>      | <span data-ttu-id="4e8bb-113">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="4e8bb-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="4e8bb-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="4e8bb-114">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="4e8bb-115">Oui</span><span class="sxs-lookup"><span data-stu-id="4e8bb-115">Yes</span></span>      | <span data-ttu-id="4e8bb-116">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="4e8bb-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="4e8bb-117">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="4e8bb-118">Non</span><span class="sxs-lookup"><span data-stu-id="4e8bb-118">No</span></span>       | <span data-ttu-id="4e8bb-119">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="4e8bb-120">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="4e8bb-120">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="4e8bb-121">Non</span><span class="sxs-lookup"><span data-stu-id="4e8bb-121">No</span></span> | <span data-ttu-id="4e8bb-122">Définit si le complément Outlook est disponible dans les scénarios délégués et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-122">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="4e8bb-123">**Important** : cet élément est disponible uniquement dans l'ensemble des conditions requises de la préversion des compléments Outlook sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-123">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="4e8bb-124">Les compléments qui utilisent cet élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="4e8bb-124">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="4e8bb-125">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="4e8bb-125">DesktopFormFactor example</span></span>

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
