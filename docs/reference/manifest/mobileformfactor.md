# <a name="mobileformfactor-element"></a><span data-ttu-id="1e464-101">Élément MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="1e464-101">MobileFormFactor element</span></span>

<span data-ttu-id="1e464-p101">Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="1e464-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="1e464-p102">Chaque définition **MobileFormFactor** contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="1e464-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="1e464-p103">L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="1e464-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1e464-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1e464-108">Child elements</span></span>

| <span data-ttu-id="1e464-109">Élément</span><span class="sxs-lookup"><span data-stu-id="1e464-109">Element</span></span>                               | <span data-ttu-id="1e464-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1e464-110">Required</span></span> | <span data-ttu-id="1e464-111">Description</span><span class="sxs-lookup"><span data-stu-id="1e464-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="1e464-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="1e464-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="1e464-113">Oui</span><span class="sxs-lookup"><span data-stu-id="1e464-113">Yes</span></span>      | <span data-ttu-id="1e464-114">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="1e464-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="1e464-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="1e464-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="1e464-116">Oui</span><span class="sxs-lookup"><span data-stu-id="1e464-116">Yes</span></span>      | <span data-ttu-id="1e464-117">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1e464-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="1e464-118">Exemple MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="1e464-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
