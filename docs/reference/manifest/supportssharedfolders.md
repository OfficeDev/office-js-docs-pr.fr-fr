# <a name="supportssharedfolders-element"></a><span data-ttu-id="6de61-101">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="6de61-101">SupportsSharedFolders element</span></span>

<span data-ttu-id="6de61-102">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="6de61-102">Defines whether the Outlook add-in is available in delegate scenarios and is set to false by default.</span></span> <span data-ttu-id="6de61-103">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="6de61-103">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="6de61-104">Il est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="6de61-104">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6de61-105">Cet élément est disponible uniquement dans [l’ensemble des conditions requises de la préversion des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)  sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="6de61-105">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="6de61-106">Les compléments qui utilisent cet élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="6de61-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="6de61-107">Vous trouverez ci-dessous un exemple de l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="6de61-107">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
