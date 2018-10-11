# <a name="allformfactors-element"></a><span data-ttu-id="94db3-101">Élément AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="94db3-101">AllFormFactors element</span></span>

<span data-ttu-id="94db3-102">Spécifie les paramètres d’un complément pour tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="94db3-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="94db3-103">Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="94db3-103">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="94db3-104">**AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="94db3-104">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="94db3-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="94db3-105">Child elements</span></span>

|  <span data-ttu-id="94db3-106">Élément</span><span class="sxs-lookup"><span data-stu-id="94db3-106">Element</span></span> |  <span data-ttu-id="94db3-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="94db3-107">Required</span></span>  |  <span data-ttu-id="94db3-108">Description</span><span class="sxs-lookup"><span data-stu-id="94db3-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="94db3-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="94db3-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="94db3-110">Oui</span><span class="sxs-lookup"><span data-stu-id="94db3-110">Yes</span></span> |  <span data-ttu-id="94db3-111">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="94db3-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="94db3-112">Exemple AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="94db3-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
