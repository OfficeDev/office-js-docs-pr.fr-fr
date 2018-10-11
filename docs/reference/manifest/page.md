# <a name="page-element"></a><span data-ttu-id="e0657-101">Élément Page</span><span class="sxs-lookup"><span data-stu-id="e0657-101">Page element</span></span>

<span data-ttu-id="e0657-102">Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="e0657-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e0657-103">Attributs</span><span class="sxs-lookup"><span data-stu-id="e0657-103">Attributes</span></span>

<span data-ttu-id="e0657-104">Aucun</span><span class="sxs-lookup"><span data-stu-id="e0657-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="e0657-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e0657-105">Child elements</span></span>

|  <span data-ttu-id="e0657-106">Élément</span><span class="sxs-lookup"><span data-stu-id="e0657-106">Element</span></span>  |  <span data-ttu-id="e0657-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e0657-107">Required</span></span>  |  <span data-ttu-id="e0657-108">Description</span><span class="sxs-lookup"><span data-stu-id="e0657-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e0657-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e0657-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="e0657-110">Oui</span><span class="sxs-lookup"><span data-stu-id="e0657-110">Yes</span></span>  | <span data-ttu-id="e0657-111">Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="e0657-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="e0657-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="e0657-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
