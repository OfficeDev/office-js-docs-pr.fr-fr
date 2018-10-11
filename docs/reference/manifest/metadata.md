# <a name="metadata-element"></a><span data-ttu-id="6f6fa-101">Élément Metadata</span><span class="sxs-lookup"><span data-stu-id="6f6fa-101">MetaData element</span></span>

<span data-ttu-id="6f6fa-102">Définit les paramètres de métadonnées utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="6f6fa-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="6f6fa-103">Attributs</span><span class="sxs-lookup"><span data-stu-id="6f6fa-103">Attributes</span></span>

<span data-ttu-id="6f6fa-104">Aucun</span><span class="sxs-lookup"><span data-stu-id="6f6fa-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="6f6fa-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="6f6fa-105">Child elements</span></span>

|  <span data-ttu-id="6f6fa-106">Élément</span><span class="sxs-lookup"><span data-stu-id="6f6fa-106">Element</span></span>  |  <span data-ttu-id="6f6fa-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="6f6fa-107">Required</span></span>  |  <span data-ttu-id="6f6fa-108">Description</span><span class="sxs-lookup"><span data-stu-id="6f6fa-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6f6fa-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6f6fa-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="6f6fa-110">Oui</span><span class="sxs-lookup"><span data-stu-id="6f6fa-110">Yes</span></span>  | <span data-ttu-id="6f6fa-111">Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="6f6fa-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="6f6fa-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="6f6fa-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
