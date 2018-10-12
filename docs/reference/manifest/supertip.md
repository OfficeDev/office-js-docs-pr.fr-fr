# <a name="supertip"></a><span data-ttu-id="6a21e-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="6a21e-101">Supertip</span></span>

<span data-ttu-id="6a21e-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles [Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="6a21e-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6a21e-104">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="6a21e-104">Child elements</span></span>

|  <span data-ttu-id="6a21e-105">Élément</span><span class="sxs-lookup"><span data-stu-id="6a21e-105">Element</span></span> |  <span data-ttu-id="6a21e-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="6a21e-106">Required</span></span>  |  <span data-ttu-id="6a21e-107">Description</span><span class="sxs-lookup"><span data-stu-id="6a21e-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6a21e-108">Title</span><span class="sxs-lookup"><span data-stu-id="6a21e-108">Title</span></span>](#title)        | <span data-ttu-id="6a21e-109">Oui</span><span class="sxs-lookup"><span data-stu-id="6a21e-109">Yes</span></span> |   <span data-ttu-id="6a21e-110">Texte du supertip.</span><span class="sxs-lookup"><span data-stu-id="6a21e-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="6a21e-111">Description</span><span class="sxs-lookup"><span data-stu-id="6a21e-111">Description</span></span>](#description)  | <span data-ttu-id="6a21e-112">Oui</span><span class="sxs-lookup"><span data-stu-id="6a21e-112">Yes</span></span> |  <span data-ttu-id="6a21e-113">Description du supertip.</span><span class="sxs-lookup"><span data-stu-id="6a21e-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="6a21e-114">Titre</span><span class="sxs-lookup"><span data-stu-id="6a21e-114">Title</span></span>

<span data-ttu-id="6a21e-p102">Obligatoire. Texte du superTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="6a21e-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="6a21e-118">Description</span><span class="sxs-lookup"><span data-stu-id="6a21e-118">Description</span></span>

<span data-ttu-id="6a21e-p103">Obligatoire. Description du superTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="6a21e-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="6a21e-122">Exemple</span><span class="sxs-lookup"><span data-stu-id="6a21e-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
