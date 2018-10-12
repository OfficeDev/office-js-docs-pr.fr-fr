# <a name="group-element"></a><span data-ttu-id="b30b4-101">Élément Group</span><span class="sxs-lookup"><span data-stu-id="b30b4-101">Group element</span></span>

<span data-ttu-id="b30b4-p101">Définit un groupe de contrôles d’interface utilisateur dans un onglet. Sur les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b30b4-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="b30b4-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="b30b4-105">Attributes</span></span>

|  <span data-ttu-id="b30b4-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="b30b4-106">Attribute</span></span>  |  <span data-ttu-id="b30b4-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b30b4-107">Required</span></span>  |  <span data-ttu-id="b30b4-108">Description</span><span class="sxs-lookup"><span data-stu-id="b30b4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b30b4-109">id</span><span class="sxs-lookup"><span data-stu-id="b30b4-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="b30b4-110">Oui</span><span class="sxs-lookup"><span data-stu-id="b30b4-110">Yes</span></span>  | <span data-ttu-id="b30b4-111">ID unique du groupe.</span><span class="sxs-lookup"><span data-stu-id="b30b4-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="b30b4-112">Attribut id</span><span class="sxs-lookup"><span data-stu-id="b30b4-112">id attribute</span></span>

<span data-ttu-id="b30b4-p102">Obligatoire. Identificateur unique du groupe. Il s’agit d’une chaîne avec un maximum de 125 caractères. Il doit être unique au sein du manifeste pour que le groupe s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="b30b4-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b30b4-117">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b30b4-117">Child elements</span></span>
|  <span data-ttu-id="b30b4-118">Élément</span><span class="sxs-lookup"><span data-stu-id="b30b4-118">Element</span></span> |  <span data-ttu-id="b30b4-119">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b30b4-119">Required</span></span>  |  <span data-ttu-id="b30b4-120">Description</span><span class="sxs-lookup"><span data-stu-id="b30b4-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b30b4-121">Label</span><span class="sxs-lookup"><span data-stu-id="b30b4-121">Label</span></span>](#label)      | <span data-ttu-id="b30b4-122">Oui</span><span class="sxs-lookup"><span data-stu-id="b30b4-122">Yes</span></span> |  <span data-ttu-id="b30b4-123">Étiquette pour le CustomTab ou un group.</span><span class="sxs-lookup"><span data-stu-id="b30b4-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="b30b4-124">Control</span><span class="sxs-lookup"><span data-stu-id="b30b4-124">Control</span></span>](#control)    | <span data-ttu-id="b30b4-125">Oui</span><span class="sxs-lookup"><span data-stu-id="b30b4-125">Yes</span></span> |  <span data-ttu-id="b30b4-126">Ensemble d’un ou de plusieurs objets Control.</span><span class="sxs-lookup"><span data-stu-id="b30b4-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="b30b4-127">Label</span><span class="sxs-lookup"><span data-stu-id="b30b4-127">Label</span></span> 

<span data-ttu-id="b30b4-p103">Obligatoire. Étiquette du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="b30b4-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="b30b4-131">Control</span><span class="sxs-lookup"><span data-stu-id="b30b4-131">Control</span></span>
<span data-ttu-id="b30b4-132">Un groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="b30b4-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```