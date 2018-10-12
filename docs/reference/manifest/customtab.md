# <a name="customtab-element"></a><span data-ttu-id="7edd3-101">Élément CustomTab</span><span class="sxs-lookup"><span data-stu-id="7edd3-101">CustomTab element</span></span>

<span data-ttu-id="7edd3-p101">Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (soit  **Accueil**,  **Message**, ou  **Réunion**), ou d’un onglet personnalisé défini par le complément.</span><span class="sxs-lookup"><span data-stu-id="7edd3-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="7edd3-p102">Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="7edd3-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="7edd3-107">L’attribut **id** doit être unique au sein du manifeste.</span><span class="sxs-lookup"><span data-stu-id="7edd3-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7edd3-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7edd3-108">Child elements</span></span>

|  <span data-ttu-id="7edd3-109">Élément</span><span class="sxs-lookup"><span data-stu-id="7edd3-109">Element</span></span> |  <span data-ttu-id="7edd3-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7edd3-110">Required</span></span>  |  <span data-ttu-id="7edd3-111">Description</span><span class="sxs-lookup"><span data-stu-id="7edd3-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7edd3-112">Groupe</span><span class="sxs-lookup"><span data-stu-id="7edd3-112">Group</span></span>](group.md)      | <span data-ttu-id="7edd3-113">Oui</span><span class="sxs-lookup"><span data-stu-id="7edd3-113">Yes</span></span> |  <span data-ttu-id="7edd3-114">Définit un groupe de commandes.</span><span class="sxs-lookup"><span data-stu-id="7edd3-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="7edd3-115">Étiquette</span><span class="sxs-lookup"><span data-stu-id="7edd3-115">Label</span></span>](#label-tab)      | <span data-ttu-id="7edd3-116">Oui</span><span class="sxs-lookup"><span data-stu-id="7edd3-116">Yes</span></span> |  <span data-ttu-id="7edd3-117">Étiquette pour CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="7edd3-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="7edd3-118">Control</span><span class="sxs-lookup"><span data-stu-id="7edd3-118">Control</span></span>](control.md)    | <span data-ttu-id="7edd3-119">Oui</span><span class="sxs-lookup"><span data-stu-id="7edd3-119">Yes</span></span> |  <span data-ttu-id="7edd3-120">Ensemble d’un ou de plusieurs objets Control.</span><span class="sxs-lookup"><span data-stu-id="7edd3-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="7edd3-121">Groupe</span><span class="sxs-lookup"><span data-stu-id="7edd3-121">Group</span></span>

<span data-ttu-id="7edd3-p103">Obligatoire. Voir [Élément Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="7edd3-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="7edd3-124">Étiquette (onglet)</span><span class="sxs-lookup"><span data-stu-id="7edd3-124">Label (Tab)</span></span>

<span data-ttu-id="7edd3-p104">Obligatoire. Étiquette de l’onglet personnalisé. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="7edd3-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="7edd3-127">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="7edd3-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```