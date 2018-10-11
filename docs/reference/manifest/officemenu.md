# <a name="officemenu-element"></a><span data-ttu-id="ca561-101">Élément OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="ca561-101">OfficeMenu element</span></span>

<span data-ttu-id="ca561-p101">Définit un ensemble d’options à ajouter au menu contextuel Office. S’applique aux compléments Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="ca561-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="ca561-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="ca561-104">Attributes</span></span>

| <span data-ttu-id="ca561-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="ca561-105">Attribute</span></span>            | <span data-ttu-id="ca561-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="ca561-106">Required</span></span> | <span data-ttu-id="ca561-107">Description</span><span class="sxs-lookup"><span data-stu-id="ca561-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="ca561-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ca561-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="ca561-109">Oui</span><span class="sxs-lookup"><span data-stu-id="ca561-109">Yes</span></span>      | <span data-ttu-id="ca561-110">Type d’OfficeMenu défini.</span><span class="sxs-lookup"><span data-stu-id="ca561-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="ca561-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="ca561-111">Child elements</span></span>

|  <span data-ttu-id="ca561-112">Élément</span><span class="sxs-lookup"><span data-stu-id="ca561-112">Element</span></span> |  <span data-ttu-id="ca561-113">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="ca561-113">Required</span></span>  |  <span data-ttu-id="ca561-114">Description</span><span class="sxs-lookup"><span data-stu-id="ca561-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ca561-115">Control</span><span class="sxs-lookup"><span data-stu-id="ca561-115">Control</span></span>](#control)    | <span data-ttu-id="ca561-116">Oui</span><span class="sxs-lookup"><span data-stu-id="ca561-116">Yes</span></span> |  <span data-ttu-id="ca561-117">Ensemble d’un ou de plusieurs objets Control</span><span class="sxs-lookup"><span data-stu-id="ca561-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="ca561-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ca561-118">xsi:type</span></span>

<span data-ttu-id="ca561-119">Indique un menu prédéfini de l’application cliente Office sur lequel ajouter ce complément Office.</span><span class="sxs-lookup"><span data-stu-id="ca561-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="ca561-p102">`ContextMenuText` -  Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur ouvre le menu contextuel (clique dessus avec le bouton droit de la souris) du texte sélectionné. S’applique à Word, Excel, PowerPoint et OneNote.</span><span class="sxs-lookup"><span data-stu-id="ca561-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="ca561-p103">`ContextMenuCell` -  Affiche l’élément dans le menu contextuel lorsque l’utilisateur ouvre le menu contextuel (clique avec le bouton droit de la souris) dans une cellule de la feuille de calcul. S’applique à Excel.</span><span class="sxs-lookup"><span data-stu-id="ca561-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="ca561-124">Control</span><span class="sxs-lookup"><span data-stu-id="ca561-124">Control</span></span>

<span data-ttu-id="ca561-125">Chaque élément **OfficeMenu** requiert un ou plusieurs contrôles [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="ca561-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="ca561-126">Exemple</span><span class="sxs-lookup"><span data-stu-id="ca561-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
