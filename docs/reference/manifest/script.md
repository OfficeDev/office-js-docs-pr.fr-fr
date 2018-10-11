# <a name="script-element"></a><span data-ttu-id="43cde-101">Élément Script</span><span class="sxs-lookup"><span data-stu-id="43cde-101">Script element</span></span>

<span data-ttu-id="43cde-102">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="43cde-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="43cde-103">Attributs</span><span class="sxs-lookup"><span data-stu-id="43cde-103">Attributes</span></span>

<span data-ttu-id="43cde-104">Aucun</span><span class="sxs-lookup"><span data-stu-id="43cde-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="43cde-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="43cde-105">Child elements</span></span>

|<span data-ttu-id="43cde-106">Éléments</span><span class="sxs-lookup"><span data-stu-id="43cde-106">Elements</span></span>  |  <span data-ttu-id="43cde-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="43cde-107">Required</span></span>  |  <span data-ttu-id="43cde-108">Description</span><span class="sxs-lookup"><span data-stu-id="43cde-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="43cde-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="43cde-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="43cde-110">Oui</span><span class="sxs-lookup"><span data-stu-id="43cde-110">Yes</span></span>  | <span data-ttu-id="43cde-111">Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="43cde-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="43cde-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="43cde-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
