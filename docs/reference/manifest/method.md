# <a name="method-element"></a><span data-ttu-id="57c7a-101">Élément Method</span><span class="sxs-lookup"><span data-stu-id="57c7a-101">Method element</span></span>

<span data-ttu-id="57c7a-102">Spécifie une méthode individuelle de l’interface API JavaScript pour Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="57c7a-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="57c7a-103">**Type de complément :** contenu, volet Office</span><span class="sxs-lookup"><span data-stu-id="57c7a-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="57c7a-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="57c7a-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="57c7a-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="57c7a-105">Contained in:</span></span>

[<span data-ttu-id="57c7a-106">Méthodes</span><span class="sxs-lookup"><span data-stu-id="57c7a-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="57c7a-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="57c7a-107">Attributes</span></span>

|<span data-ttu-id="57c7a-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="57c7a-108">**Attribute**</span></span>|<span data-ttu-id="57c7a-109">**Type**</span><span class="sxs-lookup"><span data-stu-id="57c7a-109">**Type**</span></span>|<span data-ttu-id="57c7a-110">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="57c7a-110">**Required**</span></span>|<span data-ttu-id="57c7a-111">**Description**</span><span class="sxs-lookup"><span data-stu-id="57c7a-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="57c7a-112">Name</span><span class="sxs-lookup"><span data-stu-id="57c7a-112">Name</span></span>|<span data-ttu-id="57c7a-113">string</span><span class="sxs-lookup"><span data-stu-id="57c7a-113">string</span></span>|<span data-ttu-id="57c7a-114">obligatoire</span><span class="sxs-lookup"><span data-stu-id="57c7a-114">required</span></span>|<span data-ttu-id="57c7a-p101">Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="57c7a-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="57c7a-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="57c7a-117">Remarks</span></span>

<span data-ttu-id="57c7a-118">Les éléments **Methods** et **Method** ne sont pas pris en charge dans les compléments de courrier. Pour plus d’informations sur les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="57c7a-118">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="57c7a-119">Comme il n’existe aucun moyen pour spécifier la version minimale de condition requise pour les différentes méthodes, pour vous assurer qu’une méthode est disponible à l’exécution, vous devez également utiliser une instruction **if** lors de l’appel de cette méthode dans le script de votre complément.</span><span class="sxs-lookup"><span data-stu-id="57c7a-119">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="57c7a-120">Pour plus d’informations, voir [Présentation de l’interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="57c7a-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

