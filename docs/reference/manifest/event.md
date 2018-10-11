# <a name="event-element"></a><span data-ttu-id="02cc0-101">Élément Event</span><span class="sxs-lookup"><span data-stu-id="02cc0-101">Event element</span></span>

<span data-ttu-id="02cc0-102">Définit un gestionnaire d’événements dans un complément.</span><span class="sxs-lookup"><span data-stu-id="02cc0-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="02cc0-103">L’élément `Event` est actuellement uniquement pris en charge par Outlook sur le web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="02cc0-103">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="02cc0-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="02cc0-104">Attributes</span></span>

|  <span data-ttu-id="02cc0-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="02cc0-105">Attribute</span></span>  |  <span data-ttu-id="02cc0-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="02cc0-106">Required</span></span>  |  <span data-ttu-id="02cc0-107">Description</span><span class="sxs-lookup"><span data-stu-id="02cc0-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="02cc0-108">Type</span><span class="sxs-lookup"><span data-stu-id="02cc0-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="02cc0-109">Oui</span><span class="sxs-lookup"><span data-stu-id="02cc0-109">Yes</span></span>  | <span data-ttu-id="02cc0-110">Indique l’événement à gérer.</span><span class="sxs-lookup"><span data-stu-id="02cc0-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="02cc0-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="02cc0-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="02cc0-112">Oui</span><span class="sxs-lookup"><span data-stu-id="02cc0-112">Yes</span></span>  | <span data-ttu-id="02cc0-p101">Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="02cc0-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="02cc0-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="02cc0-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="02cc0-116">Oui</span><span class="sxs-lookup"><span data-stu-id="02cc0-116">Yes</span></span>  | <span data-ttu-id="02cc0-117">Indique le nom de la fonction du gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="02cc0-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="02cc0-118">Attribut Type</span><span class="sxs-lookup"><span data-stu-id="02cc0-118">Type attribute</span></span>

<span data-ttu-id="02cc0-p102">Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="02cc0-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="02cc0-122">Type d’événement</span><span class="sxs-lookup"><span data-stu-id="02cc0-122">Event type</span></span>  |  <span data-ttu-id="02cc0-123">Description</span><span class="sxs-lookup"><span data-stu-id="02cc0-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="02cc0-124">Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.</span><span class="sxs-lookup"><span data-stu-id="02cc0-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="02cc0-125">Attribut FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="02cc0-125">FunctionExecution attribute</span></span>

<span data-ttu-id="02cc0-126">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="02cc0-126">Required.</span></span> <span data-ttu-id="02cc0-127">DOIT être défini sur `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="02cc0-127">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="02cc0-128">Attribut FunctionName</span><span class="sxs-lookup"><span data-stu-id="02cc0-128">FunctionName attribute</span></span>

<span data-ttu-id="02cc0-p104">Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.</span><span class="sxs-lookup"><span data-stu-id="02cc0-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```