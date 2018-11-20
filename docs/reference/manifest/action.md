# <a name="action-element"></a><span data-ttu-id="8c139-101">Action, élément</span><span class="sxs-lookup"><span data-stu-id="8c139-101">Action element</span></span>

<span data-ttu-id="8c139-102">Indique l’action à réaliser lorsque l’utilisateur sélectionne des contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="8c139-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="8c139-103">Attributs</span><span class="sxs-lookup"><span data-stu-id="8c139-103">Attributes</span></span>

|  <span data-ttu-id="8c139-104">Attribut</span><span class="sxs-lookup"><span data-stu-id="8c139-104">Attribute</span></span>  |  <span data-ttu-id="8c139-105">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8c139-105">Required</span></span>  |  <span data-ttu-id="8c139-106">Description</span><span class="sxs-lookup"><span data-stu-id="8c139-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8c139-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8c139-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="8c139-108">Oui</span><span class="sxs-lookup"><span data-stu-id="8c139-108">Yes</span></span>  | <span data-ttu-id="8c139-109">Type d’action à effectuer</span><span class="sxs-lookup"><span data-stu-id="8c139-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="8c139-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="8c139-110">Child elements</span></span>

|  <span data-ttu-id="8c139-111">Élément</span><span class="sxs-lookup"><span data-stu-id="8c139-111">Element</span></span> |  <span data-ttu-id="8c139-112">Description</span><span class="sxs-lookup"><span data-stu-id="8c139-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="8c139-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8c139-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="8c139-114">Spécifie le nom de la fonction à exécuter.</span><span class="sxs-lookup"><span data-stu-id="8c139-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="8c139-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8c139-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="8c139-116">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="8c139-116">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="8c139-117"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="8c139-117"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="8c139-118">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="8c139-118">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="8c139-119"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="8c139-119"> [Title](#title)</span></span> | <span data-ttu-id="8c139-120">Indique le titre personnalisé du volet Office.</span><span class="sxs-lookup"><span data-stu-id="8c139-120">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="8c139-121"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="8c139-121"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="8c139-122">Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.</span><span class="sxs-lookup"><span data-stu-id="8c139-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="8c139-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8c139-123">xsi:type</span></span>

<span data-ttu-id="8c139-p101">Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="8c139-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="8c139-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8c139-126">FunctionName</span></span>

<span data-ttu-id="8c139-p102">Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="8c139-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="8c139-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8c139-130">SourceLocation</span></span>

<span data-ttu-id="8c139-p103">Élément obligatoire lorsque  **xsi:type** est « ShowTaskpane ». Indique l’emplacement du fichier source pour cette action. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Urls** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8c139-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="8c139-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="8c139-134">TaskpaneId</span></span>

<span data-ttu-id="8c139-135">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="8c139-135">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c139-136">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="8c139-136">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="8c139-137">Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre attribut **TaskpaneId** si vous souhaitez un volet indépendant pour chacun.</span><span class="sxs-lookup"><span data-stu-id="8c139-137">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="8c139-138">Utilisez le même attribut **TaskpaneId** pour différentes actions qui partagent le même volet.</span><span class="sxs-lookup"><span data-stu-id="8c139-138">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="8c139-139">Lorsque les utilisateurs choisissent des commandes qui partagent le même attribut **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».</span><span class="sxs-lookup"><span data-stu-id="8c139-139">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="8c139-140">Cet élément n’est pas pris en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="8c139-140">Note: This element is not supported in Outlook.</span></span>

<span data-ttu-id="8c139-141">L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="8c139-141">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="8c139-p105">Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="8c139-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="8c139-144">Titre</span><span class="sxs-lookup"><span data-stu-id="8c139-144">Title</span></span>

<span data-ttu-id="8c139-145">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="8c139-145">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c139-146">Indique le titre personnalisé du volet Office pour cette action.</span><span class="sxs-lookup"><span data-stu-id="8c139-146">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="8c139-147">Les exemples ci-dessous illustrent deux différentes actions qui utilisent l’élément **title**.</span><span class="sxs-lookup"><span data-stu-id="8c139-147">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a><span data-ttu-id="8c139-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="8c139-148">SupportsPinning</span></span>

<span data-ttu-id="8c139-149">Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="8c139-149">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="8c139-150">Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="8c139-150">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="8c139-151">Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office.</span><span class="sxs-lookup"><span data-stu-id="8c139-151">Include this element with a value of `true` to support taskpane pinning.</span></span> <span data-ttu-id="8c139-152">L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée.</span><span class="sxs-lookup"><span data-stu-id="8c139-152">The user will be able to "pin" the taskpane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="8c139-153">Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="8c139-153">For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="8c139-154">SupportsPinning n’est actuellement pris en charge que par Outlook 2016 pour Windows (build 7628.1000 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="8c139-154">Note: SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
