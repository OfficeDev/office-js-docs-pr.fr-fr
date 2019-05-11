---
title: Élément Action dans le fichier manifeste
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 58dcbae57ea2c0e55c9e7708b122484b99e956fe
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952403"
---
# <a name="action-element"></a><span data-ttu-id="1d629-102">Action, élément</span><span class="sxs-lookup"><span data-stu-id="1d629-102">Action element</span></span>

<span data-ttu-id="1d629-103">Indique l’action à réaliser lorsque l’utilisateur sélectionne des contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="1d629-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="1d629-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="1d629-104">Attributes</span></span>

|  <span data-ttu-id="1d629-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="1d629-105">Attribute</span></span>  |  <span data-ttu-id="1d629-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1d629-106">Required</span></span>  |  <span data-ttu-id="1d629-107">Description</span><span class="sxs-lookup"><span data-stu-id="1d629-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1d629-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1d629-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="1d629-109">Oui</span><span class="sxs-lookup"><span data-stu-id="1d629-109">Yes</span></span>  | <span data-ttu-id="1d629-110">Type d’action à effectuer</span><span class="sxs-lookup"><span data-stu-id="1d629-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="1d629-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1d629-111">Child elements</span></span>

|  <span data-ttu-id="1d629-112">Élément</span><span class="sxs-lookup"><span data-stu-id="1d629-112">Element</span></span> |  <span data-ttu-id="1d629-113">Description</span><span class="sxs-lookup"><span data-stu-id="1d629-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1d629-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="1d629-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="1d629-115">Spécifie le nom de la fonction à exécuter.</span><span class="sxs-lookup"><span data-stu-id="1d629-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="1d629-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1d629-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="1d629-117">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="1d629-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="1d629-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="1d629-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="1d629-119">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="1d629-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="1d629-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="1d629-120"> [Title](#title)</span></span> | <span data-ttu-id="1d629-121">Indique le titre personnalisé du volet Office.</span><span class="sxs-lookup"><span data-stu-id="1d629-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="1d629-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="1d629-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="1d629-123">Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.</span><span class="sxs-lookup"><span data-stu-id="1d629-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="1d629-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1d629-124">xsi:type</span></span>

<span data-ttu-id="1d629-p101">Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="1d629-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="1d629-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="1d629-127">FunctionName</span></span>

<span data-ttu-id="1d629-p102">Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="1d629-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="1d629-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1d629-131">SourceLocation</span></span>

<span data-ttu-id="1d629-p103">Élément obligatoire lorsque  **xsi:type** est « ShowTaskpane ». Indique l’emplacement du fichier source pour cette action. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Urls** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="1d629-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="1d629-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="1d629-135">TaskpaneId</span></span>

<span data-ttu-id="1d629-136">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="1d629-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1d629-137">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="1d629-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="1d629-138">Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun.</span><span class="sxs-lookup"><span data-stu-id="1d629-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="1d629-139">Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet.</span><span class="sxs-lookup"><span data-stu-id="1d629-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="1d629-140">Lorsque les utilisateurs choisissent des commandes qui partagent le même attribut **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».</span><span class="sxs-lookup"><span data-stu-id="1d629-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="1d629-141">Cet élément n’est pas pris en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d629-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="1d629-142">L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="1d629-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="1d629-p105">Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="1d629-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="1d629-145">Titre</span><span class="sxs-lookup"><span data-stu-id="1d629-145">Title</span></span>

<span data-ttu-id="1d629-146">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="1d629-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1d629-147">Indique le titre personnalisé du volet Office pour cette action.</span><span class="sxs-lookup"><span data-stu-id="1d629-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="1d629-148">Les exemples ci-dessous illustrent deux différentes actions qui utilisent l’élément **title**.</span><span class="sxs-lookup"><span data-stu-id="1d629-148">The following examples show two different actions that use the **Title** element.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="1d629-149">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="1d629-149">SupportsPinning</span></span>

<span data-ttu-id="1d629-150">Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="1d629-150">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1d629-151">Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="1d629-151">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="1d629-152">Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office.</span><span class="sxs-lookup"><span data-stu-id="1d629-152">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="1d629-153">L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée.</span><span class="sxs-lookup"><span data-stu-id="1d629-153">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="1d629-154">Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="1d629-154">For more information, see [Implement a pinnable task pane in Outlook](/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="1d629-155">Supportspinning n’est est actuellement uniquement pris en charge par Outlook 2016 sur Windows (Build 7628,1000 ou version ultérieure) et Outlook 2016 pour Mac (Build 16.13.503 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="1d629-155">SupportsPinning is currently only supported by Outlook 2016 on Windows (build 7628.1000 or later) and Outlook 2016 for Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
