---
title: Élément Action dans le fichier manifeste
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b05da08f4995c7d8f7270e7fba6f416c9903b066
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324890"
---
# <a name="action-element"></a><span data-ttu-id="9b644-102">Action, élément</span><span class="sxs-lookup"><span data-stu-id="9b644-102">Action element</span></span>

<span data-ttu-id="9b644-103">Indique l’action à réaliser lorsque l’utilisateur sélectionne des contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="9b644-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="9b644-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b644-104">Attributes</span></span>

|  <span data-ttu-id="9b644-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="9b644-105">Attribute</span></span>  |  <span data-ttu-id="9b644-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9b644-106">Required</span></span>  |  <span data-ttu-id="9b644-107">Description</span><span class="sxs-lookup"><span data-stu-id="9b644-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9b644-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b644-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9b644-109">Oui</span><span class="sxs-lookup"><span data-stu-id="9b644-109">Yes</span></span>  | <span data-ttu-id="9b644-110">Type d’action à effectuer</span><span class="sxs-lookup"><span data-stu-id="9b644-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="9b644-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9b644-111">Child elements</span></span>

|  <span data-ttu-id="9b644-112">Élément</span><span class="sxs-lookup"><span data-stu-id="9b644-112">Element</span></span> |  <span data-ttu-id="9b644-113">Description</span><span class="sxs-lookup"><span data-stu-id="9b644-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9b644-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9b644-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="9b644-115">Spécifie le nom de la fonction à exécuter.</span><span class="sxs-lookup"><span data-stu-id="9b644-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="9b644-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9b644-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="9b644-117">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="9b644-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="9b644-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="9b644-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="9b644-119">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="9b644-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="9b644-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="9b644-120"> [Title](#title)</span></span> | <span data-ttu-id="9b644-121">Indique le titre personnalisé du volet Office.</span><span class="sxs-lookup"><span data-stu-id="9b644-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="9b644-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="9b644-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="9b644-123">Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.</span><span class="sxs-lookup"><span data-stu-id="9b644-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="9b644-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b644-124">xsi:type</span></span>

<span data-ttu-id="9b644-p101">Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="9b644-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="9b644-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9b644-127">FunctionName</span></span>

<span data-ttu-id="9b644-p102">Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="9b644-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="9b644-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9b644-131">SourceLocation</span></span>

<span data-ttu-id="9b644-132">Élément obligatoire lorsque **xsi : type** est « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="9b644-132">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9b644-133">Indique l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="9b644-133">Specifies the source file location for this action.</span></span> <span data-ttu-id="9b644-134">L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Urls** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="9b644-134">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="9b644-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="9b644-135">TaskpaneId</span></span>

<span data-ttu-id="9b644-136">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="9b644-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9b644-137">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="9b644-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="9b644-138">Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun.</span><span class="sxs-lookup"><span data-stu-id="9b644-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="9b644-139">Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet.</span><span class="sxs-lookup"><span data-stu-id="9b644-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="9b644-140">Lorsque les utilisateurs choisissent des commandes qui partagent le même attribut **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».</span><span class="sxs-lookup"><span data-stu-id="9b644-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="9b644-141">Cet élément n’est pas pris en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="9b644-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="9b644-142">L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="9b644-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="9b644-p105">Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="9b644-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="9b644-145">Titre</span><span class="sxs-lookup"><span data-stu-id="9b644-145">Title</span></span>

<span data-ttu-id="9b644-146">Élément facultatif quand  **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="9b644-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9b644-147">Indique le titre personnalisé du volet Office pour cette action.</span><span class="sxs-lookup"><span data-stu-id="9b644-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="9b644-148">L’exemple suivant montre une action qui utilise l’élément **title** .</span><span class="sxs-lookup"><span data-stu-id="9b644-148">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="9b644-149">Notez que vous n’affectez pas directement le **titre** à une chaîne.</span><span class="sxs-lookup"><span data-stu-id="9b644-149">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="9b644-150">Au lieu de cela, vous lui affectez un ID de ressource (RESID), qui est défini dans la section **ressources** du manifeste.</span><span class="sxs-lookup"><span data-stu-id="9b644-150">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="9b644-151">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="9b644-151">SupportsPinning</span></span>

<span data-ttu-id="9b644-152">Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="9b644-152">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9b644-153">Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="9b644-153">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="9b644-154">Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office.</span><span class="sxs-lookup"><span data-stu-id="9b644-154">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="9b644-155">L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée.</span><span class="sxs-lookup"><span data-stu-id="9b644-155">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="9b644-156">Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="9b644-156">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!NOTE]
> <span data-ttu-id="9b644-157">Supportspinning n’est est actuellement uniquement pris en charge par Outlook 2016 ou version ultérieure sur Windows (Build 7628,1000 ou version ultérieure) et Outlook 2016 ou version ultérieure sur Mac (Build 16.13.503 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="9b644-157">SupportsPinning is currently only supported by Outlook 2016 or later on Windows (build 7628.1000 or later) and Outlook 2016 or later on Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
