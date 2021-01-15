---
title: Élément Action dans le fichier manifeste
description: Cet élément spécifie l’action à effectuer lorsque l’utilisateur sélectionne un bouton ou un contrôle de menu.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: e345d0a1682e0125373a309e1e56eb2d6298ac7d
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771409"
---
# <a name="action-element"></a><span data-ttu-id="33705-103">Élément Action</span><span class="sxs-lookup"><span data-stu-id="33705-103">Action element</span></span>

<span data-ttu-id="33705-104">Spécifie l’action à effectuer lorsque l’utilisateur sélectionne un  [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) .</span><span class="sxs-lookup"><span data-stu-id="33705-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="33705-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="33705-105">Attributes</span></span>

|  <span data-ttu-id="33705-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="33705-106">Attribute</span></span>  |  <span data-ttu-id="33705-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="33705-107">Required</span></span>  |  <span data-ttu-id="33705-108">Description</span><span class="sxs-lookup"><span data-stu-id="33705-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="33705-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="33705-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="33705-110">Oui</span><span class="sxs-lookup"><span data-stu-id="33705-110">Yes</span></span>  | <span data-ttu-id="33705-111">Type d’action à effectuer</span><span class="sxs-lookup"><span data-stu-id="33705-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="33705-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="33705-112">Child elements</span></span>

|  <span data-ttu-id="33705-113">Élément</span><span class="sxs-lookup"><span data-stu-id="33705-113">Element</span></span> |  <span data-ttu-id="33705-114">Description</span><span class="sxs-lookup"><span data-stu-id="33705-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="33705-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="33705-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="33705-116">Spécifie le nom de la fonction à exécuter.</span><span class="sxs-lookup"><span data-stu-id="33705-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="33705-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="33705-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="33705-118">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="33705-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="33705-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="33705-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="33705-120">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="33705-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="33705-121">Title</span><span class="sxs-lookup"><span data-stu-id="33705-121">Title</span></span>](#title) | <span data-ttu-id="33705-122">Indique le titre personnalisé du volet Office.</span><span class="sxs-lookup"><span data-stu-id="33705-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="33705-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="33705-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="33705-124">Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.</span><span class="sxs-lookup"><span data-stu-id="33705-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="33705-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="33705-125">xsi:type</span></span>

<span data-ttu-id="33705-p101">Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="33705-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="33705-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="33705-128">FunctionName</span></span>

<span data-ttu-id="33705-p102">Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="33705-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="33705-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="33705-132">SourceLocation</span></span>

<span data-ttu-id="33705-133">Élément obligatoire lorsque **xsi : type** est « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="33705-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="33705-134">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="33705-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="33705-135">L’attribut **RESID** ne peut pas contenir plus de 32 caractères et doit être défini sur la valeur de l’attribut **ID** d’un élément **URL** dans l’élément **URL** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="33705-135">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="33705-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="33705-136">TaskpaneId</span></span>

<span data-ttu-id="33705-p104">Élément facultatif quand **xsi:type** est « ShowTaskpane ». Spécifie l’ID du conteneur de volet des tâches. Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun. Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet. Lorsque les utilisateurs choisissent des commandes qui partagent le même **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».</span><span class="sxs-lookup"><span data-stu-id="33705-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="33705-142">Cet élément n’est pas pris en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="33705-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="33705-143">L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="33705-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="33705-p105">Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="33705-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="33705-146">Titre</span><span class="sxs-lookup"><span data-stu-id="33705-146">Title</span></span>

<span data-ttu-id="33705-147">Élément facultatif quand **xsi:type** est « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="33705-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="33705-148">Indique le titre personnalisé du volet Office pour cette action.</span><span class="sxs-lookup"><span data-stu-id="33705-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="33705-149">L’exemple suivant montre une action qui utilise l’élément **title** .</span><span class="sxs-lookup"><span data-stu-id="33705-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="33705-150">Notez que vous n’affectez pas directement le **titre** à une chaîne.</span><span class="sxs-lookup"><span data-stu-id="33705-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="33705-151">Au lieu de cela, vous lui affectez un ID de ressource (RESID), qui est défini dans la section **ressources** du manifeste et ne peut pas comporter plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="33705-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="33705-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="33705-152">SupportsPinning</span></span>

<span data-ttu-id="33705-153">Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="33705-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="33705-154">Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="33705-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="33705-155">Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office.</span><span class="sxs-lookup"><span data-stu-id="33705-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="33705-156">L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée.</span><span class="sxs-lookup"><span data-stu-id="33705-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="33705-157">Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="33705-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33705-158">Bien que l' `SupportsPinning` élément ait été introduit dans l' [ensemble de conditions requises 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), il est actuellement uniquement pris en charge pour les abonnés Microsoft 365 à l’aide des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="33705-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="33705-159">Outlook 2016 ou version ultérieure sur Windows (version 7628,1000 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="33705-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="33705-160">Outlook 2016 ou version ultérieure sur Mac (Build 16.13.503 ou version ultérieure)</span><span class="sxs-lookup"><span data-stu-id="33705-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="33705-161">Outlook moderne sur le web</span><span class="sxs-lookup"><span data-stu-id="33705-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
