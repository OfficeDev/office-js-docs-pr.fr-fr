---
title: Élément Action dans le fichier manifeste
description: Cet élément spécifie l’action à effectuer lorsque l’utilisateur sélectionne un bouton ou un contrôle de menu.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 6be1430800dea27dbd9bf78607161d88e475c145
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505408"
---
# <a name="action-element"></a><span data-ttu-id="c2838-103">Élément Action</span><span class="sxs-lookup"><span data-stu-id="c2838-103">Action element</span></span>

<span data-ttu-id="c2838-104">Spécifie l’action à effectuer lorsque l’utilisateur sélectionne un contrôle  [Bouton](control.md#button-control) [ou](control.md#menu-dropdown-button-controls) Menu.</span><span class="sxs-lookup"><span data-stu-id="c2838-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="c2838-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="c2838-105">Attributes</span></span>

|  <span data-ttu-id="c2838-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="c2838-106">Attribute</span></span>  |  <span data-ttu-id="c2838-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="c2838-107">Required</span></span>  |  <span data-ttu-id="c2838-108">Description</span><span class="sxs-lookup"><span data-stu-id="c2838-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c2838-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c2838-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="c2838-110">Oui</span><span class="sxs-lookup"><span data-stu-id="c2838-110">Yes</span></span>  | <span data-ttu-id="c2838-111">Type d’action à effectuer</span><span class="sxs-lookup"><span data-stu-id="c2838-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="c2838-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c2838-112">Child elements</span></span>

|  <span data-ttu-id="c2838-113">Élément</span><span class="sxs-lookup"><span data-stu-id="c2838-113">Element</span></span> |  <span data-ttu-id="c2838-114">Description</span><span class="sxs-lookup"><span data-stu-id="c2838-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="c2838-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="c2838-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="c2838-116">Spécifie le nom de la fonction à exécuter.</span><span class="sxs-lookup"><span data-stu-id="c2838-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="c2838-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c2838-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="c2838-118">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="c2838-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="c2838-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="c2838-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="c2838-120">Spécifie l’ID du conteneur de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="c2838-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="c2838-121">Title</span><span class="sxs-lookup"><span data-stu-id="c2838-121">Title</span></span>](#title) | <span data-ttu-id="c2838-122">Indique le titre personnalisé du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c2838-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="c2838-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="c2838-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="c2838-124">Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.</span><span class="sxs-lookup"><span data-stu-id="c2838-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="c2838-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c2838-125">xsi:type</span></span>

<span data-ttu-id="c2838-p101">Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="c2838-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="c2838-128">L’inscription [des événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) boîte [aux](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) lettres et d’élément n’est pas disponible lorsque **xsi:type** est `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="c2838-128">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="c2838-129">FunctionName</span><span class="sxs-lookup"><span data-stu-id="c2838-129">FunctionName</span></span>

<span data-ttu-id="c2838-p102">Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="c2838-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="c2838-133">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c2838-133">SourceLocation</span></span>

<span data-ttu-id="c2838-134">Élément obligatoire lorsque **xsi:type** est « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="c2838-134">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="c2838-135">Spécifie l’emplacement du fichier source pour cette action.</span><span class="sxs-lookup"><span data-stu-id="c2838-135">Specifies the source file location for this action.</span></span> <span data-ttu-id="c2838-136">**L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Urls** dans l’élément [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="c2838-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="c2838-137">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="c2838-137">TaskpaneId</span></span>

<span data-ttu-id="c2838-p104">Élément facultatif quand **xsi:type** est « ShowTaskpane ». Spécifie l’ID du conteneur de volet des tâches. Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun. Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet. Lorsque les utilisateurs choisissent des commandes qui partagent le même **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».</span><span class="sxs-lookup"><span data-stu-id="c2838-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="c2838-143">Cet élément n’est pas pris en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2838-143">This element is not supported in Outlook.</span></span>

<span data-ttu-id="c2838-144">L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="c2838-144">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="c2838-p105">Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="c2838-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="c2838-147">Titre</span><span class="sxs-lookup"><span data-stu-id="c2838-147">Title</span></span>

<span data-ttu-id="c2838-148">Élément facultatif quand **xsi:type** est « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="c2838-148">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="c2838-149">Indique le titre personnalisé du volet Office pour cette action.</span><span class="sxs-lookup"><span data-stu-id="c2838-149">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="c2838-150">L’exemple suivant illustre une action qui utilise **l’élément Title.**</span><span class="sxs-lookup"><span data-stu-id="c2838-150">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="c2838-151">Notez que vous n’affectez pas directement le **titre** à une chaîne.</span><span class="sxs-lookup"><span data-stu-id="c2838-151">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="c2838-152">Au lieu de cela, vous lui affectez un ID de ressource (résident), qui est défini dans la section **Ressources** du manifeste et ne peut pas être plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="c2838-152">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="c2838-153">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="c2838-153">SupportsPinning</span></span>

<span data-ttu-id="c2838-154">Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="c2838-154">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="c2838-155">Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="c2838-155">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="c2838-156">Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c2838-156">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="c2838-157">L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée.</span><span class="sxs-lookup"><span data-stu-id="c2838-157">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="c2838-158">Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="c2838-158">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c2838-159">Bien que l’élément a été introduit dans l’ensemble de conditions requises 1.5, il est actuellement uniquement pris en charge pour les abonnés `SupportsPinning` Microsoft 365 à l’aide des éléments suivants. [](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)</span><span class="sxs-lookup"><span data-stu-id="c2838-159">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
>
> - <span data-ttu-id="c2838-160">Outlook 2016 ou une ultérieure sur Windows (build 7628.1000 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="c2838-160">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="c2838-161">Outlook 2016 ou une ultérieure sur Mac (build 16.13.503 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="c2838-161">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="c2838-162">Outlook moderne sur le web</span><span class="sxs-lookup"><span data-stu-id="c2838-162">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
