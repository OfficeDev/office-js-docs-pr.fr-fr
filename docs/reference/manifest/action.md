---
title: Élément Action dans le fichier manifeste
description: Cet élément spécifie l’action à effectuer lorsque l’utilisateur sélectionne un bouton ou un contrôle de menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5eee17864ecacf101d7fbfe3519b3da500768deb
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467870"
---
# <a name="action-element"></a>Élément Action

Spécifie l’action à effectuer lorsque l’utilisateur sélectionne un contrôle [Bouton ou](control-button.md) [Menu](control-menu.md).

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Type d’action à effectuer|

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Spécifie le nom de la fonction à exécuter. |
|  [SourceLocation](#sourcelocation) |    Spécifie l’emplacement du fichier source pour cette action. |
|  [TaskpaneId](#taskpaneid) | Spécifie l’ID du conteneur de volet des tâches. Non pris en charge dans Outlook des modules.|
|  [Title](#title) | Indique le titre personnalisé du volet Office. Non pris en charge dans Outlook des modules.|
|  [SupportsPinning](#supportspinning) | Indique qu’un volet des tâches prend en charge l’épinglage, ce qui conserve le volet des tâches ouvert lorsque l’utilisateur modifie la sélection.|

## <a name="xsitype"></a>xsi:type

Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> L’inscription [des événements de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) [boîte aux lettres](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) et d’élément n’est pas disponible lorsque **xsi:type** est `ExecuteFunction`.

## <a name="functionname"></a>FunctionName

Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Élément obligatoire lorsque **xsi:type** est « ShowTaskpane ». Spécifie l’emplacement du fichier source pour cette action. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **Url** dans l’élément **Urls** dans l’élément [Resources](resources.md).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Élément facultatif quand **xsi:type** est « ShowTaskpane ». Spécifie l’ID du conteneur de volet des tâches. Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun. Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet. Lorsque les utilisateurs choisissent des commandes qui partagent le même **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ».

> [!NOTE]
> Cet élément n’est pas pris en charge dans Outlook.

L’exemple suivant montre deux actions qui partagent la même valeur **TaskpaneId**.

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

Les exemples suivants montrent deux actions qui utilisent une valeur **TaskpaneId** différente. Pour voir ces exemples en contexte, consultez l’article relatif à l’[exemple de commandes de complément](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

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

## <a name="title"></a>Titre

Élément facultatif quand **xsi:type** est « ShowTaskpane ». Indique le titre personnalisé du volet Office pour cette action.

> [!NOTE]
> Cet élément enfant n’est pas pris en charge dans Outlook de développement.

L’exemple suivant montre une action qui utilise **l’élément Title** . Notez que vous n’affectez pas directement le **titre** à une chaîne. Au lieu de cela, vous lui affectez un ID de ressource (résident), qui est défini dans la section **Ressources** du manifeste et ne peut pas être plus de 32 caractères.

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

## <a name="supportspinning"></a>SupportsPinning

Élément facultatif quand **xsi:type** a la valeur « ShowTaskpane ». Les éléments [VersionOverrides](versionoverrides.md) le contenant doivent avoir une valeur d’attribut `xsi:type` de `VersionOverridesV1_1`. Incluez cet élément avec une valeur `true` pour prendre en charge l’épinglage du volet Office. L’utilisateur pourra alors « épingler » le volet Office qui restera ouvert pendant que la sélection est modifiée. Pour en savoir plus, consultez l’article relatif à l’[implémentation d’un volet Office épinglable dans Outlook](../../outlook/pinnable-taskpane.md).

> [!IMPORTANT]
> Bien que l’élément `SupportsPinning` a été introduit dans l’ensemble de conditions requises [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), il est actuellement pris en charge uniquement pour les abonnés Microsoft 365 utilisant les éléments suivants :
>
> - Outlook 2016 ou une Windows (build 7628.1000 ou ultérieure)
> - Outlook 2016 ou une ultérieure sur Mac (build 16.13.503 ou ultérieure)
> - Outlook moderne sur le web

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
