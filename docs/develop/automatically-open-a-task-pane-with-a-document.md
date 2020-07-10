---
title: Ouvrir automatiquement un volet Office avec un document
description: Découvrez comment configurer un complément Office pour qu’il s’ouvre automatiquement lors de l’ouverture d’un document.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 85b421a569ccb83c3d07f0f10fd4767929332f96
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093706"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Ouvrir automatiquement un volet Office avec un document

Vous pouvez utiliser les commandes de complément dans votre complément Office pour étendre l’interface utilisateur Office en ajoutant des boutons au ruban de l’application Office. Lorsque les utilisateurs cliquent sur le bouton de commande, une action est réalisée, comme l’ouverture d’un volet des tâches.

Certains scénarios nécessitent qu’un volet des tâches s’ouvre automatiquement quand un document s’ouvre, sans intervention explicite de l’utilisateur. Vous pouvez utiliser la fonctionnalité d’ouverture automatique du volet des tâches, présentée dans l’ensemble des conditions AddInCommands 1.1, pour ouvrir automatiquement un volet des tâches lorsque votre scénario l’exige.


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>En quoi la fonctionnalité d’ouverture automatique est-elle différente de l’insertion d’un volet des tâches ?

Quand un utilisateur lance des compléments qui n’utilisent pas les commandes de complément, par exemple, les compléments qui s’exécutent dans Office 2013, ils sont insérés et conservés dans le document. Par conséquent, lorsque d’autres utilisateurs ouvrent le document, ils sont invités à installer le complément, puis le volet des tâches s’ouvre. Le défi de ce modèle est que, dans de nombreux cas, les utilisateurs ne veulent pas que le complément soit conservé dans le document. Par exemple, un étudiant qui utilise un complément de dictionnaire dans un document Word ne voudra peut-être pas que ses camarades de classe ou enseignants soient invités à installer ce complément lorsqu’ils ouvrent le document.

Avec la fonctionnalité d’ouverture automatique, vous pouvez explicitement définir ou autoriser l’utilisateur à déterminer si un complément de volet des tâches spécifique est conservé dans un document spécifique.

## <a name="support-and-availability"></a>Prise en charge et disponibilité

La fonctionnalité d’ouverture automatique est maintenant <!-- in **developer preview** and it is only --> prise en charge dans les produits et les plateformes suivantes.

|**Produits**|**Plateformes**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Plateformes prises en charge pour tous les produits :<ul><li>Office on Windows Desktop. Build 16.0.8121.1000+</li><li>Office on Mac. Build 15.34.17051500+</li><li>Office sur le web</li></ul>|


## <a name="best-practices"></a>Meilleures pratiques

Appliquez les meilleures pratiques suivantes lorsque vous utilisez la fonctionnalité d’ouverture automatique :

- Utilisez la fonctionnalité d’ouverture automatique quand elle vous aide à rendre vos utilisateurs de complément plus efficaces, comme dans les cas suivants :
  - When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.
  - When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- Utilisez la détection de l’ensemble de conditions requises pour déterminer si la fonctionnalité d’ouverture automatique est disponible et si ce n’est pas le cas.
- N’utilisez pas la fonctionnalité d’ouverture automatique pour augmenter artificiellement l’utilisation de votre complément. S’il n’est pas logique que votre complément s’ouvre automatiquement avec certains documents, cette fonctionnalité peut gêner les utilisateurs.

    > [!NOTE]
    > Si Microsoft détecte un abus de la fonctionnalité d’ouverture automatique, votre complément peut être rejeté d’AppSource.

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## <a name="implementation"></a>Implémentation

Pour implémenter la fonctionnalité d’ouverture automatique, procédez comme suit :

- Spécifiez le volet des tâches à ouvrir automatiquement.
- Ajoutez des balises au document pour ouvrir automatiquement le volet des tâches.

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.

### <a name="step-1-specify-the-task-pane-to-open"></a>Étape 1 : Spécifier le volet des tâches à ouvrir

To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.

L’exemple suivant illustre la valeur TaskPaneId définie sur Office.AutoShowTaskpaneWithDocument.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Étape 2 : Baliser le document pour ouvrir automatiquement le volet de tâches

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  


#### <a name="tag-the-document-on-the-client-side"></a>Baliser le document côté client

Utilisez la méthode Office.js [settings.set](/javascript/api/office/office.settings) pour définir **Office.AutoShowTaskpaneWithDocument** sur **true**, comme illustré dans l’exemple suivant.

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Utilisez cette méthode si vous devez baliser le document dans le cadre de vos interactions de complément (par exemple, dès que l’utilisateur crée une liaison ou choisit une option pour indiquer qu’il souhaite que le volet s’ouvre automatiquement).

#### <a name="use-open-xml-to-tag-the-document"></a>Utiliser Open XML pour baliser le document

You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

Ajoutez deux composants Open XML dans le document :

- Un composant `webextension`
- Un composant `taskpane`

L’exemple suivant montre comment ajouter le composant `webextension`.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Le composant `webextension` inclut un conteneur de propriétés et une propriété nommée **Office.AutoShowTaskpaneWithDocument** qui doit être définie sur `true`.

Le composant `webextension` comprend également une référence au store ou au catalogue avec des attributs pour `id`, `storeType`, `store` et `version`. Parmi les valeurs `storeType`, uniquement quatre sont pertinentes pour la fonctionnalité d’ouverture automatique. Les valeurs pour les trois autres attributs dépendent de la valeur pour `storeType`, comme illustré dans le tableau suivant.

| **`storeType`valeur** | **`id`valeur**    |**`store`valeur** | **`version`valeur**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|L’ID de la ressource AppSource du complément (voir la remarque).|Les paramètres régionaux d’AppSource ; par exemple, « fr-fr ».|La version dans le catalogue AppSource (voir la remarque).|
|Système de fichiers (un partage réseau)|Le GUID du complément dans le manifeste de complément.|Le chemin du partage réseau ; par exemple, « \\\\MyComputer\\MySharedFolder ».|La version dans le manifeste de complément.|
|EXCatalog (déploiement via le serveur Exchange) |Le GUID du complément dans le manifeste de complément.|« EXCatalog » La ligne excatalog est la ligne à utiliser avec des compléments qui utilisent un déploiement centralisé dans le centre d’administration 365 de Microsoft.|La version dans le manifeste de complément.
|Registre (Registre système)|Le GUID du complément dans le manifeste de complément.|« développeur »|La version dans le manifeste de complément.|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

Pour plus d’informations sur le balisage webextension, reportez-vous à [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).

L’exemple suivant montre comment ajouter le composant `taskpane`.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Notez que dans cet exemple, l’attribut `visibility` est défini sur « 0 ». Ainsi, après l’ajout des composants `taskpane` et du volet de tâches, l’utilisateur doit installer le complément via le bouton **Complément** sur le ruban lorsqu’il ouvre le document pour la première fois. Par la suite, le volet de tâches de complément s’ouvre automatiquement lorsque le fichier est ouvert. En outre, lorsque vous définissez `visibility` sur « 0 », vous pouvez utiliser Office.js pour autoriser les utilisateurs à activer ou à désactiver la fonctionnalité d’ouverture automatique. Plus spécifiquement, le script définit le paramètre de document **Office.AutoShowTaskpaneWithDocument** sur `true` ou `false`. (Pour plus d’informations, reportez-vous à la section [Baliser le document côté client](#tag-the-document-on-the-client-side).)

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.

Définir `visibility` sur « 1 » est un bon choix lorsque le complément et le modèle ou contenu du document sont tellement étroitement intégrés que l’utilisateur ne choisirait pas de désactiver la fonctionnalité d’ouverture automatique.

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.

## <a name="test-and-verify-opening-task-panes"></a>Tester et vérifier l’ouverture des volets Office

Vous pouvez déployer une version test de votre complément qui ouvrira automatiquement un volet des tâches à l’aide du déploiement centralisé via le centre d’administration Microsoft 365. L’exemple suivant montre la façon dont les compléments sont insérés à partir du catalogue de déploiement centralisé à l’aide de la version store d’EXCatalog.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Vous pouvez tester l’exemple précédent à l’aide de votre abonnement Microsoft 365 pour tester le déploiement centralisé et vérifier que votre complément fonctionne comme prévu. Si vous ne disposez pas déjà d’un abonnement Microsoft 365, vous pouvez obtenir gratuitement un abonnement Microsoft 365 renouvelable 90 jours en joignant le [programme de développement microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="see-also"></a>Voir aussi

Pour voir un exemple illustrant comment utiliser la fonctionnalité d’ouverture automatique, reportez-vous à [Exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).
[Participez au programme de développement Microsoft 365](/office/developer-program/office-365-developer-program).
