---
title: Ouvrir automatiquement un volet Office avec un document
description: Découvrez comment configurer un Office pour qu’il s’ouvre automatiquement lorsqu’un document s’ouvre.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: a9683f63b82232f8f5697007692b359ae06b7650e96866a2425e2d900ded4d8a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081210"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Ouvrir automatiquement un volet de tâches avec un document

Vous pouvez utiliser les commandes de votre Office pour étendre l’interface utilisateur Office en ajoutant des boutons au application Office ruban. Lorsque les utilisateurs cliquent sur le bouton de commande, une action est réalisée, comme l’ouverture d’un volet des tâches.

Certains scénarios nécessitent qu’un volet des tâches s’ouvre automatiquement quand un document s’ouvre, sans intervention explicite de l’utilisateur. Vous pouvez utiliser la fonctionnalité d’ouverture automatique du volet des tâches, présentée dans l’ensemble des conditions AddInCommands 1.1, pour ouvrir automatiquement un volet des tâches lorsque votre scénario l’exige.

## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>En quoi la fonctionnalité d’ouverture automatique est-elle différente de l’insertion d’un volet des tâches ?

Quand un utilisateur lance des compléments qui n’utilisent pas les commandes de complément, par exemple, les compléments qui s’exécutent dans Office 2013, ils sont insérés et conservés dans le document. Par conséquent, lorsque d’autres utilisateurs ouvrent le document, ils sont invités à installer le complément, puis le volet des tâches s’ouvre. La difficulté de ce modèle est que, dans de nombreux cas, les utilisateurs ne veulent pas que le module soit persistant dans le document. Par exemple, un étudiant qui utilise un complément de dictionnaire dans un document Word ne voudra peut-être pas que ses camarades de classe ou enseignants soient invités à installer ce complément lorsqu’ils ouvrent le document.

Avec la fonctionnalité d’ouverture automatique, vous pouvez explicitement définir ou autoriser l’utilisateur à déterminer si un complément de volet des tâches spécifique est conservé dans un document spécifique.

## <a name="support-and-availability"></a>Prise en charge et disponibilité

La fonctionnalité d’ouverture automatique est maintenant <!-- in **developer preview** and it is only --> prise en charge dans les produits et les plateformes suivantes.

|**Produits**|**Plateformes**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Plateformes prises en charge pour tous les produits :<ul><li>Office pour bureau Windows. Build 16.0.8121.1000+</li><li>Office sur Mac. Build 15.34.17051500+</li><li>Office sur le web</li></ul>|

## <a name="best-practices"></a>Meilleures pratiques

Appliquez les meilleures pratiques suivantes lorsque vous utilisez la fonctionnalité d’ouverture automatique.

- Utilisez la fonctionnalité d’ouverture automatique quand elle vous aide à rendre vos utilisateurs de complément plus efficaces, comme dans les cas suivants :
  - Lorsque le document a besoin du complément pour fonctionner correctement. Par exemple, une feuille de calcul qui contient des valeurs de stock régulièrement actualisées par un complément. Le complément doit s’ouvrir automatiquement lorsque la feuille de calcul est ouverte pour maintenir les valeurs à jour.
  - Lorsque l’utilisateur sera le plus susceptible d’utiliser le complément avec un document particulier. Par exemple, un complément qui permet aux utilisateurs de renseigner ou de modifier des données dans un document en extrayant des informations à partir d’un système principal.
- Autorisez les utilisateurs à activer ou à désactiver la fonctionnalité d’ouverture automatique. Incluez une option dans votre interface utilisateur pour choisir de ne plus ouvrir automatiquement le volet des tâches de complément.  
- Utilisez la détection de l’ensemble de conditions requises pour déterminer si la fonctionnalité d’ouverture automatique est disponible et fournir un comportement de base si ce n’est pas le cas.
- N’utilisez pas la fonctionnalité d’ouverture automatique pour augmenter artificiellement l’utilisation de votre complément. S’il n’est pas logique que votre application s’ouvre automatiquement avec certains documents, cette fonctionnalité peut gêner les utilisateurs.

    > [!NOTE]
    > Si Microsoft détecte un abus de la fonctionnalité d’ouverture automatique, votre complément peut être rejeté d’AppSource.

- N’utilisez pas cette fonctionnalité pour épingler plusieurs volets de tâches. Vous pouvez uniquement définir l’ouverture automatique d’un volet de votre complément avec un document.  

## <a name="implement-the-autoopen-feature"></a>Implémenter la fonctionnalité d’ouverture automatique

- Spécifiez le volet des tâches à ouvrir automatiquement.
- Ajoutez des balises au document pour ouvrir automatiquement le volet des tâches.

> [!IMPORTANT]
> Le volet des tâches à ouvrir automatiquement s’ouvre uniquement si le complément est déjà installé sur l’appareil de l’utilisateur. Si le complément n’est pas installé lorsque l’utilisateur ouvre un document, la fonctionnalité d’ouverture automatique ne fonctionnera pas et le paramètre sera ignoré. Si vous avez également besoin que le complément soit distribué avec le document, vous devez définir la propriété de visibilité sur 1. Cette opération peut uniquement être effectuée à l’aide d’OpenXML. Un exemple est fourni plus loin dans cet article.

### <a name="step-1-specify-the-task-pane-to-open"></a>Étape 1 : Spécifier le volet des tâches à ouvrir

Pour spécifier le volet de tâches à ouvrir automatiquement, définissez la valeur [TaskpaneId](../reference/manifest/action.md#taskpaneid) sur **Office.AutoShowTaskpaneWithDocument**. Vous pouvez uniquement définir cette valeur sur un seul volet de tâches. Si vous définissez cette valeur sur plusieurs volets de tâches, la première occurrence de la valeur sera reconnue et les autres seront ignorées.

L’exemple suivant illustre la valeur TaskPaneId définie sur Office.AutoShowTaskpaneWithDocument.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Étape 2 : Baliser le document pour ouvrir automatiquement le volet de tâches

Vous pouvez baliser le document pour déclencher la fonctionnalité d’ouverture automatique de deux façons possibles. Choisissez l’alternative qui convient le mieux à votre scénario.  

#### <a name="tag-the-document-on-the-client-side"></a>Baliser le document côté client

Utilisez la méthode Office.js [settings.set](/javascript/api/office/office.settings) pour définir **Office.AutoShowTaskpaneWithDocument** sur **true**, comme illustré dans l’exemple suivant.

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Utilisez cette méthode si vous devez baliser le document dans le cadre de vos interactions de complément (par exemple, dès que l’utilisateur crée une liaison ou choisit une option pour indiquer qu’il souhaite que le volet s’ouvre automatiquement).

#### <a name="use-open-xml-to-tag-the-document"></a>Utiliser Open XML pour baliser le document

Vous pouvez utiliser Open XML pour créer ou modifier un document et ajouter le balisage Open Office XML approprié afin de déclencher la fonctionnalité d’ouverture automatique. Pour obtenir un exemple montrant comment procéder, voir [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

Ajoutez deux parties Open XML au document.

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
|EXCatalog (déploiement via le serveur Exchange) |Le GUID du complément dans le manifeste de complément.|« EXCatalog » La ligne EXCatalog est la ligne à utiliser avec les add-ins qui utilisent le déploiement centralisé dans le Centre d’administration Microsoft 365.|La version dans le manifeste de complément.
|Registre (Registre système)|Le GUID du complément dans le manifeste de complément.|« développeur »|La version dans le manifeste de complément.|

> [!NOTE]
> Pour trouver l’ID de ressource et la version d’un complément dans AppSource, accédez à la page d’accueil d’AppSource pour le complément. L’ID de ressource apparaît dans la barre d’adresse dans le navigateur. La version est répertoriée dans la section **Détails** de la page.

Pour plus d’informations sur le balisage webextension, reportez-vous à [[MS-OWEXML] 2.2.5. WebExtensionReference](/openspecs/office_standards/ms-owexml/d4081e0b-5711-45de-b708-1dfa1b943ad1).

L’exemple suivant montre comment ajouter le composant `taskpane`.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Notez que dans cet exemple, l’attribut `visibility` est défini sur « 0 ». Ainsi, après l’ajout des composants `taskpane` et du volet de tâches, l’utilisateur doit installer le complément via le bouton **Complément** sur le ruban lorsqu’il ouvre le document pour la première fois. Par la suite, le volet de tâches de complément s’ouvre automatiquement lorsque le fichier est ouvert. En outre, lorsque vous définissez `visibility` sur « 0 », vous pouvez utiliser Office.js pour autoriser les utilisateurs à activer ou à désactiver la fonctionnalité d’ouverture automatique. Plus spécifiquement, le script définit le paramètre de document **Office.AutoShowTaskpaneWithDocument** sur `true` ou `false`. (Pour plus d’informations, reportez-vous à la section [Baliser le document côté client](#tag-the-document-on-the-client-side).)

Si `visibility` est défini sur « 1 », le volet de tâches s’ouvre automatiquement à la première ouverture du document. L’utilisateur est invité à approuver le complément. Lorsque ce dernier est approuvé, le complément s’ouvre. Par la suite, le volet de tâches de complément s’ouvre automatiquement lorsque le fichier est ouvert. Toutefois, lorsque `visibility` est défini sur « 1 », vous ne pouvez pas utiliser Office.js pour autoriser les utilisateurs à activer ou à désactiver la fonctionnalité d’ouverture automatique.

Définir `visibility` sur « 1 » est un bon choix lorsque le complément et le modèle ou contenu du document sont tellement étroitement intégrés que l’utilisateur ne choisirait pas de désactiver la fonctionnalité d’ouverture automatique.

> [!NOTE]
> Si vous voulez distribuer votre complément avec le document, pour que les utilisateurs soient invités à l’installer, vous devez définir la propriété de visibilité sur 1. Cette opération peut uniquement être effectuée à l’aide d’Open XML.

Un moyen simple d’écrire le XML consiste à d’abord exécuter votre add-in et baliser le document côté [client](#tag-the-document-on-the-client-side) pour écrire la valeur, puis enregistrer le document et inspecter le XML qui est généré. Office détecter et fournir les valeurs d’attribut appropriées. Vous pouvez également utiliser l’outil de productivité du [SDK Open XML](https://www.nuget.org/packages/Open-XML-SDK) pour générer du code C# pour ajouter par programme le code basé sur le code XML que vous générez.

## <a name="test-and-verify-opening-task-panes"></a>Tester et vérifier l’ouverture des volets Office

Vous pouvez déployer une version de test de votre application qui ouvre automatiquement un volet Des tâches à l’aide du déploiement centralisé via le Centre d’administration Microsoft 365. L’exemple suivant montre la façon dont les compléments sont insérés à partir du catalogue de déploiement centralisé à l’aide de la version store d’EXCatalog.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Vous pouvez tester l’exemple précédent à l’aide de votre abonnement Microsoft 365 pour tester le déploiement centralisé et vérifier que votre add-in fonctionne comme prévu. Si vous n’avez pas encore d’abonnement Microsoft 365, vous pouvez obtenir un abonnement gratuit de 90 jours renouvelable Microsoft 365 en rejoignant le programme Microsoft 365 [développeur.](https://developer.microsoft.com/office/dev-program)

## <a name="see-also"></a>Voir aussi

Pour voir un exemple illustrant comment utiliser la fonctionnalité d’ouverture automatique, reportez-vous à [Exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).
[Rejoignez le Microsoft 365 développeur.](/office/developer-program/office-365-developer-program)