---
title: Ouvrir automatiquement un volet Office avec un document
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 2ebce1ce8bd95ee7802b5509d375f1986bb2877e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505915"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Ouvrir automatiquement un volet Office avec un document

Vous pouvez utiliser des commandes de complément dans votre complément Office pour étendre l’interface utilisateur Office en ajoutant des boutons au ruban Office. Lorsque les utilisateurs cliquent sur le bouton de commande, une action est réalisée, comme l’ouverture d’un volet des tâches. 

Certains scénarios nécessitent qu’un volet Office s’ouvre automatiquement lorsqu'un document s’ouvre, sans intervention explicite de l’utilisateur. Vous pouvez utiliser la fonctionnalité d’ouverture automatique du volet des tâches, présentée dans l’ensemble des conditions AddInCommands 1.1, pour ouvrir automatiquement un volet Office lorsque votre scénario l’exige. 


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>En quoi la caractéristique d’ouverture automatique est-elle différente de l’insertion d’un volet Office ? 

Lorsqu'un utilisateur lance des compléments qui n’utilisent pas les commandes de complément, par exemple, les compléments qui s’exécutent dans Office 2013, ils sont insérés et conservés dans le document. Par conséquent, lorsque d’autres utilisateurs ouvrent le document, ils sont invités à installer le complément, puis le volet des tâches s’ouvre. La difficulté avec ce modèle est que dans de nombreux cas, les utilisateurs ne veulent pas que le complément soit conservé dans le document. Par exemple, un étudiant qui utilise un complément de dictionnaire dans un document Word ne voudra peut-être pas que ses camarades de classe ou enseignants soient invités à installer ce complément lorsqu’ils ouvrent le document.  

Avec la caractéristique d’ouverture automatique, vous pouvez explicitement définir ou autoriser l’utilisateur à déterminer si un complément de volet Office spécifique est conservé dans un document spécifique. 

## <a name="support-and-availability"></a>Support et disponibilité
La caractéristique d’ouverture automatique est actuellement <!-- in **developer preview** and it is only --> prise en charge dans les plateformes et produits suivants.

|**Produits**|**Plateformes**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Plateformes supportées pour tous les produits :<ul><li>Office pour bureau Windows. Build 16.0.8121.1000 et versions ultérieures</li><li>Office pour Mac. Build 15.34.17051500 et versions ultérieures</li><li>Office Online</li></ul>|


## <a name="best-practices"></a>Meilleures pratiques

Appliquez les meilleures pratiques suivantes lorsque vous utilisez la caractéristique d’ouverture automatique :

- Utilisez la caractéristique d’ouverture automatique lorsqu'elle vous aide à rendre vos utilisateurs de complément plus efficaces, comme dans les cas suivants :
    - Lorsque le document a besoin du complément pour fonctionner correctement. Par exemple, une feuille de calcul qui contient des valeurs de stock régulièrement actualisées par un complément. Le complément doit s’ouvrir automatiquement lorsque la feuille de calcul est ouverte pour maintenir les valeurs à jour. 
    - Lorsque l’utilisateur sera le plus susceptible d’utiliser le complément avec un document particulier. Par exemple, un complément qui permet aux utilisateurs de renseigner ou de modifier des données dans un document en extrayant des informations à partir d’un système principal. 
- Autorisez les utilisateurs à activer ou à désactiver la caractéristique d’ouverture automatique. Incluez une option dans votre interface utilisateur pour choisir de ne plus ouvrir automatiquement le volet Office de complément.  
- Utilisez la détection de l’ensemble d'exigences pour déterminer si la caractéristique d’ouverture automatique est disponible, et fournissez un comportement de secours si elle ne l’est pas.
- N’utilisez pas la caractéristique d’ouverture automatique pour augmenter artificiellement l’utilisation de votre complément. Si l’ouverture automatique du complément n’est pas pertinente pour certains documents, cette caractéristique peut gêner les utilisateurs. 

    > [!NOTE]
    > Si Microsoft détecte un abus de la  caractéristique d’ouverture automatique, votre complément peut être rejeté d’AppSource. 

- N’utilisez pas cette caractéristique pour repérer plusieurs volets Office. Vous pouvez uniquement définir l’ouverture automatique d’un volet de votre complément avec un document.  

## <a name="implementation"></a>Implémentation
Pour implémenter la caractéristique d’ouverture automatique, procédez comme suit :

- Spécifiez le volet Office à ouvrir automatiquement.
- Ajoutez des balises au document pour ouvrir automatiquement le volet Office.

> [!IMPORTANT]
> Le volet des tâches à ouvrir automatiquement s’ouvre uniquement si le complément est déjà installé sur l’appareil de l’utilisateur. Si le complément n’est pas installé lorsque l’utilisateur ouvre un document, la caractéristique d’ouverture automatique ne fonctionnera pas et le paramètre sera ignoré. Si vous avez également besoin que le complément soit distribué avec le document, vous devez définir la propriété de visibilité sur 1. Cette opération peut uniquement être effectuée à l’aide d’OpenXML. Un exemple est fourni plus loin dans cet article. 

### <a name="step-1-specify-the-task-pane-to-open"></a>Étape 1 : Spécifier le volet Office à ouvrir
Pour spécifier le volet Office  à ouvrir automatiquement, définissez la valeur [TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) sur **Office.AutoShowTaskpaneWithDocument**. Vous pouvez uniquement définir cette valeur sur un seul volet Office. Si vous définissez cette valeur sur plusieurs volets Office, la première occurrence de la valeur sera reconnue et les autres seront ignorées. 

L’exemple suivant illustre la valeur TaskPaneId définie sur Office.AutoShowTaskpaneWithDocument.
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Étape 2 : Baliser le document pour ouvrir automatiquement le volet Office

Vous pouvez ajouter une balise le document pour déclencher la caractéristique ouverture automatique dans l'une des deux manières. Choisissez la solution qui convient le mieux pour votre scénario.  


#### <a name="tag-the-document-on-the-client-side"></a>Balisez le document côté client
Utilisez la méthode Office.js [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) pour définir **Office.AutoShowTaskpaneWithDocument** sur **true**, comme illustré dans l’exemple suivant.   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Utilisez cette méthode si vous devez baliser le document dans le cadre de vos interactions de complément (par exemple, dès que l’utilisateur crée une liaison ou choisit une option pour indiquer qu’il souhaite que le volet s’ouvre automatiquement).

#### <a name="use-open-xml-to-tag-the-document"></a>Utilisez Open XML pour baliser le document
Vous pouvez utiliser Open XML pour créer ou modifier un document et ajouter le balisage Open Office XML approprié afin de déclencher la  caractéristique d’ouverture automatique. Pour obtenir un exemple montrant comment procéder, voir [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin). 

Ajoutez deux composants Open XML dans le document :

- Un composant webextension
- Un composant volet Office

L’exemple suivant montre comment ajouter le composant webextension.

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

Le composant webextension inclut un conteneur de propriétés et une propriété nommée **Office.AutoShowTaskpaneWithDocument** qui doit être définie sur `true`.

Le composant webextension comprend également une référence au store ou au catalogue avec des attributs pour `id`, `storeType`, `store`, et `version`. Parmi les valeurs `storeType`, uniquement quatre sont pertinentes pour la  caractéristique d’ouverture automatique. Les valeurs pour les trois autres attributs dépendent de la valeur pour `storeType`, comme illustré dans le tableau suivant. 

| **`storeType` valeur** | **`id` valeur**    |**`store` valeur** | **`version` valeur**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|L’ID de l'élément AppSource du complément (voir la remarque).|Les paramètres régionaux d’AppSource ; par exemple, « fr-fr ».|La version dans le catalogue AppSource (voir la remarque).|
|Système de fichiers (un partage réseau)|Le GUID du complément dans le manifeste de complément.|Le chemin du partage réseau ; par exemple, « \\\\MyComputer\\MySharedFolder ».|La version dans le manifeste de complément.|
|EXCatalog (déploiement via Exchange server) |Le GUID du complément dans le manifeste de complément.|La ligne EXCatalog est la ligne à utiliser avec les compléments qui utilisent le déploiement centralisé dans le centre d’administration Office 365.|La version dans le manifeste de complément.
|Registre (Registre système)|Le GUID du complément dans le manifeste de complément.|« développeur »|La version dans le manifeste de complément.|

> [!NOTE]
> Pour trouver l’ID d'élément et la version d’un complément dans AppSource, accédez à la page d’arrivée d’AppSource pour le complément. L’ID d'élement apparaît dans la barre d’adresse dans le navigateur. La version est répertoriée dans la section **Détails** de la page.

Pour plus d’informations sur le balisage webextension, reportez-vous à [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).

L’exemple suivant montre comment ajouter le composant du volet Office.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Notez que dans cet exemple, l’attribut `visibility` est défini sur « 0 ». Ainsi, après l’ajout des composants webextension et du volet de tâches, l’utilisateur doit installer le complément via le bouton **Complément** sur le ruban lorsqu’il ouvre le document pour la première fois. Par la suite, le volet de tâches de complément s’ouvre automatiquement lorsque le fichier est ouvert. En outre, lorsque vous définissez `visibility` sur « 0 », vous pouvez utiliser Office.js pour autoriser les utilisateurs à activer ou à désactiver la caractéristique d’ouverture automatique. Plus spécifiquement, le script de jeux définit le paramètre de document **Office.AutoShowTaskpaneWithDocument** sur `true` ou `false`. (Pour plus d’informations, voir [Baliser le document côté client](#tag-the-document-on-the-client-side).) 

Si `visibility` est défini sur « 1 », le volet Office s’ouvre automatiquement à la première ouverture du document. L’utilisateur est invité à faire confiance au complément. Lorsque ce dernier est approuvé, le complément s’ouvre. Par la suite, le volet Office de complément s’ouvre automatiquement lorsque le fichier est ouvert. Toutefois, lorsque `visibility` est défini sur « 1 », vous ne pouvez pas utiliser Office.js pour autoriser les utilisateurs à activer ou à désactiver la caractéristique d’ouverture automatique. 

Définir `visibility` sur « 1 » est un bon choix lorsque le complément et le modèle ou contenu du document sont tellement étroitement intégrés que l’utilisateur ne choisirait pas de désactiver la caractéristique d’ouverture automatique. 

> [!NOTE]
> Si vous voulez distribuer votre complément avec le document, pour que les utilisateurs soient invités à l’installer, vous devez définir la propriété de visibilité sur 1. Cette opération peut uniquement être effectuée à l’aide d’Open XML.

Une méthode simple d’écriture du code XML consiste à exécuter d’abord votre complément, puis à [baliser le document côté client](#tag-the-document-on-the-client-side) pour écrire la valeur, à enregistrer le document et à inspecter le code XML généré. Office détectera et fournira les valeurs d’attribut appropriées. Vous pouvez également utiliser l’outil de productivité [Kit de développement logiciel Open XML 2.5](https://www.microsoft.com/download/details.aspx?id=30425) pour générer le code C# pour ajouter par programme le balisage en fonction du XML vous générez.

## <a name="test-and-verify-opening-taskpanes"></a>Tester et vérifier les tâches d'ouverture
Vous pouvez déployer une version d’évaluation de votre complément qui ouvre automatiquement un volet Office à l’aide du Déploiement Centralisé via le centre d’administration Office 365. L’exemple suivant montre comment les compléments sont insérés à partir du catalogue de Déploiement Centralisé à l’aide de la version magasin de EXCatalog.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
Pour tester l’exemple précédent, nous vous recommandons de participer au [Programme de développement Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program) et de vous inscrire pour un [compte de développeur Office 365](https://developer.microsoft.com/office/dev-program) si vous ne possédez pas encore un abonnement à Office 365. Vous pouvez réellement tester le déploiement centralisé et vérifier que votre complément fonctionne comme prévu.


## <a name="see-also"></a>Voir aussi

Pour obtenir un exemple qui montre comment utiliser la caractéristique ouverture automatique, voir [exemples de commandes Office Add-in](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane). [Participer au programme du développeur Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program). 

