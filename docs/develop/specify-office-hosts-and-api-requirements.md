---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: Découvrez comment spécifier Office applications et les conditions requises de l’API pour que votre module fonctionne comme prévu.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e0cf0a99706861a5446512542b28f3b27db54d8d
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263050"
---
# <a name="specify-office-applications-and-api-requirements"></a>Spécifier les applications Office et les exigences de l’API

Votre Office peut dépendre d’une application Office spécifique (également appelée hôte Office) ou de membres spécifiques de l’API JavaScript Office (office.js). Par exemple, votre complément peut :

- Exécuter dans une application Office (Word ou Excel), ou plusieurs applications.
- Utilisez des API JavaScript Office disponibles uniquement dans certaines versions de Office. Par exemple, la version d’achat Excel 2016 ne prend pas en charge toutes les API Excel de la bibliothèque JavaScript Office.

Dans ces situations, vous devez vous assurer que votre application n’est jamais installée sur des applications Office ou des versions Office dans lesquelles elle ne peut pas s’exécuter.

Il existe également des scénarios dans lesquels vous souhaitez contrôler les fonctionnalités de votre add-in qui sont visibles par les utilisateurs en fonction de leur application Office et de leur version Office version. Deux exemples sont les suivants :

- Votre complément comporte des fonctionnalités utiles dans Word et PowerPoint, telles que la manipulation de texte, mais il comporte des fonctionnalités supplémentaires qui n’ont de sens que dans PowerPoint, telles que les fonctionnalités de gestion des diapositives. Vous devez masquer les fonctionnalités PowerPoint uniquement lorsque le module est en cours d’exécution dans Word.
- Votre application dispose d’une fonctionnalité qui nécessite une méthode d’API JavaScript Office qui est prise en charge dans certaines versions d’une application Office, telles que les Excel d’abonnement, mais qui n’est pas prise en charge dans d’autres, telle que les Excel 2016 d’achat unique. Toutefois, votre application comporte d’autres fonctionnalités qui nécessitent  uniquement Office méthodes d’API JavaScript qui sont pris en charge dans Excel 2016. Dans ce scénario, vous avez besoin que le module soit installable sur Excel 2016, mais la fonctionnalité qui nécessite la méthode non pris en Excel 2016.

Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

> [!NOTE]
> Pour obtenir une vue d’Office l’endroit où les Office sont actuellement pris en charge, consultez la page Office sur la disponibilité des applications [clientes](../overview/office-add-in-availability.md) et de la plateforme pour les Office de recherche.

> [!TIP]
> Bon nombre des tâches décrites dans cet article sont réalisées pour vous, entièrement ou en partie, lorsque vous créez votre projet de add-in à l’aide d’un outil tel que Yo Office ou l’un des modèles de Office Add-in dans Visual Studio. Dans ce cas, interprétez la tâche comme si vous devait vérifier qu’elle a été effectuée.

## <a name="use-the-latest-office-javascript-api-library"></a>Utiliser la dernière bibliothèque Office’API JavaScript

Votre application doit charger la version la plus récente de la bibliothèque d’API JavaScript Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, assurez-vous que vous avez la balise suivante dans le premier fichier `script` HTML que votre application ouvre. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>Spécifier les Office applications peuvent héberger votre application

Par défaut, un add-in est installable dans toutes les applications Office pris en charge par le type de add-in spécifié (c’est-à-dire, Courrier, Volet De tâches ou Contenu). Par exemple, un add-in du volet Des tâches est installable par défaut sur Access, Excel, OneNote, PowerPoint, Project et Word. 

Pour vous assurer que votre application est installable dans un sous-ensemble d’applications Office, utilisez les éléments [Hosts](../reference/manifest/hosts.md) et [Host](../reference/manifest/host.md) dans le manifeste.

Par exemple, la déclaration **Hosts** et **Host** suivante spécifie que le add-in peut être installé sur n’importe quelle version de Excel, qui inclut Excel sur le Web, Windows et iPad, mais ne peut pas être installé sur une autre application Office.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

**L’élément Hosts** peut contenir un ou plusieurs **éléments Host.** Il doit y avoir un **élément Host** distinct pour chaque application Office sur laquelle le module doit être installé. `Name`L’attribut est obligatoire et peut être définie sur l’une des valeurs suivantes.

| Nom          | Office applications clientes                     | Types de add-in disponibles |
|:--------------|:-----------------------------------------------|:-----------------------|
| Base de données      | applications web Access                                | Volet de tâches              |
| Document      | Word sur le web, Windows, Mac, iPad            | Volet de tâches              |
| Boîte aux lettres       | Outlook sur le web, Windows, Mac, Android, iOS | Courrier                   |
| Bloc-notes      | OneNote sur le web                             | Volet De tâches, Contenu     |
| Présentation  | PowerPoint sur le web, Windows, Mac, iPad      | Volet De tâches, Contenu     |
| Project       | Project sur Windows                             | Volet de tâches              |
| Classeur      | Excel sur le Web, Windows, Mac, iPad           | Volet De tâches, Contenu     |

> [!NOTE]
> Office applications sont pris en charge sur différentes plateformes et s’exécutent sur des ordinateurs de bureau, des navigateurs web, des tablettes et des appareils mobiles. En règle générale, vous ne pouvez pas spécifier la plateforme qui peut être utilisée pour exécuter votre add-in. Par exemple, si vous spécifiez , les deux Excel sur le Web et sur Windows peuvent être utilisés pour `Workbook` exécuter votre add-in. Toutefois, si vous spécifiez , votre application ne s’exécutera pas sur Outlook clients mobiles, sauf si vous définissez le `Mailbox` [point d’extension mobile.](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface)

> [!NOTE]
> Il n’est pas possible qu’un manifeste de add-in s’applique à plusieurs types : courrier, volet Des tâches ou Contenu. Cela signifie que si vous souhaitez que votre application soit installable sur Outlook et sur l’une des autres applications Office, vous devez créer deux applications, l’une avec un manifeste de type messagerie et l’autre avec un manifeste de type de contenu ou de volet De tâches. 

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>Spécifier les Office et les plateformes qui peuvent héberger votre add-in

Vous ne pouvez pas spécifier explicitement les versions et builds de Office ou les plateformes sur lesquelles votre module doit être installé, et vous ne le souhaiteriez pas, car vous deriez devoir réviser votre manifeste chaque fois que la prise en charge des fonctionnalités de votre add-in est étendue à une nouvelle version ou plateforme. Au lieu de cela, spécifiez dans le manifeste les API dont votre add-in a besoin. Office empêche le module d’être installé sur des combinaisons de version et de plateforme Office qui ne sont pas en charge et garantit que le module n’apparaîtra pas dans Mes **modules.**

> [!IMPORTANT]
> Utilisez uniquement le manifeste de base pour spécifier les membres d’API que votre application doit avoir de toute valeur significative. Si votre application utilise une API pour certaines fonctionnalités, mais qu’elle comporte d’autres fonctionnalités utiles qui ne nécessitent pas l’API, vous devez concevoir le module de sorte qu’il soit installable sur la plateforme et les combinaisons de versions Office qui ne prend pas en charge l’API, mais offre une expérience réduite sur ces combinaisons. Pour plus d’informations, [voir Conception pour d’autres expériences.](#design-for-alternate-experiences)

### <a name="requirement-sets"></a>Ensembles de conditions requises

Pour simplifier le processus de spécification des API requises par votre Office, vous pouvez grouper la plupart des API ensemble dans des ensembles *de conditions requises.* Les API du modèle objet [API](understanding-the-javascript-api-for-office.md#api-models) commun sont regroupées par la fonctionnalité de développement qu’elles supportent. Par exemple, toutes les API connectées à des liaisons de tableau sont dans l’ensemble de conditions requises appelé « TableBindings 1.1 ». Les API dans les modèles objet spécifiques de [l’application](understanding-the-javascript-api-for-office.md#api-models) sont regroupées par date de publication pour être utilisés dans les applications de production.

Les ensembles de conditions requises sont en version. Par exemple, les API qui la prise en charge [des boîtes de](../design/dialog-boxes.md) dialogue sont dans l’ensemble de conditions requises DialogApi 1.1. Lorsque des API supplémentaires qui activent la messagerie à partir d’un volet Des tâches vers une boîte de dialogue ont été publiées, elles ont été regroupées dans DialogApi 1.2, ainsi que toutes les API de DialogApi 1.1. *Chaque version d’un ensemble de conditions requises est un sur-ensemble de toutes les versions antérieures.*

La prise en charge de l’ensemble de conditions requises varie selon Office application, la version de l’application Office et la plateforme sur laquelle elle est en cours d’exécution. Par exemple, DialogApi 1.2 n’est pas pris en charge sur les versions d’achat one-time de Office avant Office 2021, mais DialogApi 1.1 est pris en charge sur toutes les versions d’achat à prix simple antérieures à Office 2013. Vous souhaitez que votre add-in soit installable sur chaque combinaison de plateforme et de version Office qui prend en charge les API qu’il utilise. Vous devez donc toujours spécifier dans le manifeste la version *minimale* de chaque ensemble de conditions requises par votre add-in. Pour plus d’informations sur la façon de le faire, voir plus loin dans cet article.

> [!TIP]
> Pour plus d’informations sur [](office-versions-and-requirement-sets.md#office-requirement-sets-availability)le contrôle de version de l’ensemble de conditions requises, voir Office [la](../reference/requirement-sets/office-add-in-requirement-sets.md)disponibilité des ensembles de conditions requises et pour obtenir la liste complète des ensembles de conditions requises et des informations sur les API dans chacune d’elles, commencez par les ensembles de conditions requises du Office. Les rubriques de référence pour la plupart Office.js API spécifient également l’ensemble de conditions requises à qui elles appartiennent (le cas nécessaire).

> [!NOTE]
> Certains ensembles de conditions requises sont également associés à des éléments de manifeste. Voir Spécification des conditions requises dans un [élément VersionOverrides](#specify-requirements-in-a-versionoverrides-element) pour plus d’informations sur la pertinence de ce fait pour la conception de votre add-in.

#### <a name="apis-not-in-a-requirement-set"></a>API non dans un ensemble de conditions requises

Toutes les API des modèles spécifiques de l’application sont dans des ensembles de conditions requises, mais certaines d’entre elles dans le modèle API commun ne le sont pas. Il existe également un moyen de spécifier l’une de ces API non définies dans le manifeste lorsque votre add-in en requiert une. Cet article contient d’autres détails plus avant.

### <a name="requirements-element"></a>Élément Requirements

Utilisez [l’élément Requirements](../reference/manifest/requirements.md) et ses éléments enfants [Sets](../reference/manifest/sets.md) and Methods pour spécifier les ensembles de conditions [requises](../reference/manifest/methods.md) minimum ou les membres d’API qui doivent être pris en charge par l’application Office pour installer votre application. 

Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API **spécifiés** dans l’élément **Requirements,** le module ne s’exécutera pas dans cette application ou plateforme et ne s’affichera pas dans Mes applications.

> [!NOTE]
> **L’élément Requirements** est facultatif pour tous les modules, à l’exception Outlook les autres. Lorsque l’attribut de l’élément racine est , il doit y avoir un élément Requirements qui spécifie la version minimale de l’ensemble de conditions requises mailBox requise `xsi:type` `OfficeApp` par le `MailBox` module.  Pour plus d’informations, [voir Outlook conditions requises de l’API JavaScript.](../reference/requirement-sets/outlook-api-requirement-sets.md)

L’exemple de code suivant montre comment configurer un add-in installable dans toutes les applications Office qui prendre en charge les applications suivantes :

-  `TableBindings` ensemble de conditions requises, dont la version minimale est « 1.1 ».
-  `OOXML` ensemble de conditions requises, dont la version minimale est « 1.1 ».
-  `Document.getSelectedDataAsync` .

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```
Notez ce qui suit à propos de cet exemple.

- **L’élément Requirements** contient les **éléments enfants Sets** et **Methods.**
- **L’élément Sets** peut contenir un ou plusieurs **éléments Set.** `DefaultMinVersion` spécifie la valeur par `MinVersion` défaut de tous les éléments **Set** enfants.
- Un [élément Set](../reference/manifest/set.md) spécifie un ensemble de conditions requises que l’application Office doit prendre en charge pour rendre le module installable. `Name`L’attribut spécifie le nom de l’ensemble de conditions requises. Spécifie `MinVersion` la version minimale de l’ensemble de conditions requises. `MinVersion`remplace la valeur de `DefaultMinVersion` l’attribut dans les jeux parents .
- **L’élément Methods** peut contenir un ou plusieurs [éléments Method.](../reference/manifest/method.md) Vous ne pouvez pas utiliser l’élément **Methods** avec des compléments Outlook.
- **L’élément** Method spécifie une méthode individuelle que l’application Office doit prendre en charge pour rendre le module installable. `Name`L’attribut est obligatoire et spécifie le nom de la méthode qualifiée avec son objet parent.

## <a name="design-for-alternate-experiences"></a>Conception pour d’autres expériences

Les fonctionnalités d’extensibilité que la plateforme de Office de service fournit peuvent être divisées en trois types :

- Fonctionnalités d’extensibilité disponibles immédiatement après l’installation du module. Vous pouvez utiliser ce type de fonctionnalité en configurant un [élément VersionOverrides](../reference/manifest/versionoverrides.md) dans le manifeste. Les commandes de ce type de fonctionnalité, qui sont des menus et des boutons de ruban [personnalisés,](../design/add-in-commands.md)sont un exemple de ce type de fonctionnalité.
- Fonctionnalités d’extensibilité qui sont disponibles uniquement lorsque le module est en cours d’exécution et qui sont implémentées avec Office.js API JavaScript ; par exemple, [boîtes de dialogue](../design/dialog-boxes.md).
- Fonctionnalités d’extensibilité disponibles uniquement au moment de l’exécution, mais implémentées avec une combinaison de javascript Office.js et de configuration dans un **élément VersionOverrides.** Voici quelques exemples [Excel fonctions personnalisées,](../excel/custom-functions-overview.md)l’personnalisation de l' [sign-on](sso-in-office-add-ins.md)et des [onglets contextuels personnalisés.](../design/contextual-tabs.md)

Si votre add-in utilise une fonctionnalité d’extensibilité spécifique pour certaines de ses fonctionnalités, mais dispose d’autres fonctionnalités utiles qui ne nécessitent pas la fonctionnalité d’extensibilité, vous devez concevoir le module de sorte qu’il soit installable sur les combinaisons de plateforme et de version Office qui ne prend pas en charge la fonctionnalité d’extensibilité. Il peut fournir une expérience précieuse, bien que réduite, sur ces combinaisons. 

Vous implémentez cette conception différemment selon la façon dont la fonctionnalité d’extensibilité est implémentée : 

- Pour les fonctionnalités entièrement implémentées avec JavaScript, voir Vérifications à l’exécution de la prise en charge des méthodes et des [ensembles de conditions requises.](#runtime-checks-for-method-and-requirement-set-support)
- Pour les fonctionnalités qui nécessitent la configuration d’un élément **VersionOverrides,** voir Spécifications requises dans un [élément VersionOverrides.](#specify-requirements-in-a-versionoverrides-element)

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>L’runtime vérifie la prise en charge des méthodes et des ensembles de conditions requises 

Vous testez au moment de l’utilisation pour déterminer si l’Office utilisateur prend en charge un ensemble de conditions requises avec [la méthode isSetSupported.](/javascript/api/office/office.requirementsetsupport#isSetSupported_name__minVersion_) Passez le nom de l’ensemble de conditions requises et la version minimale en tant que paramètres. Si l’ensemble de conditions requises est pris en charge, `isSetSupported` renvoie **true**. Le code ci-dessous vous montre un exemple.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
Tenez compte du code suivant :

- Le premier paramètre est obligatoire. Il s’agit d’une chaîne qui représente le nom de l’ensemble de conditions requises. Pour plus d’informations concernant les ensembles de conditions requises disponibles, voir [Ensembles de conditions requises pour complément Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- Le deuxième paramètre est facultatif. Il s’agit d’une chaîne qui spécifie la version minimale de l’ensemble de conditions requises que l’application Office doit prendre en charge pour que le code de l’instruction s’exécute `if` (par exemple, «**1,9**»). S’il n’est pas utilisé, la version « 1.1 » est supposée.

> [!WARNING]
> Lors de l’appel de la méthode, la valeur du deuxième paramètre (s’il est spécifié) doit être `isSetSupported` une chaîne et non un nombre. L’parseur JavaScript ne peut pas différencier les valeurs numériques telles que 1.1 et 1.10, contrairement aux valeurs de chaînes telles que « 1.1 » et « 1.10 ».

Le tableau suivant indique les noms des ensembles de conditions requises pour les modèles d’API spécifiques à l’application.

|Application Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Boîte aux lettres|
|PowerPoint|PowerPointApi|
|Word|WordApi|

Voici un exemple d’utilisation de la méthode avec l’un des ensembles de conditions requises du modèle API commun.

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> La méthode et les ensembles de conditions requises pour ces applications sont disponibles dans le dernier `isSetSupported` fichier Office.js sur le CDN. Si vous n’utilisez pas Office.js du CDN, votre module peut générer des exceptions si vous utilisez une ancienne version de la bibliothèque dans laquelle il n’est pas `isSetSupported` définie. Pour plus d’informations, [voir Utiliser la dernière Office de l’API JavaScript.](#use-the-latest-office-javascript-api-library)

Lorsque votre application dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, utilisez la vérification à l’runtime pour déterminer si la méthode est prise en charge par l’application Office, comme illustré dans l’exemple de code suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’application Office prend en charge `document.setSelectedDataAsync` .

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>Spécifier les conditions requises dans un élément VersionOverrides

[L’élément VersionOverrides](../reference/manifest/versionoverrides.md) a été ajouté au schéma de manifeste principalement, mais pas exclusivement, pour prendre en charge les fonctionnalités qui doivent être disponibles immédiatement après l’installation d’un module, telles que les commandes de add-in (boutons et menus personnalisés du ruban). Office connaître ces fonctionnalités lorsqu’il pare le manifeste du add-in. 

Supposons que votre add-in utilise l’une de ces fonctionnalités, mais qu’il est utile et qu’il doit être installable, même sur les versions Office qui ne la prend pas en charge. Dans ce scénario, identifiez la fonctionnalité à l’aide d’un élément [Requirements](../reference/manifest/requirements.md) (et de ses éléments [Sets](../reference/manifest/sets.md) et [Methods](../reference/manifest/methods.md) enfants) que vous incluez en tant qu’enfant de l’élément **VersionOverrides lui-même** plutôt qu’en tant qu’enfant de l’élément de `OfficeApp` base. Cela a pour effet que Office autorise l’installation du module, mais Office ignore certains des éléments enfants de l’élément **VersionOverrides** sur les versions Office où la fonctionnalité n’est pas prise en charge.

Plus précisément, les éléments enfants de **VersionOverrides** qui remplacent les éléments dans le manifeste de base, tels qu’un élément **Hosts,** sont ignorés et les éléments correspondants du manifeste de base sont utilisés à la place. Toutefois, il peut y avoir des éléments enfants dans **une VersionOverrides** qui implémentent des fonctionnalités supplémentaires plutôt que de remplacer les paramètres dans le manifeste de base. Deux exemples sont les `WebApplicationInfo` `EquivalentAddins` suivants : Ces parties de **VersionOverrides** ne seront pas ignorées, en supposant que la plateforme et la version de Office la fonctionnalité correspondante.   

Pour plus d’informations sur les éléments descendants de l’élément **Requirements,** voir [l’élément Requirements](#requirements-element) plus tôt dans cet article.

Voici un exemple.

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> Faites très attention avant d’utiliser un élément **Requirements** dans **une VersionOverrides,** car sur les combinaisons de plateforme et de version qui ne prend pas en charge la condition *requise,* aucune des commandes de module ne sera installée, même celles qui appellent des fonctionnalités qui *n’ont* pas besoin de cette exigence. Prenons l’exemple d’un add-in qui possède deux boutons de ruban personnalisés. L’un d’eux Office des API JavaScript disponibles dans l’ensemble de conditions **requises ExcelApi 1.4** (et version ultérieure). Les autres appellent des API qui sont uniquement disponibles dans **ExcelApi 1.9** (et ultérieur). Si vous avez placé une condition requise pour **ExcelApi 1.9** dans **VersionOverrides,** alors lorsque la version 1.9 n’est pas prise en charge, aucun bouton n’apparaît sur le ruban.  Une meilleure stratégie dans ce scénario consisterait à utiliser la technique décrite dans les vérifications runtime pour la prise en charge des méthodes et des ensembles [de conditions requises.](#runtime-checks-for-method-and-requirement-set-support) Le code appelé par le deuxième bouton utilise d’abord pour vérifier la prise en charge `isSetSupported` **d’ExcelApi 1.9**. S’il n’est pas pris en charge, le code envoie à l’utilisateur un message lui disant que cette fonctionnalité du module n’est pas disponible sur sa version de Office. 

> [!TIP]
> Il n’est pas nécessaire de répéter un élément **Requirement** dans **une versionOverrides** qui apparaît déjà dans le manifeste de base. Si l’exigence est spécifiée dans le manifeste de base, le add-in ne peut pas s’installer lorsque la condition n’est pas prise en charge, de sorte que Office n’a pas même l’élément **VersionOverrides.** 

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](add-in-manifests.md)
- [Ensembles de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
