---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: Découvrez comment spécifier les applications Office et les exigences d’API pour que votre complément fonctionne comme prévu.
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: f4953d18a0d9d09c7f7e15a5fdfbad8525a1fdb8
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68466907"
---
# <a name="specify-office-applications-and-api-requirements"></a>Spécifier les applications Office et les exigences de l’API

Votre complément Office peut dépendre d’une application Office spécifique (également appelée hôte Office) ou de membres spécifiques de l’API JavaScript Office (office.js). Par exemple, votre complément peut :

- Exécuter dans une application Office (Word ou Excel), ou plusieurs applications.
- Utilisez les API JavaScript Office disponibles uniquement dans certaines versions d’Office. Par exemple, la version perpétuelle sous licence en volume de Excel 2016 ne prend pas en charge toutes les API liées à Excel dans la bibliothèque JavaScript Office.

Dans ces situations, vous devez vous assurer que votre complément n’est jamais installé sur les applications Office ou les versions d’Office dans lesquelles il ne peut pas s’exécuter.

Il existe également des scénarios dans lesquels vous souhaitez contrôler les fonctionnalités de votre complément qui sont visibles par les utilisateurs en fonction de leur application Office et de leur version d’Office. Voici deux exemples :

- Votre complément comporte des fonctionnalités utiles dans Word et PowerPoint, telles que la manipulation de texte, mais il a des fonctionnalités supplémentaires qui n’ont de sens que dans PowerPoint, telles que les fonctionnalités de gestion des diapositives. Vous devez masquer les fonctionnalités PowerPoint uniquement lorsque le complément est en cours d’exécution dans Word.
- Votre complément dispose d’une fonctionnalité qui nécessite une méthode d’API JavaScript Office prise en charge dans certaines versions d’une application Office, comme l’abonnement Microsoft 365 Excel, mais qui n’est pas prise en charge dans d’autres, telles que les Excel 2016 perpétuelles sous licence en volume. Toutefois, votre complément a d’autres fonctionnalités qui nécessitent uniquement des méthodes d’API JavaScript *Office prises* en charge dans les Excel 2016 perpétuelles sous licence en volume. Dans ce scénario, vous avez besoin que le complément soit installable sur cette version de Excel 2016, mais la fonctionnalité qui nécessite la méthode non prise en charge doit être masquée pour ces utilisateurs.

Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

> [!NOTE]
> Pour obtenir une vue générale de l’emplacement où les compléments Office sont actuellement pris en charge, consultez la page [Sur la disponibilité de l’application cliente et de la plateforme Office pour les compléments Office](/javascript/api/requirement-sets) .

> [!TIP]
> La plupart des tâches décrites dans cet article sont effectuées pour vous, en totalité ou en partie, lorsque vous créez votre projet de complément avec un outil, tel que le [générateur Yeoman pour les compléments Office](yeoman-generator-overview.md) ou l’un des modèles de complément Office dans Visual Studio. Dans ce cas, veuillez interpréter la tâche comme indiquant que vous devez vérifier qu’elle a été effectuée.

## <a name="use-the-latest-office-javascript-api-library"></a>Utiliser la dernière bibliothèque d’API JavaScript Office

Votre complément doit charger la version la plus récente de la bibliothèque d’API JavaScript Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, assurez-vous que vous avez la balise suivante `script` dans le premier fichier HTML que votre complément s’ouvre. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>Spécifier les applications Office pouvant héberger votre complément

Par défaut, un complément est installable dans toutes les applications Office prises en charge par le type de complément spécifié (courrier, volet Office ou contenu). Par exemple, un complément du volet Office est installable par défaut sur Access, Excel, OneNote, PowerPoint, Project et Word.

Pour vous assurer que votre complément est installable dans un sous-ensemble d’applications Office, utilisez les éléments [Hosts](/javascript/api/manifest/hosts) et [Host](/javascript/api/manifest/host) dans le manifeste.

Par exemple, la déclaration suivante **\<Hosts\>** **\<Host\>** spécifie que le complément peut être installé sur n’importe quelle version d’Excel, qui inclut Excel sur le Web, Windows et iPad, mais ne peut pas être installé sur une autre application Office.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

L’élément **\<Hosts\>** peut contenir un ou plusieurs **\<Host\>** éléments. Il doit y avoir un élément distinct **\<Host\>** pour chaque application Office sur laquelle le complément doit être installable. L’attribut `Name` est obligatoire et peut être défini sur l’une des valeurs suivantes.

| Nom          | Applications clientes Office                     | Types de compléments disponibles |
|:--------------|:-----------------------------------------------|:-----------------------|
| Base de données      | applications web Access                                | Volet de tâches              |
| Document      | Word sur le web, Windows, Mac, iPad            | Volet de tâches              |
| Boîte aux lettres       | Outlook sur le web, Windows, Mac, Android, iOS | Courrier                   |
| Bloc-notes      | OneNote sur le web                             | Volet Office, Contenu     |
| Présentation  | PowerPoint sur le web, Windows, Mac, iPad      | Volet Office, Contenu     |
| Project       | Project sur Windows                             | Volet de tâches              |
| Classeur      | Excel sur le Web, Windows, Mac, iPad           | Volet Office, Contenu     |

> [!NOTE]
> Les applications Office sont prises en charge sur différentes plateformes et exécutées sur des ordinateurs de bureau, des navigateurs web, des tablettes et des appareils mobiles. Vous ne pouvez généralement pas spécifier la plateforme qui peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez `Workbook`, Excel sur le Web et sur Windows peuvent être utilisés pour exécuter votre complément. Toutefois, si vous spécifiez `Mailbox`, votre complément ne s’exécutera pas sur les clients mobiles Outlook, sauf si vous définissez le [point d’extension mobile](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface).

> [!NOTE]
> Il n’est pas possible qu’un manifeste de complément s’applique à plusieurs types : Courrier, volet Office ou Contenu. Cela signifie que si vous souhaitez que votre complément soit installable sur Outlook et sur l’une des autres applications Office, vous devez créer *deux* compléments, l’un avec un manifeste de type Courrier et l’autre avec un volet Office ou un manifeste de type contenu.

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>Spécifier les versions et plateformes Office pouvant héberger votre complément

Vous ne pouvez pas spécifier explicitement les versions et builds Office ou les plateformes sur lesquelles votre complément doit être installable, et vous ne le souhaitez pas, car vous devez réviser votre manifeste chaque fois que la prise en charge des fonctionnalités de complément utilisées par votre complément est étendue à une nouvelle version ou plateforme. Au lieu de cela, spécifiez dans le manifeste les API dont votre complément a besoin. Office empêche l’installation du complément sur des combinaisons de versions et de plateformes Office qui ne prennent pas en charge les API et garantit que le complément n’apparaît pas dans **Mes compléments**.

> [!IMPORTANT]
> Utilisez uniquement le manifeste de base pour spécifier les membres de l’API dont votre complément doit avoir une valeur significative. Si votre complément utilise une API pour certaines fonctionnalités, mais possède d’autres fonctionnalités utiles qui ne nécessitent pas l’API, vous devez concevoir le complément afin qu’il soit installable sur des combinaisons de versions de plateforme et Office qui ne prennent pas en charge l’API, mais offrent une expérience réduite sur ces combinaisons. Pour plus d’informations, consultez [Conception pour d’autres expériences](#design-for-alternate-experiences).

### <a name="requirement-sets"></a>Ensembles de conditions requises

Pour simplifier le processus de spécification des API dont votre complément a besoin, Office regroupe la plupart des API dans des *ensembles de conditions requises*. Les API du [modèle objet d’API commune](understanding-the-javascript-api-for-office.md#api-models) sont regroupées par la fonctionnalité de développement qu’elles prennent en charge. Par exemple, toutes les API connectées aux liaisons de table se trouvent dans l’ensemble de conditions requises appelé « TableBindings 1.1 ». Les API des [modèles objet spécifiques à l’application](understanding-the-javascript-api-for-office.md#api-models) sont regroupées lorsqu’elles ont été publiées pour être utilisées dans les compléments de production.

Les ensembles de conditions requises sont versionnées. Par exemple, les API qui prennent en charge [les boîtes de dialogue](../develop/dialog-api-in-office-add-ins.md) se trouvent dans l’ensemble de conditions requises DialogApi 1.1. Lorsque des API supplémentaires qui activent la messagerie d’un volet Office vers une boîte de dialogue ont été publiées, elles ont été regroupées dans DialogApi 1.2, ainsi que toutes les API dans DialogApi 1.1. *Chaque version d’un ensemble de conditions requises est un sur-ensemble de toutes les versions antérieures.*

La prise en charge de l’ensemble de conditions requises varie en fonction de l’application Office, de la version de l’application Office et de la plateforme sur laquelle elle s’exécute. Par exemple, DialogApi 1.2 n’est pas pris en charge sur les versions perpétuelles avec licence en volume d’Office avant Office 2021, mais DialogApi 1.1 est pris en charge sur toutes les versions perpétuelles d’Office 2013. Vous souhaitez que votre complément soit installable sur chaque combinaison de plateforme et de version d’Office qui prend en charge les API qu’il utilise. Vous devez donc toujours spécifier dans le manifeste la version *minimale* de chaque ensemble de conditions requises requise par votre complément. Vous trouverez plus d’informations sur la procédure à suivre plus loin dans cet article.

> [!TIP]
> Pour plus d’informations sur le contrôle de version des ensembles de conditions requises, consultez [la disponibilité des ensembles de conditions requises Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) et, pour obtenir des listes complètes d’ensembles de conditions requises et des informations sur les API de chacun d’eux, commencez par les [ensembles de conditions requises du complément Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets). Les rubriques de référence de la plupart des API Office.js spécifient également l’ensemble de conditions requises auquel elles appartiennent (le cas échéant).

> [!NOTE]
> Certains ensembles de conditions requises ont également des éléments manifestes associés. Consultez [Spécification des exigences dans un élément VersionOverrides](#specify-requirements-in-a-versionoverrides-element) pour plus d’informations sur le moment où ce fait est pertinent pour votre conception de complément.

#### <a name="apis-not-in-a-requirement-set"></a>API non dans un ensemble de conditions requises

Toutes les API des modèles spécifiques à l’application se trouvent dans des ensembles de conditions requises, mais certaines d’entre elles ne le sont pas dans le modèle d’API commun. Il existe également un moyen de spécifier l’une de ces API sans définition dans le manifeste lorsque votre complément en a besoin. Cet article contient d’autres détails plus avant.

### <a name="requirements-element"></a>Élément Requirements

Utilisez l’élément [Requirements](/javascript/api/manifest/requirements) et ses éléments enfants [Sets](/javascript/api/manifest/sets) and Methods pour spécifier les [ensembles](/javascript/api/manifest/methods) de conditions requises minimum ou les membres d’API qui doivent être pris en charge par l’application Office pour installer votre complément.

Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément **\<Requirements\>** , le complément ne s’exécute pas dans cette application ou cette plateforme et ne s’affiche pas dans **Mes compléments**.

> [!NOTE]
> L’élément **\<Requirements\>** est facultatif pour tous les compléments, à l’exception des compléments Outlook. Lorsque l’attribut `xsi:type` de l’élément racine `OfficeApp` est `MailApp`, il doit y avoir un **\<Requirements\>** élément qui spécifie la version minimale de l’ensemble de conditions requises de boîte aux lettres requise par le complément. Pour plus d’informations, consultez [les ensembles de conditions requises de l’API JavaScript Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

L’exemple de code suivant montre comment configurer un complément installable dans toutes les applications Office qui prennent en charge les éléments suivants :

- `TableBindings` l’ensemble de conditions requises, qui a une version minimale de « 1.1 ».
- `OOXML` l’ensemble de conditions requises, qui a une version minimale de « 1.1 ».
- `Document.getSelectedDataAsync` Méthode.

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

- L’élément **\<Requirements\>** contient les éléments enfants et **\<Methods\>** les **\<Sets\>** éléments.
- L’élément **\<Sets\>** peut contenir un ou plusieurs **\<Set\>** éléments. `DefaultMinVersion` spécifie la valeur par défaut `MinVersion` de tous les éléments enfants **\<Set\>** .
- Un élément [Set](/javascript/api/manifest/set) spécifie un ensemble de conditions requises que l’application Office doit prendre en charge pour rendre le complément installable. L’attribut `Name` spécifie le nom de l’ensemble de conditions requises. Spécifie `MinVersion` la version minimale de l’ensemble de conditions requises. `MinVersion` remplace la valeur de l’attribut `DefaultMinVersion` dans le parent **\<Sets\>**.
- L’élément **\<Methods\>** peut contenir un ou plusieurs éléments [Method](/javascript/api/manifest/method) . Vous ne pouvez pas utiliser l’élément **\<Methods\>** avec des compléments Outlook.
- L’élément **\<Method\>** spécifie une méthode individuelle que l’application Office doit prendre en charge pour rendre le complément installable. L’attribut `Name` est obligatoire et spécifie le nom de la méthode qualifiée avec son objet parent.

## <a name="design-for-alternate-experiences"></a>Conception pour d’autres expériences

Les fonctionnalités d’extensibilité fournies par la plateforme de complément Office peuvent être utilement divisées en trois types :

- Fonctionnalités d’extensibilité disponibles immédiatement après l’installation du complément. Vous pouvez utiliser ce type de fonctionnalité en configurant un élément [VersionOverrides](/javascript/api/manifest/versionoverrides) dans le manifeste. Les commandes de complément, qui sont [des boutons](../design/add-in-commands.md) et des menus personnalisés du ruban, sont un exemple de ce type de fonctionnalité.
- Fonctionnalités d’extensibilité qui sont disponibles uniquement lorsque le complément est en cours d’exécution et qui sont implémentées avec Office.js API JavaScript ; par exemple, [boîtes de dialogue](../develop/dialog-api-in-office-add-ins.md).
- Fonctionnalités d’extensibilité disponibles uniquement au moment de l’exécution, mais implémentées avec une combinaison de Office.js JavaScript et de configuration dans un **\<VersionOverrides\>** élément. Les [fonctions personnalisées Excel](../excel/custom-functions-overview.md), l’authentification [unique](sso-in-office-add-ins.md) et les [onglets contextuels personnalisés en sont des exemples](../design/contextual-tabs.md).

Si votre complément utilise une fonctionnalité d’extensibilité spécifique pour certaines de ses fonctionnalités, mais possède d’autres fonctionnalités utiles qui ne nécessitent pas la fonctionnalité d’extensibilité, vous devez concevoir le complément afin qu’il soit installable sur les combinaisons de versions de plateforme et d’Office qui ne prennent pas en charge la fonctionnalité d’extensibilité. Il peut fournir une expérience précieuse, bien que diminuée, sur ces combinaisons.

Vous implémentez cette conception différemment en fonction de la façon dont la fonctionnalité d’extensibilité est implémentée :

- Pour les fonctionnalités entièrement implémentées avec JavaScript, consultez [les vérifications d’exécution pour la prise en charge de la méthode et de l’ensemble de conditions requises](#runtime-checks-for-method-and-requirement-set-support).
- Pour les fonctionnalités qui vous obligent à configurer un **\<VersionOverrides\>** élément, consultez [Spécification des exigences dans un élément VersionOverrides](#specify-requirements-in-a-versionoverrides-element).

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>Vérifications d’exécution de la prise en charge de la méthode et de l’ensemble de conditions requises

Vous testez au moment de l’exécution pour déterminer si Office de l’utilisateur prend en charge un ensemble de conditions requises avec la méthode [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) . Transmettez le nom de l’ensemble de conditions requises et la version minimale en tant que paramètres. Si l’ensemble de conditions requises est pris en charge, `isSetSupported` retourne `true`. Le code ci-dessous vous montre un exemple.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is perpetual Word 2013 (which does not support WordApi 1.1).
}
```

Tenez compte du code suivant :

- Le premier paramètre est obligatoire. Il s’agit d’une chaîne qui représente le nom de l’ensemble de conditions requises. Pour plus d’informations concernant les ensembles de conditions requises disponibles, voir [Ensembles de conditions requises pour complément Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets).
- Le deuxième paramètre est facultatif. Il s’agit d’une chaîne qui spécifie la version minimale de l’ensemble de conditions requises que l’application Office doit prendre en charge pour que le code de l’instruction `if` s’exécute (par exemple, « **1,9** »). S’il n’est pas utilisé, la version « 1.1 » est supposée.

> [!WARNING]
> Lors de l’appel de la `isSetSupported` méthode, la valeur du deuxième paramètre (si spécifié) doit être une chaîne et non un nombre. L’analyseur JavaScript ne peut pas différencier les valeurs numériques telles que 1.1 et 1.10, tandis qu’il le peut pour les valeurs de chaîne telles que « 1.1 » et « 1.10 ».

Le tableau suivant présente les noms des ensembles de conditions requises pour les modèles d’API spécifiques à l’application.

|Application Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Boîte aux lettres|
|PowerPoint|PowerPointApi|
|Word|WordApi|

Voici un exemple d’utilisation de la méthode avec l’un des ensembles de conditions requises du modèle d’API commun.

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
> La `isSetSupported` méthode et les ensembles de conditions requises pour ces applications sont disponibles dans le dernier fichier Office.js sur le CDN. Si vous n’utilisez pas Office.js à partir du CDN, votre complément peut générer des exceptions si vous utilisez une ancienne version de la bibliothèque dans laquelle `isSetSupported` il n’est pas défini. Pour plus d’informations, consultez [Utiliser la dernière bibliothèque d’API JavaScript Office](#use-the-latest-office-javascript-api-library).

Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, utilisez la vérification de l’exécution pour déterminer si la méthode est prise en charge par l’application Office, comme illustré dans l’exemple de code suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’application Office prend en charge `document.setSelectedDataAsync`.

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>Spécifier des exigences dans un élément VersionOverrides

L’élément [VersionOverrides](/javascript/api/manifest/versionoverrides) a été ajouté au schéma de manifeste principalement, mais pas exclusivement, pour prendre en charge les fonctionnalités qui doivent être disponibles immédiatement après l’installation d’un complément, telles que les commandes de complément (boutons et menus personnalisés du ruban). Office doit connaître ces fonctionnalités lorsqu’il analyse le manifeste du complément.

Supposons que votre complément utilise l’une de ces fonctionnalités, mais que le complément est précieux et qu’il doit être installable, même sur les versions d’Office qui ne prennent pas en charge la fonctionnalité. Dans ce scénario, identifiez la fonctionnalité à l’aide d’un élément [Requirements](/javascript/api/manifest/requirements) (et de ses éléments [enfants Sets](/javascript/api/manifest/sets) and [Methods](/javascript/api/manifest/methods) ) que vous incluez en tant qu’enfant de l’élément **\<VersionOverrides\>** lui-même plutôt qu’en tant qu’enfant de l’élément de base `OfficeApp` . Cela a pour effet qu’Office autorise l’installation du complément, mais Office ignore certains éléments enfants de l’élément **\<VersionOverrides\>** dans les versions d’Office où la fonctionnalité n’est pas prise en charge.

Plus précisément, les éléments enfants des **\<VersionOverrides\>** éléments qui remplacent les éléments dans le manifeste de base, tels qu’un **\<Hosts\>** élément, sont ignorés et les éléments correspondants du manifeste de base sont utilisés à la place. Toutefois, il peut y avoir des éléments enfants dans un **\<VersionOverrides\>** élément qui implémentent réellement des fonctionnalités supplémentaires plutôt que des paramètres de remplacement dans le manifeste de base. Deux exemples sont le `WebApplicationInfo` et `EquivalentAddins`. Ces parties ne **\<VersionOverrides\>** seront *pas* ignorées, en supposant que la plateforme et la version d’Office prennent en charge la fonctionnalité correspondante.  

Pour plus d’informations sur les éléments descendants de l’élément, consultez l’élément **\<Requirements\>** [Requirements](#requirements-element) plus haut dans cet article.

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
> Utilisez un grand soin avant d’utiliser un **\<Requirements\>** élément dans un **\<VersionOverrides\>**, car sur les combinaisons de plateformes et de versions qui ne prennent pas en charge l’exigence, *aucune* des commandes de complément n’est installée, *même celles qui appellent des fonctionnalités qui n’en ont pas besoin*. Prenons l’exemple d’un complément qui a deux boutons de ruban personnalisés. L’un d’eux appelle les API JavaScript Office disponibles dans l’ensemble de conditions requises **ExcelApi 1.4** (et versions ultérieures). Les autres appellent des API disponibles uniquement dans **ExcelApi 1.9** (et versions ultérieures). Si vous placez une exigence pour **ExcelApi 1.9** dans le **\<VersionOverrides\>**, quand 1.9 n’est pas pris en charge *, aucun bouton n’apparaît* sur le ruban. Une meilleure stratégie dans ce scénario serait d’utiliser la technique décrite dans [runtime checks for method and requirement set support](#runtime-checks-for-method-and-requirement-set-support). Le code appelé par le deuxième bouton permet `isSetSupported` d’abord de vérifier la prise en charge **d’ExcelApi 1.9**. S’il n’est pas pris en charge, le code fournit à l’utilisateur un message indiquant que cette fonctionnalité du complément n’est pas disponible sur sa version d’Office.

> [!TIP]
> Il est inutile de répéter un élément **Requirement** dans un **\<VersionOverrides\>** élément qui apparaît déjà dans le manifeste de base. Si l’exigence est spécifiée dans le manifeste de base, le complément ne peut pas s’installer là où l’exigence n’est pas prise en charge, de sorte qu’Office n’analyse même pas l’élément **\<VersionOverrides\>** .

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](add-in-manifests.md)
- [Ensembles de conditions requises pour les compléments Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
