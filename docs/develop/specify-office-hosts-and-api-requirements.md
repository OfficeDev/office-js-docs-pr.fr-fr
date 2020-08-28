---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: Découvrez comment spécifier les applications Office et les conditions requises de l’API pour que votre complément fonctionne comme prévu.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 90ee7c3a5ad01252336608c02f995bbcbbe94212
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292628"
---
# <a name="specify-office-applications-and-api-requirements"></a>Spécification des exigences en matière d’applications et d’API Office

Votre complément Office peut dépendre d’une application Office spécifique, d’un ensemble de conditions requises, d’un membre de l’API ou d’une version de l’API pour fonctionner comme prévu. Par exemple, votre complément peut :

- Exécuter dans une application Office (Word ou Excel), ou plusieurs applications.

- utiliser des API JavaScript disponibles uniquement dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui fonctionne dans Excel 2016 ;

- s’exécuter uniquement dans les versions d’Office qui prennent en charge les membres d’API utilisés par votre complément.

Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

> [!NOTE]
> Pour obtenir une vue d’ensemble de l’emplacement où les compléments Office sont actuellement pris en charge, consultez la page [application cliente Office et disponibilité de la plateforme pour les compléments Office](../overview/office-add-in-availability.md) .

Le tableau suivant répertorie les concepts principaux décrits dans cet article.

|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application cliente Office|Application Office utilisée pour exécuter votre complément. Par exemple, Word, Excel, etc.|
|Plateforme|Emplacement d’exécution de l’application Office, par exemple dans un navigateur ou sur un iPad.|
|Ensemble de conditions requises|Groupe nommé de membres d’API associés. Les compléments utilisent des ensembles de conditions requises pour déterminer si l’application Office prend en charge les membres d’API utilisés par votre complément. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres individuels d’API. La prise en charge de l’ensemble de conditions requises varie en fonction de l’application Office et de la version de l’application Office. <br >Les ensembles de conditions requises sont spécifiés dans le fichier manifeste. Lorsque vous spécifiez des ensembles de conditions requises dans le manifeste, vous définissez le niveau minimal de prise en charge de l’API que l’application Office doit fournir afin d’exécuter votre complément. Les applications Office qui ne prennent pas en charge les ensembles de conditions requises spécifiés dans le manifeste ne peuvent pas exécuter votre complément et votre complément ne s’affichera pas dans <span class="ui">mes compléments</span>. Cela limite l’emplacement où votre complément est disponible. Dans le code utilisant les vérifications à l’exécution. Pour obtenir la liste complète des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md).|
|Vérification à l’exécution|Test effectué au moment de l’exécution pour déterminer si l’application Office qui exécute votre complément prend en charge les ensembles de conditions requises ou les méthodes utilisées par votre complément. Pour effectuer une vérification à l’exécution, vous utilisez une instruction **If** avec la `isSetSupported` méthode, les ensembles de conditions requises ou les noms de méthodes qui ne font pas partie d’un ensemble de conditions requises. Les vérifications à l’exécution permettent de veiller à ce que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’exécution ne spécifient pas le niveau minimal de prise en charge de l’API que l’application Office doit fournir pour l’exécution de votre complément. Au lieu de cela, vous utilisez l’instruction **If** pour déterminer si un membre de l’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affiche toujours dans **Mes compléments** quand vous effectuez des vérifications à l’exécution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez des vérifications à l’exécution dans votre complément, assurez-vous d’utiliser la dernière bibliothèque de l’API JavaScript pour Office (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du schéma de manifeste de complément le plus récent

Le manifeste de votre du complément doit utiliser la version 1.1 du schéma de manifeste de complément. Définissez l' `OfficeApp` élément dans le manifeste de votre complément comme suit.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>Spécifier la dernière bibliothèque d’API JavaScript pour Office

Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de la bibliothèque de l’API JavaScript pour Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante à votre code HTML. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>Options permettant de spécifier les applications Office ou les conditions requises d’API

Lorsque vous spécifiez des applications Office ou des exigences d’API, il existe plusieurs facteurs à prendre en compte. Le diagramme suivant montre comment choisir la technique à utiliser dans votre complément.

![Choisir la meilleure option pour votre complément lorsque vous spécifiez les applications Office ou les conditions requises d’API](../images/options-for-office-hosts.png)

- Si votre complément s’exécute dans une application Office, définissez l' `Hosts` élément dans le manifeste. Pour plus d’informations, consultez [Définition de l’élément Hosts](#set-the-hosts-element).

- Pour définir l’ensemble de conditions minimales ou les membres de l’API qu’une application Office doit prendre en charge pour exécuter votre complément, définissez l' `Requirements` élément dans le manifeste. Pour plus d’informations, consultez la section [ Définition de l’élément Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).

- Si vous souhaitez fournir des fonctionnalités supplémentaires si des ensembles de conditions requises ou des membres d’API spécifiques sont disponibles dans l’application Office, effectuez une vérification à l’exécution dans le code JavaScript de votre complément. Par exemple, si votre complément est exécuté dans Excel 2016, utilisez les membres d’API de l’API JavaScript Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, consultez la section [Utilisation des vérifications à l’exécution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).

## <a name="set-the-hosts-element"></a>Définition de l’élément Hosts

Pour que votre complément s’exécute dans une application cliente Office, utilisez les `Hosts` `Host` éléments et dans le manifeste. Si vous ne spécifiez pas l' `Hosts` élément, votre complément s’exécutera dans toutes les applications Office prises en charge par les compléments Office.

Par exemple, la `Hosts` déclaration et suivante `Host` spécifie que le complément fonctionnera avec n’importe quelle version d’Excel, y compris Excel sur le Web, Windows et iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

L' `Hosts` élément peut contenir un ou plusieurs `Host` éléments. L' `Host` élément spécifie l’application Office requise par votre complément. L' `Name` attribut est obligatoire et peut prendre la valeur de l’une des valeurs suivantes.

| Nom          | Applications clientes Office                      |
|:--------------|:----------------------------------------------|
| Base de données      | applications web Access                               |
| Document      | Word sur le Web, Windows, Mac, iPad           |
| Boîte aux lettres       | Outlook sur le Web, Windows, Mac, Android, iOS|
| Présentation  | PowerPoint sur le Web, Windows, Mac, iPad     |
| Project       | Project sur Windows                            |
| Classeur      | Excel sur le Web, Windows, Mac, iPad          |

> [!NOTE]
> L' `Name` attribut spécifie l’application cliente Office qui peut exécuter votre complément. Les applications Office sont prises en charge sur différentes plateformes et s’exécutent sur des ordinateurs de bureau, des navigateurs Web, des tablettes et des appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez `Mailbox` , vous pouvez utiliser Outlook sur le Web et Windows pour exécuter votre complément.

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

## <a name="set-the-requirements-element-in-the-manifest"></a>Définition de l’élément Requirements dans le manifeste

L' `Requirements` élément spécifie les ensembles de conditions requises minimum ou les membres d’API qui doivent être pris en charge par l’application Office pour exécuter votre complément. L' `Requirements` élément peut spécifier des ensembles de conditions requises et des méthodes individuelles utilisées dans votre complément. Dans la version 1,1 du schéma de manifeste de complément, l' `Requirements` élément est facultatif pour tous les compléments, à l’exception des compléments Outlook.

> [!WARNING]
> Utilisez uniquement l' `Requirements` élément pour spécifier des ensembles de conditions requises critiques ou des membres d’API que votre complément doit utiliser. Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l' `Requirements` élément, le complément ne s’exécutera pas dans cette application ou cette plateforme, et ne s’affichera pas dans **mes compléments**. Au lieu de cela, nous vous recommandons de faire en sorte que votre complément soit disponible sur toutes les plateformes d’une application Office, comme Excel sur le Web, Windows et iPad. Pour que votre complément soit disponible sur  _toutes les_ applications et plateformes Office, utilisez des vérifications à l’exécution à la place de l' `Requirements` élément.

L’exemple de code suivant montre un complément qui se charge dans toutes les applications clientes Office qui prennent en charge les éléments suivants :

-  `TableBindings` ensemble de conditions requises, dont la version minimale est « 1,1 ».

-  `OOXML` ensemble de conditions requises, dont la version minimale est « 1,1 ».

-  `Document.getSelectedDataAsync` procédé.

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- L' `Requirements` élément contient les `Sets` `Methods` éléments enfants et.

- L' `Sets` élément peut contenir un ou plusieurs `Set` éléments. `DefaultMinVersion` spécifie la `MinVersion` valeur par défaut de tous les `Set` éléments enfants.

- L' `Set` élément spécifie les ensembles de conditions requises que l’application Office doit prendre en charge pour exécuter le complément. L' `Name` attribut spécifie le nom de l’ensemble de conditions requises. L' `MinVersion` spécifie la version minimale de l’ensemble de conditions requises. `MinVersion` remplace la valeur de `DefaultMinVersion` pour plus d’informations sur les ensembles de conditions requises et les versions d’ensemble de conditions requises auxquelles appartiennent les membres de l’API, consultez la rubrique [ensembles de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md).

- L' `Methods` élément peut contenir un ou plusieurs `Method` éléments. Vous ne pouvez pas utiliser l' `Methods` élément avec des compléments Outlook.

- L' `Method` élément spécifie une méthode individuelle qui doit être prise en charge dans l’application Office dans laquelle votre complément est exécuté. L' `Name` attribut est obligatoire et spécifie le nom de la méthode qualifiée avec son objet parent.

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des vérifications à l’exécution dans votre code JavaScript

Vous souhaiterez peut-être fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’application Office. Par exemple, vous pouvez utiliser les interfaces API JavaScript de Word dans votre complément existant si ce dernier est exécuté dans Word 2016. Pour ce faire, utilisez la méthode [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) avec le nom de l’ensemble de conditions requises. `isSetSupported` détermine, au moment de l’exécution, si l’application Office qui exécute le complément prend en charge l’ensemble de conditions requises. Si l’ensemble de conditions requises est pris en charge, `isSetSupported` renvoie la **valeur true** et exécute le code supplémentaire qui utilise les membres de l’API à partir de cet ensemble de conditions requises. Si l’application Office ne prend pas en charge l’ensemble de conditions requises, `isSetSupported` renvoie la **valeur false** et le code supplémentaire ne s’exécute pas. Le code suivant indique la syntaxe à utiliser avec `isSetSupported`

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (obligatoire) est une chaîne qui représente le nom de l’ensemble de la configuration requise (p. ex., « **ExcelApi** », « **Mailbox** », etc.). Pour plus d’informations concernant les ensembles de conditions requises disponibles, voir [Ensembles de conditions requises pour complément Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- _MinimumVersion_ (facultatif) est une chaîne qui spécifie la version minimale de l’ensemble de conditions requises que l’application Office doit prendre en charge afin que le code au sein de l' `if` instruction s’exécute (par exemple, «**1,9**»).

> [!WARNING]
> Lors de l’appel de la `isSetSupported` méthode, la valeur du `MinimumVersion` paramètre (s’il est spécifié) doit être une chaîne. En effet, l’analyseur syntaxique JavaScript ne peut pas différencier les valeurs numériques, telles que 1.1 et 1.10, mais le peut pour les valeurs chaîne, telles que « 1.1 » et « 1.10 ».
> La surcharge `number` est déconseillée.

Utilisez- `isSetSupported` le avec l' `RequirementSetName` application Office, comme suit.

|Application Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Boîte aux lettres|
|Word|WordApi|

La `isSetSupported` méthode et les ensembles de conditions requises pour ces applications sont disponibles dans le dernier fichier Office.js du CDN. Si vous n’utilisez pas Office.js du CDN, votre complément peut générer des exceptions, car il n’est pas `isSetSupported` défini. Pour plus d’informations, consultez [la rubrique spécifier la dernière bibliothèque de l’API JavaScript pour Office](#specify-the-latest-office-javascript-api-library).

L’exemple de code suivant montre comment un complément peut fournir des fonctionnalités différentes pour différentes applications Office qui peuvent prendre en charge différents ensembles de conditions requises ou membres d’API.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises

Certains membres API n’appartiennent pas à des ensembles de conditions requises. Ceci ne s’applique qu’aux membres de l’API qui font partie de l’espace de noms de l' [API JavaScript pour Office](../reference/javascript-api-for-office.md) (tout élément sous `Office.` sauf les [API de boîte aux lettres Outlook](/javascript/api/outlook)), mais pas les membres de l’API qui appartiennent à l' [API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) (quoi que ce soit `Word.` ), l' [API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md) (tout élément dans `Excel.` ) ou l' [API JavaScript OneNote](../reference/overview/onenote-add-ins-javascript-reference.md) `OneNote.` Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’exécution pour déterminer si la méthode est prise en charge par l’application Office, comme illustré dans l’exemple de code suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’application Office prend en charge `document.setSelectedDataAsync` .

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](add-in-manifests.md)
- [Ensembles de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
