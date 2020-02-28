---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: ''
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 4ee8dabd5a364a2c5566b2918c173da9b6d04a5a
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327788"
---
# <a name="specify-office-hosts-and-api-requirements"></a>Spécification des exigences en matière d’hôtes Office et d’API

Il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API pour fonctionner correctement. Par exemple, votre complément peut :

- Exécuter dans une application Office (Word ou Excel), ou plusieurs applications.

- utiliser des API JavaScript disponibles uniquement dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui fonctionne dans Excel 2016 ;

- s’exécuter uniquement dans les versions d’Office qui prennent en charge les membres d’API utilisés par votre complément.

Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

> [!NOTE]
> Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).

Le tableau suivant répertorie les concepts principaux décrits dans cet article.

|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application hôte Office ou hôte Office|Application Office utilisée pour exécuter votre complément. Par exemple, Word, Excel, etc.|
|Plateforme|Emplacement d’exécution de l’hôte Office, par exemple, dans un navigateur ou sur un iPad.|
|Ensemble de conditions requises|Groupe nommé de membres d’API associés. Les compléments utilisent des ensembles de conditions requises pour déterminer si l’hôte Office prend en charge les membres d’API utilisés par votre complément. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres individuels d’API. La prise en charge de l’ensemble des conditions requises varie selon l’hôte Office et la version de ce dernier. <br >Les ensembles de conditions requises sont spécifiés dans le fichier manifeste. Quand vous définissez des ensembles de conditions requises dans le fichier manifeste, vous définissez le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour exécuter votre complément. Les hôtes Office qui ne prennent pas en charge les ensembles de conditions requises spécifiés dans le manifeste ne peuvent pas exécuter votre complément, et votre complément ne sera pas affiché dans <span class="ui">Mes compléments</span>. Cela limite les emplacements où votre complément sera disponible. Dans le code utilisant les vérifications à l’exécution. Pour obtenir la liste complète des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).|
|Vérification à l’exécution|Test effectué à l’exécution pour déterminer si l’hôte Office qui exécute votre complément prend en charge les ensembles de conditions requises ou les méthodes utilisés par votre complément. Pour effectuer une vérification à l’exécution, vous **** utilisez une instruction if `isSetSupported` avec la méthode, les ensembles de conditions requises ou les noms de méthodes qui ne font pas partie d’un ensemble de conditions requises. Les vérifications à l’exécution permettent de veiller à ce que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’exécution ne précisent pas le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour l’exécution de votre complément. Au lieu de cela, vous utilisez l’instruction **If** pour déterminer si un membre de l’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affiche toujours dans **Mes compléments** quand vous effectuez des vérifications à l’exécution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez des vérifications à l’exécution dans votre complément, assurez-vous d’utiliser la dernière bibliothèque de l’API JavaScript pour Office (Office. js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du schéma de manifeste de complément le plus récent

Le manifeste de votre complément doit utiliser la version 1,1 du schéma de manifeste du complément. Définissez l' `OfficeApp` élément dans le manifeste de votre complément comme suit.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>Spécifier la dernière bibliothèque d’API JavaScript pour Office

Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de la bibliothèque de l’API JavaScript pour Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, ajoutez la balise suivante `script` à votre code html. L' `/1/` utilisation de dans l’URL du CDN garantit que vous faites référence à la version la plus récente d’Office. js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Options pour spécifier des hôtes Office ou les conditions requises d’API

Lors de la spécification des hôtes Office ou des conditions requises d’API, vous devez tenir compte de plusieurs facteurs. Le diagramme suivant montre comment choisir la technique à utiliser dans votre complément.

![Optez pour la meilleure solution pour votre complément lorsque vous spécifiez des hôtes Office ou des exigences d’API](../images/options-for-office-hosts.png)

- Si votre complément s’exécute dans un hôte Office, définissez l' `Hosts` élément dans le manifeste. Pour plus d’informations, consultez [la rubrique Set the hosts Element](#set-the-hosts-element).

- Pour définir l’ensemble de conditions minimales ou les membres de l’API qu’un hôte Office doit prendre en charge pour exécuter votre `Requirements` complément, définissez l’élément dans le manifeste. Pour plus d’informations, voir [définir l’élément Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).

- Si vous souhaitez proposer des fonctionnalités supplémentaires lorsque des ensembles de conditions requises spécifiques ou des membres d’API sont disponibles dans l’hôte Office, effectuez une vérification à l’exécution dans le code JavaScript de votre complément. Par exemple, si votre complément est exécuté dans Excel 2016, utilisez les membres d’API de l’API JavaScript Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, consultez la section [Utilisation des vérifications à l’exécution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).

## <a name="set-the-hosts-element"></a>Définition de l’élément Hosts

Pour que votre complément s’exécute dans une application hôte Office, utilisez les `Hosts` éléments et `Host` dans le manifeste. Si vous ne spécifiez `Hosts` pas l’élément, votre complément s’exécutera sur tous les hôtes.

Par exemple, la Déclaration `Hosts` et `Host` suivante spécifie que le complément fonctionnera avec n’importe quelle version d’Excel, y compris Excel sur le Web, Windows et iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

L' `Hosts` élément peut contenir un ou plusieurs `Host` éléments. L' `Host` élément spécifie l’hôte Office dont votre complément a besoin. L' `Name` attribut est obligatoire et peut prendre la valeur de l’une des valeurs suivantes.

| Name          | Applications hôtes Office                                                                  |
|:--------------|:------------------------------------------------------------------------------------------|
| Base de données      | applications web Access                                                                           |
| Document      | Word pour Windows, Word pour Mac, Word pour iPad, Word sur le web                               |
| Boîte aux lettres       | Outlook pour Windows, Outlook pour Mac, Outlook sur le web, Outlook pour Android, Outlook pour iOS|
| Présentation  | PowerPoint pour Windows, PowerPoint pour Mac, PowerPoint pour iPad, PowerPoint sur le web       |
| Project       | Project sur Windows                                                                        |
| Classeur      | Excel pour Windows, Excel pour Mac, Excel pour iPad, Excel sur le web                           |

> [!NOTE]
> L' `Name` attribut spécifie l’application hôte Office qui peut exécuter votre complément. Les hôtes Office sont pris en charge sur différentes plateformes et sont exécutés sur les ordinateurs de bureau, les navigateurs web, les tablettes et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez `Mailbox`, Outlook sur le web et Outlook sur Windows peuvent être utilisés pour exécuter votre complément.

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.


## <a name="set-the-requirements-element-in-the-manifest"></a>Définition de l’élément Requirements dans le manifeste

L' `Requirements` élément spécifie les ensembles de conditions requises minimum ou les membres d’API qui doivent être pris en charge par l’hôte Office pour exécuter votre complément. L' `Requirements` élément peut spécifier des ensembles de conditions requises et des méthodes individuelles utilisées dans votre complément. Dans la version 1,1 du schéma de manifeste de complément, l' `Requirements` élément est facultatif pour tous les compléments, à l’exception des compléments Outlook.

> [!WARNING]
> Utilisez uniquement l' `Requirements` élément pour spécifier des ensembles de conditions requises critiques ou des membres d’API que votre complément doit utiliser. Si l’hôte ou la plateforme Office ne prend pas en charge les ensembles de conditions requises `Requirements` ou les membres d’API spécifiés dans l’élément, le complément ne s’exécutera pas sur cet hôte ou cette plateforme, et ne s’affichera pas dans **mes compléments**. Au lieu de cela, nous vous recommandons de faire en sorte que votre complément soit disponible sur toutes les plateformes d’un hôte Office, comme Excel sur le Web, Windows et iPad. Pour que votre complément soit disponible sur _tous les_ hôtes et plateformes Office, utilisez des vérifications à l' `Requirements` exécution à la place de l’élément.

Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge les éléments suivants :

-  `TableBindings`ensemble de conditions requises, dont la version minimale est « 1,1 ».

-  `OOXML`ensemble de conditions requises, dont la version minimale est « 1,1 ».

-  `Document.getSelectedDataAsync`procédé.

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

- L' `Requirements` élément contient les `Sets` éléments `Methods` enfants et.

- L' `Sets` élément peut contenir un ou plusieurs `Set` éléments. `DefaultMinVersion` spécifie la `MinVersion` valeur par défaut de `Set` tous les éléments enfants.

- L' `Set` élément spécifie les ensembles de conditions requises que l’hôte Office doit prendre en charge pour exécuter le complément. L' `Name` attribut spécifie le nom de l’ensemble de conditions requises. L `MinVersion` 'spécifie la version minimale de l’ensemble de conditions requises. `MinVersion`remplace la valeur de `DefaultMinVersion` pour plus d’informations sur les ensembles de conditions requises et les versions d’ensemble de conditions requises auxquelles appartiennent les membres de l’API, consultez la rubrique ensembles de conditions requises pour les [Compléments Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).

- L' `Methods` élément peut contenir un ou plusieurs `Method` éléments. Vous ne pouvez pas `Methods` utiliser l’élément avec des compléments Outlook.

- L' `Method` élément spécifie une méthode individuelle qui doit être prise en charge dans l’hôte Office où s’exécute votre complément. L' `Name` attribut est obligatoire et spécifie le nom de la méthode qualifiée avec son objet parent.

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des vérifications à l’exécution dans votre code JavaScript

Vous pouvez fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’hôte Office. Par exemple, vous pouvez utiliser les interfaces API JavaScript de Word dans votre complément existant si ce dernier est exécuté dans Word 2016. Pour ce faire, utilisez la méthode [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) avec le nom de l’ensemble de conditions requises. `isSetSupported`détermine, lors de l’exécution, si l’hôte Office qui exécute le complément prend en charge l’ensemble de conditions requises. Si l’ensemble de conditions requises est `isSetSupported` pris en charge, renvoie la **valeur true** et exécute le code supplémentaire qui utilise les membres de l’API à partir de cet ensemble de conditions requises. Si l’hôte Office ne prend pas en charge l' `isSetSupported` ensemble de conditions requises, renvoie la **valeur false** et le code supplémentaire ne s’exécute pas. Le code suivant illustre la syntaxe à utiliser avec `isSetSupported`.

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (obligatoire) est une chaîne qui représente le nom de l’ensemble de la configuration requise (p. ex., « **ExcelApi** », « **Mailbox** », etc.). Pour plus d’informations concernant les ensembles de conditions requises disponibles, voir [Ensembles de conditions requises pour complément Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).
- _MinimumVersion_ (facultatif) est une chaîne qui spécifie la version minimale requise que l’hôte doit prendre en charge afin de permettre l’exécution de l’instruction `if` dans le code (p. ex. « **1.9** »).

> [!WARNING]
> Lors de l' `isSetSupported` appel de la méthode, la `MinimumVersion` valeur du paramètre (s’il est spécifié) doit être une chaîne. En effet, l’analyseur syntaxique JavaScript ne peut pas différencier les valeurs numériques, telles que 1.1 et 1.10, mais le peut pour les valeurs chaîne, telles que « 1.1 » et « 1.10 ».
> La surcharge `number` est déconseillée.

Utilisez `isSetSupported` -le `RequirementSetName` avec l’hôte Office comme suit.

|Hôte Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Boîte aux lettres|
|Word|WordApi|

La `isSetSupported` méthode et les ensembles de conditions requises pour ces hôtes sont disponibles dans le fichier Office. js le plus récent sur le CDN. Si vous n’utilisez pas Office. js à partir du CDN, votre complément peut générer des exceptions `isSetSupported` , car il n’est pas défini. Pour plus d’informations, consultez [la rubrique spécifier la dernière bibliothèque de l’API JavaScript pour Office](#specify-the-latest-office-javascript-api-library).

L’exemple de code suivant montre comment un complément peut fournir des fonctionnalités différentes pour divers hôtes Office qui peuvent prendre en charge plusieurs ensembles de conditions requises ou membres d’API.

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
    // Run additional code when the Office host is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises

Certains membres API n’appartiennent pas à des ensembles de conditions requises. Ceci ne s’applique qu’aux membres de l’API qui font partie de l’espace de noms `Office.` de l' [API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office) (tout élément sous sauf les API de [boîte aux lettres Outlook](/javascript/api/outlook)), `Word.`mais pas les membres de l’API qui appartiennent à l' [API JavaScript pour Word](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview) (quoi que ce soit), l' [API JavaScript pour Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) (tout élément `Excel.`dans) ou l' [API](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference) `OneNote.`JavaScript OneNote Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’exécution pour déterminer si la méthode est prise en charge par l’hôte Office, comme indiqué dans l’exemple suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’hôte `document.setSelectedDataAsync`prend en charge.

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](add-in-manifests.md)
- [Ensembles de conditions requises pour les compléments Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
