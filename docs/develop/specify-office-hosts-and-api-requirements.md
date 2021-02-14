---
title: Spécification des exigences en matière d’hôtes Office et d’API
description: Découvrez comment spécifier les applications Office et les conditions requises d’API pour que votre application fonctionne comme prévu.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 948e86e99150ebf2d0bc7deaa5512627679b025f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237839"
---
# <a name="specify-office-applications-and-api-requirements"></a>Spécifier les applications Office et les exigences de l’API

Votre application Office peut dépendre d’une application Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API pour fonctionner comme prévu. Par exemple, votre complément peut :

- Exécuter dans une application Office (Word ou Excel), ou plusieurs applications.

- utiliser des API JavaScript disponibles uniquement dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui fonctionne dans Excel 2016 ;

- s’exécuter uniquement dans les versions d’Office qui prennent en charge les membres d’API utilisés par votre complément.

Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

> [!NOTE]
> Pour obtenir une vue générale de l’endroit où les applications Office sont actuellement pris en charge, consultez la page disponibilité de la plateforme et de l’application cliente Office pour les [applications Office.](../overview/office-add-in-availability.md)

Le tableau suivant répertorie les concepts principaux décrits dans cet article.

|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application cliente Office|Application Office utilisée pour exécuter votre complément. Par exemple, Word, Excel, etc.|
|Plateforme|L’endroit où l’application Office s’exécute, par exemple dans un navigateur ou sur un iPad.|
|Ensemble de conditions requises|Groupe nommé de membres d’API associés. Les add-ins utilisent des ensembles de conditions requises pour déterminer si l’application Office prend en charge les membres d’API utilisés par votre application. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres individuels d’API. La prise en charge de l’ensemble de conditions requises varie selon l’application Office et la version de l’application Office. <br >Les ensembles de conditions requises sont spécifiés dans le fichier manifeste. Lorsque vous spécifiez des ensembles de conditions requises dans le manifeste, vous définissez le niveau minimal de prise en charge de l’API que l’application Office doit fournir pour exécuter votre application. Les applications Office qui ne peuvent pas prendre en charge les ensembles de conditions <span class="ui">requises spécifiés</span>dans le manifeste ne peuvent pas exécuter votre application et votre application ne s’affichera pas dans Mes applications. Cela limite l’endroit où votre add-in est disponible. Dans le code utilisant les vérifications à l’exécution. Pour obtenir la liste complète des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md).|
|Vérification à l’exécution|Test effectué au moment de l’exécution pour déterminer si l’application Office exécutant votre application prend en charge les ensembles de conditions requises ou les méthodes utilisées par votre application. Pour effectuer une vérification à l’runtime, vous utilisez une instruction **if** avec la méthode, les ensembles de conditions requises ou les noms de méthodes qui ne font pas partie `isSetSupported` d’un ensemble de conditions requises. Les vérifications à l’exécution permettent de veiller à ce que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’runtime ne spécifient pas le niveau minimal de prise en charge de l’API que l’application Office doit fournir pour que votre application s’exécute. À la place, vous utilisez **l’instruction if** pour déterminer si un membre d’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affiche toujours dans **Mes compléments** quand vous effectuez des vérifications à l’exécution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez les vérifications à l’runtime dans votre application, assurez-vous d’utiliser la dernière bibliothèque d’API JavaScript (office.js) Office.

### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du schéma de manifeste de complément le plus récent

Le manifeste de votre du complément doit utiliser la version 1.1 du schéma de manifeste de complément. Définissez `OfficeApp` l’élément dans le manifeste de votre add-in comme suit.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>Spécifier la dernière bibliothèque d’API JavaScript pour Office

Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de la bibliothèque d’API JavaScript Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante à votre code HTML. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>Options permettant de spécifier les applications Office ou les conditions requises des API

Lorsque vous spécifiez des applications Office ou des exigences d’API, plusieurs facteurs sont à prendre en compte. Le diagramme suivant montre comment choisir la technique à utiliser dans votre complément.

![Choisissez la meilleure option pour votre application lors de la spécification d’applications Office ou d’exigences d’API](../images/options-for-office-hosts.png)

- Si votre add-in s’exécute dans une application Office, définissez `Hosts` l’élément dans le manifeste. Pour plus d’informations, consultez [Définition de l’élément Hosts](#set-the-hosts-element).

- Pour définir l’ensemble minimal de conditions requises ou les membres d’API qu’une application Office doit prendre en charge pour exécuter votre application, définissez l’élément `Requirements` dans le manifeste. Pour plus d’informations, consultez la section [ Définition de l’élément Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).

- Si vous souhaitez fournir des fonctionnalités supplémentaires si des ensembles de conditions requises ou des membres d’API spécifiques sont disponibles dans l’application Office, effectuez une vérification à l’runtime dans le code JavaScript de votre complément. Par exemple, si votre complément est exécuté dans Excel 2016, utilisez les membres d’API de l’API JavaScript Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, consultez la section [Utilisation des vérifications à l’exécution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).

## <a name="set-the-hosts-element"></a>Définition de l’élément Hosts

Pour exécuter votre application dans une application cliente Office, utilisez les éléments et `Hosts` `Host` les éléments du manifeste. Si vous ne spécifiez pas l’élément, votre application s’exécutera dans toutes les applications Office pris en charge par `Hosts` les applications Office.

Par exemple, l’exemple suivant et la déclaration spécifient que le add-in fonctionne avec n’importe quelle version d’Excel, qui inclut Excel sur le web, Windows et `Hosts` `Host` iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

`Hosts`L’élément peut contenir un ou plusieurs `Host` éléments. `Host`L’élément spécifie l’application Office dont votre application a besoin. `Name`L’attribut est obligatoire et peut être définie sur l’une des valeurs suivantes.

| Nom          | Applications clientes Office                      |
|:--------------|:----------------------------------------------|
| Base de données      | applications web Access                               |
| Document      | Word sur le web, Windows, Mac, iPad           |
| Boîte aux lettres       | Outlook sur le web, Windows, Mac, Android, iOS|
| Présentation  | PowerPoint sur le web, Windows, Mac, iPad     |
| Project       | Project sur Windows                            |
| Classeur      | Excel sur le web, Windows, Mac, iPad          |

> [!NOTE]
> `Name`L’attribut spécifie l’application cliente Office qui peut exécuter votre application. Les applications Office sont pris en charge sur différentes plateformes et s’exécutent sur des ordinateurs de bureau, des navigateurs web, des tablettes et des appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez , Outlook sur le web et sur Windows peuvent être utilisés pour `Mailbox` exécuter votre application.

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

## <a name="set-the-requirements-element-in-the-manifest"></a>Définition de l’élément Requirements dans le manifeste

L’élément spécifie les ensembles de conditions requises minimum ou les membres d’API qui doivent être pris en charge par `Requirements` l’application Office pour exécuter votre application. L’élément peut spécifier des ensembles de conditions requises et des méthodes `Requirements` individuelles utilisées dans votre add-in. Dans la version 1.1 du schéma de manifeste du add-in, l’élément est facultatif pour tous les modules, à l’exception des `Requirements` autres.

> [!WARNING]
> Utilisez uniquement l’élément pour spécifier des ensembles de conditions requises critiques ou des membres `Requirements` d’API que votre application doit utiliser. Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres `Requirements` d’API **spécifiés** dans l’élément, le module ne s’exécutera pas dans cette application ou plateforme et ne s’affichera pas dans Mes applications. Au lieu de cela, nous vous recommandons de rendre votre application disponible sur toutes les plateformes d’une application Office, telles qu’Excel sur le web, Windows et iPad. Pour rendre votre application disponible sur toutes les applications  _et_ plateformes Office, utilisez les vérifications à l’runtime à la place de `Requirements` l’élément.

L’exemple de code suivant montre un add-in qui se charge dans toutes les applications clientes Office qui supportent les données suivantes :

-  `TableBindings` ensemble de conditions requises, dont la version minimale est « 1.1 ».

-  `OOXML` ensemble de conditions requises, dont la version minimale est « 1.1 ».

-  `Document.getSelectedDataAsync` .

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

- `Requirements`L’élément contient les éléments enfants et les `Sets` `Methods` éléments.

- `Sets`L’élément peut contenir un ou plusieurs `Set` éléments. `DefaultMinVersion` spécifie la valeur par `MinVersion` défaut de tous les éléments `Set` enfants.

- `Set`L’élément spécifie les ensembles de conditions requises que l’application Office doit prendre en charge pour exécuter le module. `Name`L’attribut spécifie le nom de l’ensemble de conditions requises. Spécifie `MinVersion` la version minimale de l’ensemble de conditions requises. `MinVersion`remplace la valeur de Pour plus d’informations sur les ensembles de conditions requises et les versions d’ensembles de conditions requises dont appartiennent les membres de votre API, consultez les ensembles de conditions requises pour les `DefaultMinVersion` [add-ins Office.](../reference/requirement-sets/office-add-in-requirement-sets.md)

- `Methods`L’élément peut contenir un ou plusieurs `Method` éléments. Vous ne pouvez pas utiliser `Methods` l’élément avec des add-ins Outlook.

- L’élément spécifie une méthode individuelle qui doit être prise en charge dans `Method` l’application Office dans laquelle votre application s’exécute. `Name` L’attribut est obligatoire et spécifie le nom de la méthode qualifiée avec son objet parent.

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des vérifications à l’exécution dans votre code JavaScript

Vous souhaitez peut-être fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’application Office. Par exemple, vous pouvez utiliser les interfaces API JavaScript de Word dans votre complément existant si ce dernier est exécuté dans Word 2016. Pour ce faire, utilisez la méthode [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) avec le nom de l’ensemble de conditions requises. `isSetSupported` détermine, au moment de l’exécution, si l’application Office qui exécute le add-in prend en charge l’ensemble de conditions requises. Si l’ensemble de conditions requises est pris en charge, renvoie true et exécute le code supplémentaire qui utilise les membres `isSetSupported` d’API  de cet ensemble de conditions requises. Si l’application Office ne prend pas en charge l’ensemble de conditions requises, renvoie false et le code supplémentaire `isSetSupported` ne s’exécute  pas. Le code suivant indique la syntaxe à utiliser avec `isSetSupported`

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (obligatoire) est une chaîne qui représente le nom de l’ensemble de la configuration requise (p. ex., « **ExcelApi** », « **Mailbox** », etc.). Pour plus d’informations concernant les ensembles de conditions requises disponibles, voir [Ensembles de conditions requises pour complément Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- _MinimumVersion_ (facultatif) est une chaîne qui spécifie la version minimale de l’ensemble de conditions requises que l’application Office doit prendre en charge pour que le code de l’instruction s’exécute `if` (par exemple, «**1,9**»).

> [!WARNING]
> Lors de l’appel de la méthode, la valeur du paramètre `isSetSupported` `MinimumVersion` (si spécifié) doit être une chaîne. En effet, l’analyseur syntaxique JavaScript ne peut pas différencier les valeurs numériques, telles que 1.1 et 1.10, mais le peut pour les valeurs chaîne, telles que « 1.1 » et « 1.10 ».
> La surcharge `number` est déconseillée.

À `isSetSupported` utiliser avec `RequirementSetName` l’application Office associée comme suit.

|Application Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Boîte aux lettres|
|Word|WordApi|

La méthode et les ensembles de conditions requises pour ces applications sont disponibles dans le dernier `isSetSupported` fichier Office.js sur le CDN. Si vous n’utilisez pas Office.js à partir du CDN, votre add-in peut générer des exceptions, car elle ne sera `isSetSupported` pas définie. Pour plus d’informations, voir [Spécifier la dernière bibliothèque d’API JavaScript pour Office.](#specify-the-latest-office-javascript-api-library)

L’exemple de code suivant montre comment un application peut fournir des fonctionnalités différentes pour différentes applications Office qui peuvent prendre en charge différents ensembles de conditions requises ou membres d’API.

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

Certains membres API n’appartiennent pas à des ensembles de conditions requises. Cela s’applique uniquement aux membres d’API qui font partie de l’espace de noms de [l’API JavaScript pour Office](../reference/javascript-api-for-office.md) (à l’exception des API de boîte aux lettres Outlook), mais pas aux membres d’API qui appartiennent à l’API JavaScript pour Word (quelque chose dans), à l’API JavaScript pour Excel (tout dans) ou à l’API `Office.` [](/javascript/api/outlook) [](../reference/overview/word-add-ins-reference-overview.md) `Word.` [](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` [JavaScript OneNote](../reference/overview/onenote-add-ins-javascript-reference.md) (tout élément dans les espaces de `OneNote.` noms). Lorsque votre application dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’runtime pour déterminer si la méthode est prise en charge par l’application Office, comme illustré dans l’exemple de code suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

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
