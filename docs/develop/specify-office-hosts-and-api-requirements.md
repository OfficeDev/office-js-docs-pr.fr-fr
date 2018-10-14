---
title: Spécification des hôtes Office et des conditions requises pour les API
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: ff6c0e4b4b2f8a517a62932722c34142ffdab609
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505978"
---
# <a name="specify-office-hosts-and-api-requirements"></a>Spécification des hôtes Office et des conditions requises pour les API

Il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version d’API pour fonctionner correctement. Par exemple, votre complément peut :

- s'exécuter dans une seule application Office (Word ou Excel), ou dans plusieurs applications ;
    
- utiliser des API JavaScript qui ne sont disponibles que dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui s'exécute dans Excel 2016 ; 
    
- s’exécuter seulement dans les versions d’Office qui prennent en charge les membres d’API que votre complément utilise.
    
Cet article vous aidera à comprendre quelles options vous devez choisir pour vous assurer que votre complément fonctionne comme attendu et qu'il atteigne l’audience la plus large possible.

> [!NOTE]
> Pour une vue d'ensemble des emplacements où les compléments Office sont actuellement pris en charge, voir la page [Disponibilité des hôtes et des plateformes pour un complément Office](../overview/office-add-in-availability.md). 

La table suivante liste les concepts de base décrits tout au long de cet article.

|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application hôte Office, hôte Office, ou hôte|L'application Office utilisée pour exécuter votre complément. Par exemple, Word, Word Online, Excel, et ainsi de suite.|
|Plateforme|L'emplacement où l’hôte Office s'exécute, comme Office Online ou Office pour iPad.|
|Ensemble de conditions requises|Un groupe nommé de membres d’API associés. Les compléments utilisent les ensembles de conditions requises pour déterminer si l’hôte Office prend en charge les membres d’API utilisés par votre complément. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres d’API individuels. La prise en charge d’un ensemble des conditions requises varie selon l’hôte Office et la version de l'hôte Office. <br >Les ensembles de conditions requises sont indiqués dans le fichier manifeste. Lorsque vous indiquez des ensembles de conditions requises dans le manifeste, vous définissez le niveau minimal de prise en charge d'API que l’hôte Office doit fournir pour exécuter votre complément. Les hôtes Office qui ne prennent pas en charge les ensembles de conditions requises indiqués dans le manifeste ne peuvent pas exécuter votre complément, et votre complément ne s'affichera pas dans <span class="ui">Mes compléments</span>. Ceci limite les emplacements où votre complément sera disponible. Dans le code utilisant des vérifications à l’exécution. Pour la liste complète des ensembles de conditions requises, voir [Ensembles de conditions requises pour les compléments Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).|
|Vérification à l’exécution|Un test qui est effectué à l’exécution pour déterminer si l’hôte Office qui exécute votre complément prend en charge les ensembles de conditions requises ou les méthodes utilisées par votre complément. Pour effectuer une vérification à l’exécution, vous pouvez utiliser une instruction **if** avec la méthode **isSetSupported**, les ensembles de conditions requises, ou les noms de méthode qui ne font pas partie d’un ensemble de conditions requises. Utilisez les vérifications à l’exécution pour vous assurer que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’exécution n'indiquent pas le niveau minimal de prise en charge d’API que l’hôte Office doit fournir pour que votre complément s'exécute. A la place, vous devez utiliser l’instruction **if** pour déterminer si un membre d’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affichera toujours dans **Mes compléments** quand vous utilisez des vérifications à l’exécution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez les vérifications à l’exécution dans votre complément, assurez-vous que vous utilisez la toute dernière bibliothèque API JavaScript pour Office (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du tout dernier schéma de manifeste de complément

Le manifeste de votre complément doit utiliser la version 1.1 du schéma de manifeste de complément. Définissez l’élément **AppOffice** dans le manifeste de votre complément comme suit.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a>Indication de la toute dernière bibliothèque API JavaScript pour Office

Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de la bibliothèque API JavaScript pour Office à partir du réseau de distribution de contenu (CDN). Pour ce faire, ajoutez la balise suivante `script` à votre code HTML. Le fait d’utiliser `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Options pour indiquer des hôtes Office ou des conditions requises pour les API

Quand vous indiquez des hôtes Office ou des conditions requises pour les API, il y a plusieurs facteurs à considérer. Le diagramme suivant montre comment décider quelle technique utiliser dans votre complément.

![Choisissez la meilleure option pour votre complément lorsque vous indiquez des hôtes Office ou des conditions requises pour les API](../images/options-for-office-hosts.png)

- Si votre complément s'exécute dans un seul hôte Office, définissez l'élément **Hosts** dans le manifeste. Pour plus d'informations, voir [Définir l'élément Hosts](#set-the-hosts-element).
    
- Pour définir l’ensemble minimal de conditions requises ou les membres d’API qu’un hôte Office doit prendre en charge pour exécuter votre complément, définissez l’élément **Requirements** dans le manifeste. Pour plus d’informations, voir la section [Définition de l’élément Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).
    
- Si vous souhaitez proposer des fonctionnalités supplémentaires lorsque des ensembles de conditions requises ou des membres d’API particuliers sont disponibles dans l’hôte Office, effectuez une vérification à l’exécution dans le code JavaScript de votre complément. Par exemple, si votre complément s'exécute dans Excel 2016, utilisez les membres d’API de la nouvelle API JavaScript pour Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, voir [Utilisation des vérifications à l’exécution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).
    
## <a name="set-the-hosts-element"></a>Définition de l’élément Hosts

Pour faire que votre complément s'exécute dans une seule application hôte Office, utilisez les éléments **Hosts** et **Host** dans le manifeste. Si vous ne définissez pas l’élément **Hosts**, votre complément s'exécutera dans tous les hôtes.

Par exemple, les déclarations **Hosts** et **Host** suivantes indiquent que le complément fonctionnera avec n’importe quelle version d’Excel, ce qui comprend Excel pour Windows, Excel Online, et Excel pour iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

L’élément **Hosts** peut contenir un ou plusieurs éléments  **Host**. L’élément  **Host** indique l’hôte Office que votre complément requiert. L’attribut **Name** est requis et peut être défini à l’une des valeurs suivantes.

| Name          | Applications hôtes Office                      |
|:--------------|:----------------------------------------------|
| Base de données      | Applications web Access                               |
| Document      | Word pour Windows, Mac, iPad et Online        |
| Boîte aux lettres       | Outlook pour Windows, Mac, Web et Outlook.com | 
| Présentation  | PowerPoint pour Windows, Mac, iPad et Online  |
| Projet       | Projet                                       |
| Classeur      | Excel Windows, Mac, iPad et Online           |

> [!NOTE]
> L’attribut `Name` indique l'application hôte Office qui peut exécuter votre complément. Les hôtes Office sont pris en charge sur différentes plateformes et s'exécutent sur les ordinateurs de bureau, les navigateurs web, les tablettes, et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous indiquez `Mailbox`, à la fois Outlook et Outlook Web App peuvent être utilisés pour exécuter votre complément. 


## <a name="set-the-requirements-element-in-the-manifest"></a>Définition de l’élément Requirements dans le manifeste

L’élément **Requirements** indique les ensembles de conditions requises minimaux ou les membres d’API qui doivent être pris en charge par l’hôte Office pour exécuter votre complément. L’élément **Requirements** peut indiquer à la fois des ensembles de conditions requises et des méthodes individuelles utilisés dans votre complément. Dans la version 1.1 du schéma de manifeste de complément, l’élément **Requirements** est facultatif pour tous les compléments, sauf pour les compléments Outlook.

> [!WARNING]
> Utilisez seulement l’élément **Requirements** pour indiquer des ensembles de conditions requises ou des membres d’API cruciaux que votre complément doit utiliser. Si l’hôte ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API indiqués dans l’élément **Requirements**, le complément ne s’exécutera pas sur cet hôte ou cette plateforme, et il ne s’affichera pas dans **Mes compléments**. A la place, nous vous recommandons de rendre votre complément disponible sur toutes les plateformes d’un hôte Office, tel qu'Excel pour Windows, Excel Online, et Excel pour iPad. Pour rendre votre complément disponible sur _tous_ les hôtes et plateformes Office, utilisez des vérifications à l’exécution à la place de l’élément **Requirements**.

L'exemple de code suivant montre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge ce qui suit :

-  Un ensemble de conditions requises **TableBindings**, qui a une version minimale de 1.1.
    
-  Un ensemble de conditions requises **OOXML**, qui a une version minimale de 1.1.
    
-  La méthode **Document.getSelectedDataAsync**.

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

- L’élément **Requirements** contient les éléments enfants **Sets** et **Methods**.
    
- L’élément **Sets** peut contenir un ou plusieurs éléments **Set**.  **DefaultMinVersion** indique la valeur **MinVersion** par défaut de tous les éléments **Set** enfants.
    
- L’élément **Set** indique les ensembles de conditions requises que l’hôte Office doit prendre en charge pour exécuter le complément. L’attribut **Name** indique le nom de l’ensemble de conditions requises. **MinVersion** indique la version minimale de l’ensemble de conditions requises. **MinVersion** remplace la valeur de **DefaultMinVersion**. Pour plus d’informations sur les ensembles de conditions requises et les versions des ensembles de conditions requises auxquels vos membres d'API appartiennent, voir [Les ensembles de conditions requises pour un complément Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).
    
- L’élément **Methods** peut contenir un ou plusieurs éléments **Method**. Vous ne pouvez pas utiliser l’élément **Methods** avec des compléments Outlook.
    
- L’élément **Method** indique une méthode individuelle qui doit être prise en charge dans l’hôte Office dans lequel votre complément s'exécute. L’attribut **Name** est requis et indique le nom de la méthode qualifiée avec son objet parent.
    

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des vérifications à l’exécution dans votre code JavaScript


Vous pouvez vouloir fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’hôte Office. Par exemple, vous pouvez vouloir utiliser les nouvelles API JavaScript de Word dans votre complément existant si votre complément s'exécute dans Word 2016. Pour ce faire, utilisez la méthode **isSetSupported** avec le nom de l’ensemble de conditions requises. **isSetSupported** détermine, à l’exécution, si l’hôte Office exécutant le complément prend en charge l’ensemble de conditions requises. Si l’ensemble de conditions requises est pris en charge, **isSetSupported** retourne **true** et exécute le code supplémentaire qui utilise les membres d’API provenant de cet ensemble de conditions requises. Si l’hôte Office ne prend pas en charge l’ensemble de conditions requises, **isSetSupported** retourne **false** et le code supplémentaire ne s'exécutera pas. Le code suivant indique la syntaxe à utiliser avec **isSetSupported**.


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  _RequirementSetName_ (requis) est une chaîne qui représente le nom de l’ensemble de conditions requises. Pour plus d’informations sur les ensembles de conditions requises disponibles, voir [Les ensembles de conditions requises pour un complément Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).
    
-  _VersionNumber_ (facultatif) est la version de l’ensemble de conditions requises.
    
Dans Excel 2016 ou Word 2016, utilisez **isSetSupported** avec les ensembles de conditions requises **ExcelAPI** ou **WordAPI**. La méthode **isSetSupported**, et les ensembles de conditions requises **ExcelAPI** et **WordAPI**, sont disponibles dans le tout dernier fichier Office.js disponible depuis le CDN. Si vous n’utilisez pas Office.js à partir du CDN, votre complément pourra générer des exceptions du fait que la méthode **isSetSupported** ne sera pas définie. Pour plus d’informations, voir [Indication de la toute dernière bibliothèque API JavaScript pour Office](#specify-the-latest-javascript-api-for-office-library). 


> [!NOTE]
> **isSetSupported** ne fonctionne pas dans Outlook ou Outlook Web App. Pour utiliser une vérification à l’exécution dans Outlook ou Outlook Web App, utilisez la technique décrite dans la section [Vérifications à l’exécution en utilisant des méthodes ne faisant pas partie d’un ensemble de conditions requises](#runtime-checks-using-methods-not-in-a-requirement-set).

L’exemple de code suivant montre comment un complément peut fournir des fonctionnalités différentes pour divers hôtes Office qui peuvent prendre en charge différents ensembles de conditions requises ou membres d’API.




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Vérifications à l’exécution utilisant des méthodes ne faisant pas partie d’un ensemble de conditions requises


Certains membres d’API n’appartiennent pas à des ensembles de conditions requises. Ceci ne s’applique qu'aux membres d’API qui font partie de l’espace de noms [API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) (tout qui se trouve sous Office.), pas aux membres de l’API qui appartiennent à l’API JavaScript de Word (tout se qui se trouve dans Word) ou aux espaces de noms [Référence de l’API JavaScript des compléments Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) (tout se qui se trouve dans Excel.). Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’exécution pour déterminer si la méthode est prise en charge par l’hôte Office, comme montré dans l’exemple de code suivant. Pour une liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Les ensembles de conditions requises pour un complément Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).


> [!NOTE]
> Nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’hôte prend en charge **document.setSelectedDataAsync**.




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](add-in-manifests.md)
- [Ensembles d'exigences pour les compléments Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)