---
title: Sp?cification des exigences en mati?re d?h?tes Office et d?API
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd517dee1faf8d3f3009a0b9ce7127f5760e730d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="specify-office-hosts-and-api-requirements"></a>Sp?cification des exigences en mati?re d?h?tes Office et d?API

Il se peut que votre compl?ment Office d?pende d?un h?te Office sp?cifique, d?un ensemble de conditions requises, d?un membre d?API ou d?une version de l?API pour fonctionner correctement. Par exemple, votre compl?ment peut :

- ex?cuter une ou plusieurs application Office (Word ou Excel) ;
    
- utiliser des API JavaScript disponibles uniquement dans certaines versions d?Office. Par exemple, vous pouvez utiliser les API JavaScript d?Excel dans un compl?ment qui fonctionne dans Excel 2016 ; 
    
- s?ex?cuter uniquement dans les versions d?Office qui prennent en charge les membres d?API utilis?s par votre compl?ment.
    
Cet article vous aidera ? comprendre les options que vous devez choisir afin de vous assurer que votre compl?ment fonctionne comme pr?vu et atteint l?audience la plus large possible.

> [!NOTE]
> Pour savoir de mani?re d?taill?e quelle version d?Office prend en charge les compl?ments Office, consultez la page relative ? la [disponibilit? des compl?ments Office sur les plateformes et les h?tes](../overview/office-add-in-availability.md). 

Le tableau suivant r?pertorie les concepts principaux d?crits dans cet article.

|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application h?te Office ou h?te Office|Application Office utilis?e pour ex?cuter votre compl?ment. Par exemple, Word, Word Online ou Excel.|
|Plateforme|Application sur laquelle l?h?te Office est ex?cut?, comme Office Online ou Office pour iPad.|
|Ensemble de conditions requises|Groupe nomm? de membres d?API associ?s. Les compl?ments utilisent des ensembles de conditions requises pour d?terminer si l?h?te Office prend en charge les membres d?API utilis?s par votre compl?ment. Il est plus facile de tester la prise en charge d?un ensemble de conditions requises, plut?t que la prise en charge de membres individuels d?API. La prise en charge de l?ensemble des conditions requises varie selon l?h?te Office et la version de ce dernier. <br >Les ensembles de conditions requises sont sp?cifi?s dans le fichier manifeste. Quand vous d?finissez des ensembles de conditions requises dans le fichier manifeste, vous d?finissez le niveau minimal de prise en charge de l?API que l?h?te Office doit fournir pour ex?cuter votre compl?ment. Les h?tes Office qui ne prennent pas en charge les ensembles de conditions requises sp?cifi?s dans le manifeste ne peuvent pas ex?cuter votre compl?ment, et votre compl?ment ne sera pas affich? dans <span class="ui">Mes compl?ments</span>. Cela limite les emplacements o? votre compl?ment sera disponible. Dans le code utilisant les v?rifications ? l?ex?cution. Pour obtenir la liste compl?te des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).|
|V?rification ? l?ex?cution|Test effectu? ? l?ex?cution pour d?terminer si l?h?te Office qui ex?cute votre compl?ment prend en charge les ensembles de conditions requises ou les m?thodes utilis?s par votre compl?ment. Pour effectuer une v?rification ? l?ex?cution, vous pouvez utiliser une instruction **if** avec la m?thode **isSetSupported**, les ensembles de conditions requises ou les noms de m?thode qui ne font pas partie d?un ensemble de conditions requises. Les v?rifications ? l?ex?cution permettent de veiller ? ce que votre compl?ment atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les v?rifications ? l?ex?cution ne pr?cisent pas le niveau minimal de prise en charge de l?API que l?h?te Office doit fournir pour l?ex?cution de votre compl?ment. Au lieu de cela, vous devez utiliser l?instruction **if** afin de d?terminer si un membre d?API est pris en charge. Si c?est le cas, vous pouvez fournir des fonctionnalit?s suppl?mentaires dans votre compl?ment. Votre compl?ment s?affiche toujours dans **Mes compl?ments** quand vous effectuez des v?rifications ? l?ex?cution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre compl?ment doit utiliser la version la plus r?cente du sch?ma de manifeste de compl?ment. Si vous utilisez les v?rifications ? l?ex?cution dans votre compl?ment, assurez-vous que vous utilisez la derni?re API JavaScript pour la biblioth?que Office (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du sch?ma de manifeste de compl?ment le plus r?cent

Le manifeste de votre du compl?ment doit utiliser la version 1.1 du sch?ma de manifeste de compl?ment. D?finissez l??l?ment **App_office** dans votre manifeste compl?ment comme suit.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a>Indication de l?API JavaScript la plus r?cente pour la biblioth?que Office

Si vous utilisez des v?rifications ? l?ex?cution, r?f?rencez la version la plus r?cente de l?API JavaScript pour la biblioth?que Office ? partir du r?seau de livraison de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante ? votre code HTML. L?utilisation de `/1/` dans l?URL CDN garantit que vous r?f?rencez la version d?Office.js la plus r?cente.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Options pour sp?cifier des h?tes Office ou les conditions requises d?API

Lors de la sp?cification des h?tes Office ou des conditions requises d?API, vous devez tenir compte de plusieurs facteurs. Le diagramme suivant montre comment choisir la technique ? utiliser dans votre compl?ment.

![Optez pour la meilleure solution pour votre compl?ment lorsque vous sp?cifiez des h?tes Office ou des exigences d?API](../images/options-for-office-hosts.png)

- Si votre compl?ment s?ex?cute dans un h?te Office, d?finissez l??l?ment **Hosts** dans le manifeste. Pour plus d?informations, consultez [D?finition de l??l?ment Hosts](#set-the-hosts-element).
    
- Pour d?finir l?ensemble minimal de conditions requises ou les membres minimaux d?API qu?un h?te Office doit prendre en charge pour ex?cuter votre compl?ment, d?finissez l??l?ment **Requirements** dans le manifeste. Pour plus d?informations, consultez la section [ D?finition de l??l?ment Requirements dans le manifeste](#set-the-requirements-element-in-the-manifest).
    
- Si vous souhaitez proposer des fonctionnalit?s suppl?mentaires lorsque des ensembles de conditions requises sp?cifiques ou des membres d?API sont disponibles dans l?h?te Office, effectuez une v?rification ? l?ex?cution dans le code JavaScript de votre compl?ment. Par exemple, si votre compl?ment est ex?cut? dans Excel 2016, utilisez les membres d?API de la nouvelle API JavaScript pour Excel pour fournir des fonctionnalit?s suppl?mentaires. Pour plus d?informations, consultez la section [Utilisation des v?rifications ? l?ex?cution dans votre code JavaScript](#use-runtime-checks-in-your-javascript-code).
    
## <a name="set-the-hosts-element"></a>D?finition de l??l?ment Hosts

Pour ex?cuter votre compl?ment dans une application h?te Office, utilisez les ?l?ments **Hosts** et **Host** dans le manifeste. Si vous ne d?finissez pas l??l?ment **Hosts**, votre compl?ment sera ex?cut? dans tous les h?tes.

Par exemple, les d?clarations  **Hosts** et **Host** suivantes indiquent que le compl?ment fonctionnera avec n?importe quelle version d?Excel, y compris Excel pour Windows, Excel Online et Excel pour iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

L??l?ment  **Hosts** peut contenir un ou plusieurs ?l?ments  **Host**. L??l?ment  **Host** indique l?h?te Office requis par votre compl?ment. L?attribut **Name** est requis et peut ?tre d?fini sur l?une des valeurs suivantes.

| Nom          | Applications h?tes Office                      |
|:--------------|:----------------------------------------------|
| Base de donn?es      | applications web Access                               |
| Document      | Word pour Windows, Mac, iPad et Online        |
| Bo?te aux lettres       | Outlook pour Windows, Mac, Web et Outlook.com | 
| Pr?sentation  | PowerPoint pour Windows, Mac, iPad et Online  |
| Projet       | Projet                                       |
| Classeur      | Excel pour Windows, Mac, iPad et Online           |

> [!NOTE]
> L?attribut `Name` sp?cifie l?application h?te Office pouvant ex?cuter votre compl?ment. Les h?tes Office sont pris en charge sur diff?rentes plateformes et sont ex?cut?s sur les ordinateurs de bureau, les navigateurs web, les tablettes et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut ?tre utilis?e pour ex?cuter votre compl?ment. Par exemple, si vous sp?cifiez `Mailbox`, Outlook et Outlook Web App peuvent ?tre utilis?s pour ex?cuter votre compl?ment. 


## <a name="set-the-requirements-element-in-the-manifest"></a>D?finition de l??l?ment Requirements dans le manifeste

L??l?ment **Requirements** indique les ensembles de conditions minimales requises ou les membres d?API qui doivent ?tre pris en charge par l?h?te Office en vue d?ex?cuter votre compl?ment. L??l?ment **Requirements** peut indiquer des ensembles de conditions requises et des m?thodes individuelles utilis?s dans votre compl?ment. Dans la version 1.1 du sch?ma de manifeste du compl?ment, l??l?ment **Requirements** est facultatif pour tous les compl?ments, sauf pour les compl?ments Outlook.

> [!WARNING]
> Utilisez uniquement l??l?ment **Requirements** pour sp?cifier des ensembles de conditions requises essentiels ou des membres d?API que votre compl?ment doit utiliser. Si la plateforme ou l?h?te Office ne prend pas en charge les ensembles de conditions requises ou les membres d?API sp?cifi?s dans l??l?ment **Requirements**, le compl?ment ne s?ex?cute pas dans cet h?te ou cette plateforme et ne s?affiche pas dans **Mes compl?ments**. Nous vous recommandons plut?t de rendre votre compl?ment disponible sur toutes les plateformes d?un h?te Office, comme Excel pour Windows, Excel Online et Excel pour iPad. Pour rendre votre compl?ment disponible sur _tous_ les h?tes et plateformes Office, utilisez des v?rifications ? l?ex?cution ? la place de l??l?ment **Requirements**.

Cet exemple de code illustre un compl?ment qui se charge dans toutes les applications h?tes Office qui prennent en charge les ?l?ments suivants :

-  Un ensemble de conditions requises **TableBindings**, dont la version minimale est 1.1.
    
-  Un ensemble de conditions requises **OOXML**, dont la version minimale est 1.1.
    
-  La m?thode **Document.getSelectedDataAsync**.

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

- L??l?ment **Requirements** contient les ?l?ments enfants **Sets** et **Methods**.
    
- L??l?ment  **Sets** peut contenir un ou plusieurs ?l?ments  **Set**.  **DefaultMinVersion** indique la valeur **MinVersion** par d?faut de tous les ?l?ments  **Set** enfants.
    
- L??l?ment **Set** sp?cifie les ensembles de conditions requises que l?h?te Office doit prendre en charge pour ex?cuter le compl?ment. L?attribut **Name** indique le nom de l?ensemble de conditions requises. L?attribut **MinVersion** sp?cifie la version minimale de l?ensemble de conditions requises. L?attribut **MinVersion** remplace la valeur de **DefaultMinVersion**. Pour plus d?informations sur les ensembles de conditions requises et les versions auxquelles les membres de votre API appartiennent, consultez [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).
    
- L??l?ment **Methods** peut contenir un ou plusieurs ?l?ments **Method**. Vous ne pouvez pas utiliser l??l?ment **Methods** avec des compl?ments Outlook.
    
- L??l?ment  **Method** sp?cifie une m?thode individuelle qui doit ?tre prise en charge dans l?h?te Office o? votre compl?ment est ex?cut?. L?attribut **Name** est obligatoire et indique le nom de la m?thode qualifi?e avec son objet parent.
    

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des v?rifications ? l?ex?cution dans votre code JavaScript


Vous pouvez fournir des fonctionnalit?s suppl?mentaires dans votre compl?ment si certains ensembles de conditions requises sont pris en charge par l?h?te Office. Par exemple, vous pouvez utiliser les nouvelles interfaces API JavaScript de Word dans votre compl?ment existant si ce dernier est ex?cut? dans Word 2016. Pour ce faire, utilisez la m?thode **isSetSupported** avec le nom de l?ensemble de conditions requises. **isSetSupported** d?termine, lors de l?ex?cution, si l?h?te Office ex?cutant le compl?ment prend en charge l?ensemble des conditions requises. Si l?ensemble de conditions requises est pris en charge, **isSetSupported** renvoie **True** et ex?cute le code suppl?mentaire qui utilise les membres d?API provenant de l?ensemble de conditions requises. Si l?h?te Office ne prend pas en charge l?ensemble de conditions requises, **isSetSupported** renvoie **False** et le code suppl?mentaire n?est pas ex?cut?. Le code suivant indique la syntaxe ? utiliser avec **isSetSupported**.


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  _RequirementSetName_ (obligatoire) est une cha?ne repr?sentant le nom de l?ensemble de conditions requises. Pour plus d?informations sur les ensembles de conditions requises disponibles, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).
    
-  _VersionNumber_ (facultatif) correspond ? la version de l?ensemble de conditions requises.
    
Dans Excel 2016 ou Word 2016, utilisez **isSetSupported** avec les ensembles de conditions requises  **ExcelAPI** ou **WordAPI**. La m?thode  **isSetSupported**, ainsi que les ensembles de conditions requises  **ExcelAPI** et **WordAPI**, sont disponibles dans le dernier fichier Office.js du CDN. Si vous n?utilisez pas Office.js ? partir du CDN, votre compl?ment peut g?n?rer des exceptions, car la m?thode  **isSetSupported** ne sera pas d?finie. Pour plus d?informations, voir [ Indication de l?API JavaScript la plus r?cente pour la biblioth?que Office](#specify-the-latest-javascript-api-for-office-library). 


> [!NOTE]
> **isSetSupported** ne fonctionne pas dans Outlook ou Outlook Web App. Pour utiliser une v?rification ? l?ex?cution dans Outlook ou Outlook Web App, utilisez la technique d?crite dans la section [V?rifications ? l?ex?cution ? l?aide de m?thodes ne faisant pas partie d?un ensemble de conditions requises](#runtime-checks-using-methods-not-in-a-requirement-set).

L?exemple de code suivant montre comment un compl?ment peut fournir des fonctionnalit?s diff?rentes pour divers h?tes Office qui peuvent prendre en charge plusieurs ensembles de conditions requises ou membres d?API.




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


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>V?rifications ? l?ex?cution ? l?aide de m?thodes ne faisant pas partie d?un ensemble de conditions requises


Certains membres API n?appartiennent pas ? des ensembles de conditions requises. Cela s?applique uniquement aux membres d?API qui font partie de l?espace de noms de l?[interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) (rien sous Office), et non aux membres d?API qui appartiennent ? l?espace de noms de l?interface API JavaScript pour Word (rien dans Word) ou de la [r?f?rence de l?API JavaScript pour les compl?ments Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (rien dans Excel). Lorsque votre compl?ment d?pend d?une m?thode qui ne fait pas partie d?un ensemble de conditions requises, vous pouvez utiliser la v?rification ? l?ex?cution pour d?terminer si la m?thode est prise en charge par l?h?te Office, comme indiqu? dans l?exemple suivant. Pour consulter la liste compl?te des m?thodes qui n?appartiennent pas ? un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).


> [!NOTE]
> Nous vous recommandons de limiter l?utilisation de ce type de v?rification ? l?ex?cution dans le code de votre compl?ment.

L?exemple de code suivant v?rifie si l?h?te prend en charge **document.setSelectedDataAsync**.




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compl?ments Office](add-in-manifests.md)
- [Ensembles de conditions requises pour les compl?ments Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)