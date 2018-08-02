---
title: Modèle objet API JavaScript pour Office
description: ''
ms.date: 07/27/2018
ms.openlocfilehash: a1fca2c8dd5a92cadb6f2798c3732093578360e5
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21719880"
---
# <a name="office-javascript-api-object-model"></a>Modèle objet API JavaScript pour Office
Les compléments JavaScript Office donnent accès aux fonctionnalités sous-jacentes de l'hôte. Une grande partie de cet accès passe par quelques objets importants. L'objet [Contexte](#context-object) donne accès à l'environnement d'exécution après l'initialisation. L'objet [Document](#document-object) donne le contrôle utilisateur sur un document Excel, PowerPoint ou Word. L'objet [Boite aux lettres](#mailbox-object) donne l'accès aux messages et aux profils utilisateur à un complément Outlook. Comprendre les relations entre ces objets de haut niveau est la base d'un complément JavaScript.

## <a name="context-object"></a>Context, objet

**S’applique à :** tous les types de complément

Lorsqu'un complément est [initialisé](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in), il a de nombreux objets différents avec lesquels il peut interagir dans l'environnement d'exécution. Le contexte d'exécution du complément est reflété dans l'API par l'objet [Contexte](https://dev.office.com/reference/add-ins/shared/office.context). L'objet **Contexte** est l'objet principal qui donne accès aux objets les plus importants de l'API, tels que les objets [Document](https://dev.office.com/reference/add-ins/shared/document) et [Boites aux lettres](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) qui à leur tour, fournissent un accès au contenu du document et de la boîte aux lettres.

Par exemple, dans les compléments de contenu ou du volet Office, vous pouvez utiliser la propriété [document](https://dev.office.com/reference/add-ins/shared/office.context.document) de l’objet **Context** pour accéder aux propriétés et aux méthodes de l’objet **Document** afin d’interagir avec le contenu de documents Word, de feuilles de calcul Excel ou de planifications Project. De même, dans les compléments Outlook, vous pouvez utiliser la propriété [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) de l’objet **Context** pour accéder aux méthodes et aux propriétés de l’objet **Mailbox** afin d’interagir avec le contenu des messages, des demandes de réunion ou des rendez-vous.

L’objet **Context** donne également accès aux propriétés [contentLanguage](https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) et [displayLanguage](https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) qui vous permettent de déterminer les paramètres régionaux (langue) utilisés dans le document ou dans l’élément, ou par l’application hôte. La propriété [roamingSettings](https://dev.office.com/reference/add-ins/outlook/Office.context) vous permet d'accéder aux membres de l'objet [RoamingSettings,](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) qui stocke les paramètres spécifiques à votre complément pour les boîtes aux lettres de chaque utilisateur. Enfin, l’objet **Context** fournit une propriété [ui](https://dev.office.com/reference/add-ins/shared/officeui) qui permet à votre complément de lancer des boîtes de dialogue contextuelles.


## <a name="document-object"></a>Objet Document

**S’applique à :** types de complément de contenu et du volet Office

Pour permettre l’interaction avec les données de document dans Excel, PowerPoint et Word, l’API fournit l’objet [Document](https://dev.office.com/reference/add-ins/shared/document). Vous pouvez utiliser les membres de l’objet  **Document** pour accéder aux données de différentes façons :

- Lecture et écriture dans les sélections actives sous forme de texte, de cellules contiguës (matrices) ou de tableaux.
    
- Données tabulaires (matrices ou tableaux).
    
- Liaisons (créées avec les méthodes « add » de l’objet  **Bindings**).
    
- Parties XML personnalisées (uniquement pour Word).
    
- Paramètres ou état de complément persistant par complément dans le document.
    
Vous pouvez également utiliser l’objet  **Document** pour interagir avec les données des documents Project. La fonctionnalité propre à Project de l’API est documentée dans la classe abstraite [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) des membres. Pour plus d’informations sur la création de compléments du volet Office pour Project, voir [Compléments du volet Office pour Project](../project/project-add-ins.md).

Tous ces types d’accès aux données utilisent une instance de l’objet abstrait  **Document**.

Vous pouvez accéder à une instance de l’objet  **Document** lors de l’initialisation du complément de contenu ou du volet Office en utilisant la propriété [document](https://dev.office.com/reference/add-ins/shared/office.context.document) de l’objet **Context**. L’objet  **Document** définit les fonctions communes d’accès aux données dans les documents Word et Excel, et donne également accès à l’objet **CustomXmlParts** pour les documents Word.

L’objet  **Document** prend en charge quatre moyens pour les développeurs d’accéder au contenu des documents :


- Accès basé sur les sélections
    
- Accès basé sur les liaisons
    
- Accès basé sur les parties XML personnalisées (Word uniquement)
    
- Accès basé sur l’intégralité du document (PowerPoint et Word uniquement)
    
Pour vous aider à comprendre comment fonctionnent les méthodes d’accès aux données par sélection et par liaison, nous expliquerons tout d’abord comment les API d’accès aux données fournissent un accès aux données cohérent parmi les différentes applications Office.


### <a name="consistent-data-access-across-office-applications"></a>Accès aux données cohérent sur les différentes applications Office

 **S’applique à :** types de complément de contenu et du volet Office

Pour créer des extensions qui fonctionnent de manière transparente parmi les différents documents Office, l’API JavaScript pour Office fait abstraction des particularités de chaque application Office par l’intermédiaire des types de données communs et par le forçage de type des différents contenus de documents selon trois types de données communs.


#### <a name="common-data-types"></a>Type de données communs

Dans l’accès aux données basé sur la sélection et basé sur la liaison, les contenus de documents sont exposés via des types de données qui sont communs à toutes les applications Office prises en charge. Dans Office 2013, trois principaux types de données sont pris en charge :



|**Type de données**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Texte|Fournit une représentation sous forme de chaîne des données de la sélection ou de la liaison.|Dans Excel 2013, Project 2013 et PowerPoint 2013, seul le texte brut est pris en charge. Dans Word 2013, trois formats de texte sont pris en charge : texte brut, HTML et Office Open XML (OOXML). Lorsque du texte est sélectionné dans une cellule d’Excel, les méthodes reposant sur la sélection lisent et écrivent le contenu entier de la cellule, même si uniquement une partie du texte est sélectionnée dans la cellule. Lorsque du texte est sélectionné dans Word et PowerPoint, les méthodes reposant sur la sélection lisent et écrivent uniquement la série de caractères sélectionnés. Project 2013 et PowerPoint 2013 prennent uniquement en charge l’accès aux données basé sur les sélections.|
|Matrice|Fournit les données de la sélection ou de la liaison sous forme d’un **tableau** bidimensionnel implémenté dans JavaScript sous forme de tableau de tableaux.Par exemple, deux lignes de valeurs **string** dans deux colonnes donneront ` [['a', 'b'], ['c', 'd']]`, et une seule colonne de trois lignes donnera `[['a'], ['b'], ['c']]`.|L’accès aux données de matrice est pris en charge uniquement dans Excel 2013 et Word 2013.|
|Tableau|Fournit les données dans la sélection ou la liaison sous forme d’objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L’objet  **TableData** expose les données via les propriétés **headers** et **rows**.|L’accès aux données de tableau est pris en charge uniquement dans Excel 2013 et Word 2013.|

#### <a name="data-type-coercion"></a>Contrainte du type de données

Les méthodes d’accès aux données des objets **Document** et [Binding](https://dev.office.com/reference/add-ins/shared/binding) prennent en charge la spécification du type de données voulu à l’aide du paramètre _coercionType_ de ces méthodes, ainsi que les valeurs d’énumération [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) correspondantes. Quelle que soit la forme réelle de la liaison, les différentes applications Office prennent en charge les types de données communs en tentant de forcer le type des données selon le type demandé. Par exemple, si un tableau ou un paragraphe Word est sélectionné, le développeur peut indiquer qu’il souhaite le lire en tant que texte brut, HTML, Office Open XML ou en tant que tableau, et l’implémentation de l’API gère les transformations et conversions de données nécessaires.


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si les données tabulaires doivent croître de façon dynamique lors de l’ajout de lignes et de colonnes, et que vous devez travailler avec des en-têtes de tableaux, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ de la méthode d’accès aux données d’objet **Document** ou **Binding** en tant que `"table"` ou **Office.CoercionType.Table**). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas la fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre _coercionType_ de la méthode d’accès aux données en tant que `"matrix"` ou **Office.CoercionType.Matrix**), qui fournit un modèle plus simple d’interaction avec les données.

Si les données sont d’un type qui ne peut pas être forcé vers le type spécifié, la propriété [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) du rappel renvoie `"failed"`. Par ailleurs, vous pouvez utiliser la propriété [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) pour accéder à un objet [Error](https://dev.office.com/reference/add-ins/shared/error) incluant des informations sur la raison de l’échec de l’appel de la méthode.


## <a name="working-with-selections-using-the-document-object"></a>Utilisation des sélections à l’aide de l’objet Document


L’objet **Document** expose des méthodes qui vous permettent de lire et d’écrire dans la sélection actuelle de l’utilisateur sur le mode « obtenir et définir ». Pour cela, l’objet **Document** fournit les méthodes **getSelectedDataAsync** et **setSelectedDataAsync**.

Pour obtenir des exemples de code montrant comment effectuer des tâches avec les sélections, voir [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Utilisation des liaisons à l’aide des objets Bindings et Binding


L’accès aux données basé sur les liaisons permet aux compléments de contenu et du volet Office d’accéder de manière cohérente à une région particulière d’un document ou d’une feuille de calcul par l’intermédiaire d’un identificateur associé à une liaison. Le complément doit d’abord établir la liaison en appelant l’une des méthodes qui associent une partie du document à un identificateur unique : [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) ou [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). Une fois la liaison établie, le complément peut utiliser l’identificateur fourni pour accéder aux données contenues dans la région associée du document ou de la feuille de calcul. La création de liaisons apporte à votre complément les avantages suivants :


- Elle permet l’accès aux structures de données communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (série contiguë de caractères).
    
- Elle permet les opérations de lecture/écriture sans exiger que l’utilisateur effectue une sélection.
    
- Elle établit une relation entre le complément et les données du document. Les liaisons persistent dans le document et sont accessibles par la suite.
    
L’établissement d’une liaison vous permet également de vous abonner aux données et aux événements de changement de sélection qui sont concernés par cette région particulière du document ou de la feuille de calcul. Cela signifie que le complément est seulement notifié des changements qui surviennent dans la région délimitée, par opposition aux changements généraux affectant l’ensemble du document ou de la feuille de calcul.

L’objet [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) expose une méthode [getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) qui donne accès à toutes les liaisons établies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID à l’aide de la méthode [Bindings.getBindingByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ou [Office.select](https://dev.office.com/reference/add-ins/shared/office.select). Vous pouvez établir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l’une des méthodes suivantes de l’objet  **Bindings** : [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) ou [releaseByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

Vous spécifiez trois types de liaisons différents avec le paramètre  _bindingType_ lorsque vous créez une liaison avec les méthodes **addFromSelectionAsync**, **addFromPromptAsync** ou **addFromNamedItemAsync**  :



|**Type de liaison**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Liaison de texte|Établit une liaison à une zone du document qui est représentée en tant que texte.|Dans Word, la plupart des sélections contiguës sont valides, tandis que dans Excel, seules les sélections de cellules uniques peuvent être la cible d’une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.|
|Liaison de matrices|Établit une liaison à une zone fixe d’un document qui contient des données tabulaires sans en-tête. Les données dans une liaison de matrice sont écrites ou lues comme un **tableau** bidimensionnel, implémenté dans JavaScript sous forme de tableau de tableaux. Par exemple, deux lignes de valeurs **string** dans deux colonnes peuvent être écrites ou lues comme ` [['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut être écrite ou lue comme `[['a'], ['b'], ['c']]`.|Dans Excel, toute sélection contiguë de cellules peut être utilisée pour établir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.|
|Liaison de tableaux|Établit une liaison à une zone d’un document qui contient un tableau avec des en-têtes. Les données dans une liaison de tableau sont écrites ou lues comme un objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L’objet **TableData** expose les données via les propriétés **headers** et **rows**.|Tout tableau Excel ou Word peut être la base d’une liaison de tableau. Une fois que vous établissez une liaison de tableau, chaque nouvelle ligne ou colonne qu’un utilisateur ajoute au tableau est automatiquement incluse dans la liaison. |

<br/>

Une fois la liaison créée à l’aide de l’une des trois méthodes « add » de l’objet  **Bindings**, vous pouvez travailler avec les données et les propriétés de la liaison en utilisant les méthodes de l’objet correspondant : [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) ou [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding). Ces trois objets héritent des méthodes [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) et [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync) de l’objet **Binding** qui vous permettent d’interagir avec les données liées.

Pour obtenir des exemples de code qui montrent comment effectuer des tâches avec les liaisons, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Utilisation de parties XML personnalisées à l’aide des objets CustomXmlParts et CustomXmlPart


 **S’applique à :** compléments du volet Office pour Word et PowerPoint

Les objets [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) et [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) de l’API donnent accès à des parties XML personnalisées dans les documents Word, qui permettent une manipulation orientée XML du contenu du document. Pour une démonstration de l’utilisation des objets **CustomXmlParts** et **CustomXmlPart**, voir l’exemple de code [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Utilisation de l’intégralité du document à l’aide de la méthode getFileAsync


 **S’applique à :** compléments du volet Office pour Word et PowerPoint

La méthode [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) et les membres des objets [File](https://dev.office.com/reference/add-ins/shared/file) et [Slice](https://dev.office.com/reference/add-ins/shared/slice) fournissent les fonctionnalités permettant d’obtenir l’intégralité des fichiers Word et PowerPoint sous forme de sections (blocs) de 4 Mo maximum à la fois. Pour plus d’informations, reportez-vous à la rubrique [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="mailbox-object"></a>Objet Mailbox


 **S’applique à :** compléments Outlook

Les compléments Outlook utilisent principalement un sous-ensemble de l’API exposée via l’objet [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox). Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item), utilisez la propriété [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

De plus, les compléments Outlook peuvent utiliser les objets suivants :


-  Objet **Office** : pour l’initialisation.
    
-  Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.
    
-  Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.
    
Pour plus d’informations sur l’utilisation de JavaScript dans les compléments Outlook, reportez-vous à la rubrique [Compléments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/).