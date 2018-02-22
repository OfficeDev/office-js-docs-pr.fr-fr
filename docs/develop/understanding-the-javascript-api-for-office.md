---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 01/23/2018
---


# <a name="understanding-the-javascript-api-for-office"></a>Présentation de l’API JavaScript pour Office

Cet article fournit des informations sur l’API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’API JavaScript pour Office, voir [Mettre à jour la version de votre API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/fr-fr/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/fr-fr/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Référence à la bibliothèque de l’interface API JavaScript pour Office dans votre complément

La bibliothèque de l’[interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément pour garantir qu’elle utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.

Pour obtenir plus d’informations sur le CDN Office.js et la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Office.js fournit un événement d’initialisation qui se déclenche lorsque l’API est entièrement chargée et prête à interagir avec l’utilisateur. Vous pouvez utiliser le gestionnaire d’événements **initialize** afin de mettre en œuvre des scénarios d’initialisation de complément courants, comme inviter l’utilisateur à sélectionner des cellules dans Excel, puis insérer un graphique initialisé avec les valeurs sélectionnées. Vous pouvez également utiliser le gestionnaire d’événements initialize pour initialiser d’autres logiques personnalisées pour votre complément, telles que l’établissement de liaisons, la demande de valeurs de paramètres de complément par défaut, et ainsi de suite.

Voici à quoi ressemblerait l’événement initialize :     

```js
Office.initialize = function () { };
```
Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être placées dans l’événement Office.initialize. Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Toutes les pages au sein d’un complément Office sont nécessaires pour attribuer un gestionnaire d’événements à l’événement initialize, **Office.initialize**. Si vous ne parvenez pas à attribuer un gestionnaire d’événements, votre complément peut générer une erreur lors de son démarrage. En outre, si un utilisateur essaie d’utiliser votre complément avec un client web Office Online, notamment Excel Online, PowerPoint Online ou Outlook Web App, l’exécution du complément échouera. Si vous n’avez pas besoin de code d’initialisation, le corps de la fonction attribuée à **Office.initialize** peut être vide, comme dans le premier exemple ci-dessus.

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Initialisation du paramètre Reason
Pour les compléments de contenu et du volet Office, Office.initialize fournit un paramètre _reason_ supplémentaire. Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif. Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) et à l’[énumération InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration) 

## <a name="context-object"></a>Context, objet

**S’applique à :** tous les types de complément

Lorsqu’un complément est initialisé, il peut interagir avec plusieurs objets dans l’environnement d’exécution. Le contexte d’exécution du complément est indiqué dans l’API par l’objet [Context](https://dev.office.com/reference/add-ins/shared/office.context). **Context** est l’objet principal qui permet d’accéder aux objets les plus importants de l’API (par exemple, les objets [Document](https://dev.office.com/reference/add-ins/shared/document) et [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox)) qui, quant à eux, donnent accès au contenu des documents et de la boîte aux lettres.

Par exemple, dans les compléments de contenu ou du volet Office, vous pouvez utiliser la propriété [document](https://dev.office.com/reference/add-ins/shared/office.context.document) de l’objet **Context** pour accéder aux propriétés et aux méthodes de l’objet **Document** afin d’interagir avec le contenu de documents Word, de feuilles de calcul Excel ou de planifications Project. De même, dans les compléments Outlook, vous pouvez utiliser la propriété [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) de l’objet **Context** pour accéder aux méthodes et aux propriétés de l’objet **Mailbox** afin d’interagir avec le contenu des messages, des demandes de réunion ou des rendez-vous.

L’objet **Context** donne également accès aux propriétés [contentLanguage](https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) et [displayLanguage](https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) qui vous permettent de déterminer les paramètres régionaux (langue) utilisés dans le document ou l’élément, ou par l’application hôte. La propriété [roamingSettings](https://dev.office.com/reference/add-ins/outlook/Office.context) vous permet d’accéder aux membres de l’objet [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings). Enfin, l’objet **Context** fournit une propriété [ui](https://dev.office.com/reference/add-ins/shared/officeui) qui permet à votre complément de lancer des boîtes de dialogue contextuelles.


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

Dans l’accès aux données basé sur la sélection et basé sur la liaison, les contenus de documents sont exposés via des types de données qui sont communs à toutes les applications Office prises en charge. Dans Office 2013, trois principaux types de données sont pris en charge :



|**Type de données**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Texte|Fournit une représentation sous forme de chaîne des données de la sélection ou de la liaison.|Dans Excel 2013, Project 2013 et PowerPoint 2013, seul le texte brut est pris en charge. Dans Word 2013, trois formats de texte sont pris en charge : texte brut, HTML et Office Open XML (OOXML). Lorsque du texte est sélectionné dans une cellule d’Excel, les méthodes reposant sur la sélection lisent et écrivent le contenu entier de la cellule, même si uniquement une partie du texte est sélectionnée dans la cellule. Lorsque du texte est sélectionné dans Word et PowerPoint, les méthodes reposant sur la sélection lisent et écrivent uniquement la série de caractères sélectionnés. Project 2013 et PowerPoint 2013 prennent uniquement en charge l’accès aux données basé sur les sélections.|
|Matrice|Fournit les données de la sélection ou de la liaison sous forme d’un **tableau** bidimensionnel implémenté dans JavaScript sous forme de tableau de tableaux.Par exemple, deux lignes de valeurs **string** dans deux colonnes donneront ` [['a', 'b'], ['c', 'd']]`, et une seule colonne de trois lignes donnera `[['a'], ['b'], ['c']]`.|L’accès aux données de matrice est pris en charge uniquement dans Excel 2013 et Word 2013.|
|Tableau|Fournit les données dans la sélection ou la liaison sous forme d’objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L’objet  **TableData** expose les données via les propriétés **headers** et **rows**.|L’accès aux données de tableau est pris en charge uniquement dans Excel 2013 et Word 2013.|

#### <a name="data-type-coercion"></a>Contrainte du type de données

Les méthodes d’accès aux données des objets **Document** et [Binding](https://dev.office.com/reference/add-ins/shared/binding) prennent en charge la spécification du type de données voulu à l’aide du paramètre _coercionType_ de ces méthodes, ainsi que les valeurs d’énumération [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) correspondantes. Quelle que soit la forme réelle de la liaison, les différentes applications Office prennent en charge les types de données communs en tentant de forcer le type des données selon le type demandé. Par exemple, si un tableau ou un paragraphe Word est sélectionné, le développeur peut indiquer qu’il souhaite le lire en tant que texte brut, HTML, Office Open XML ou en tant que tableau, et l’implémentation de l’API gère les transformations et conversions de données nécessaires.


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si les données tabulaires doivent croître de façon dynamique lors de l’ajout de lignes et de colonnes, et que vous devez travailler avec des en-têtes de tableaux, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ de la méthode d’accès aux données d’objet **Document** ou **Binding** en tant que `"table"` ou **Office.CoercionType.Table**). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas la fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre _coercionType_ de la méthode d’accès aux données en tant que `"matrix"` ou **Office.CoercionType.Matrix**), qui fournit un modèle plus simple d’interaction avec les données.

Si les données sont d’un type qui ne peut pas être forcé vers le type spécifié, la propriété [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) du rappel renvoie `"failed"`. Par ailleurs, vous pouvez utiliser la propriété [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) pour accéder à un objet [Error](https://dev.office.com/reference/add-ins/shared/error) incluant des informations sur la raison de l’échec de l’appel de la méthode.


## <a name="working-with-selections-using-the-document-object"></a>Utilisation des sélections à l’aide de l’objet Document


L’objet **Document** expose des méthodes qui vous permettent de lire et d’écrire dans la sélection actuelle de l’utilisateur sur le mode « obtenir et définir ». Pour cela, l’objet **Document** fournit les méthodes **getSelectedDataAsync** et **setSelectedDataAsync**.

Pour obtenir des exemples de code montrant comment effectuer des tâches avec les sélections, voir [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Utilisation des liaisons à l’aide des objets Bindings et Binding


L’accès aux données basé sur les liaisons permet aux compléments de contenu et du volet Office d’accéder de manière cohérente à une région particulière d’un document ou d’une feuille de calcul par l’intermédiaire d’un identificateur associé à une liaison. Le complément doit d’abord établir la liaison en appelant l’une des méthodes qui associent une partie du document à un identificateur unique : [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) ou [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). Une fois la liaison établie, le complément peut utiliser l’identificateur fourni pour accéder aux données contenues dans la région associée du document ou de la feuille de calcul. La création de liaisons apporte à votre complément les avantages suivants :


- Elle permet l’accès aux structures de données communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (série contiguë de caractères).
    
- Elle permet les opérations de lecture/écriture sans exiger que l’utilisateur effectue une sélection.
    
- Elle établit une relation entre le complément et les données du document. Les liaisons persistent dans le document et sont accessibles par la suite.
    
L’établissement d’une liaison vous permet également de vous abonner aux données et aux événements de changement de sélection qui sont concernés par cette région particulière du document ou de la feuille de calcul. Cela signifie que le complément est seulement notifié des changements qui surviennent dans la région délimitée, par opposition aux changements généraux affectant l’ensemble du document ou de la feuille de calcul.

L’objet [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) expose une méthode [getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) qui donne accès à toutes les liaisons établies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID à l’aide de la méthode [Bindings.getBindingByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ou [Office.select](https://dev.office.com/reference/add-ins/shared/office.select). Vous pouvez établir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l’une des méthodes suivantes de l’objet  **Bindings** : [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) ou [releaseByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

Vous spécifiez trois types de liaisons différents avec le paramètre  _bindingType_ lorsque vous créez une liaison avec les méthodes **addFromSelectionAsync**, **addFromPromptAsync** ou **addFromNamedItemAsync**  :



|**Type de liaison**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Liaison de texte|Établit une liaison à une zone du document qui est représentée en tant que texte.|Dans Word, la plupart des sélections contiguës sont valides, tandis que dans Excel, seules les sélections de cellules uniques peuvent être la cible d’une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.|
|Matrix binding|Établit une liaison à une zone fixe d’un document qui contient des données tabulaires sans en-tête. Les données dans une liaison de matrice sont écrites ou lues comme un **tableau** bidimensionnel, implémenté dans JavaScript sous forme de tableau de tableaux. Par exemple, deux lignes de valeurs **string** dans deux colonnes peuvent être écrites ou lues comme ` [['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut être écrite ou lue comme `[['a'], ['b'], ['c']]`.|Dans Excel, toute sélection contiguë de cellules peut être utilisée pour établir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.|
|Table binding|Établit une liaison à une zone d’un document qui contient un tableau avec des en-têtes. Les données dans une liaison de tableau sont écrites ou lues comme un objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L’objet **TableData** expose les données via les propriétés **headers** et **rows**.|Tout tableau Excel ou Word peut être la base d’une liaison de tableau. Une fois que vous établissez une liaison de tableau, chaque nouvelle ligne ou colonne qu’un utilisateur ajoute au tableau est automatiquement incluse dans la liaison. |

<br/>

Une fois la liaison créée à l’aide de l’une des trois méthodes « add » de l’objet  **Bindings**, vous pouvez travailler avec les données et les propriétés de la liaison en utilisant les méthodes de l’objet correspondant : [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) ou [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding). Ces trois objets héritent des méthodes [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) et [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync) de l’objet **Binding** qui vous permettent d’interagir avec les données liées.

Pour obtenir des exemples de code qui montrent comment effectuer des tâches avec les liaisons, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Utilisation de parties XML personnalisées à l’aide des objets CustomXmlParts et CustomXmlPart


 **S’applique à :** compléments du volet Office pour Word et PowerPoint

Les objets [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) et [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) de l’API donnent accès à des parties XML personnalisées dans les documents Word, qui permettent une manipulation orientée XML du contenu du document. Pour une démonstration de l’utilisation des objets **CustomXmlParts** et **CustomXmlPart**, voir l’exemple de code [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Utilisation de l’intégralité du document à l’aide de la méthode getFileAsync


 **S’applique à :** compléments du volet Office pour Word et PowerPoint

La méthode [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) et les membres des objets [File](https://dev.office.com/reference/add-ins/shared/file) et [Slice](https://dev.office.com/reference/add-ins/shared/slice) fournissent les fonctionnalités permettant d’obtenir l’intégralité des fichiers Word et PowerPoint sous forme de sections (blocs) de 4 Mo maximum à la fois. Pour plus d’informations, reportez-vous à la rubrique [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="mailbox-object"></a>Objet Mailbox


 **S’applique à :** compléments Outlook

Les compléments Outlook utilisent principalement un sous-ensemble de l’API exposée via l’objet [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox). Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item), utilisez la propriété [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

De plus, les compléments Outlook peuvent utiliser les objets suivants :


-  Objet **Office** : pour l’initialisation.
    
-  Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.
    
-  Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.
    
Pour plus d’informations sur l’utilisation de JavaScript dans les compléments Outlook, reportez-vous à la rubrique [Compléments Outlook](https://docs.microsoft.com/fr-fr/outlook/add-ins/).


## <a name="api-support-matrix"></a>Matrice de prise en charge d’API


Ce tableau récapitule l’API et les fonctionnalités prises en charge dans les types de complément (contenu, volet Office et Outlook), ainsi que les applications Office qui peuvent les héberger lorsque vous indiquez les applications hôte Office prises en charge par votre complément à l’aide du [schéma de manifeste de complément 1.1 et des fonctionnalités prises en charge par la version 1.1 de l’interface API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nom de l’hôte**|Base de données|Classeur|Boîte aux lettres|Présentation|Document|Projet|
||**Applications hôtes** **prises en charge**|Applications web Access|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA pour les périphériques|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Types de compléments pris en charge**|Contenu|v|v||v|||
||Volet de tâches||v||v|v|v|
||Outlook|||O||||
|**Fonctionnalités d’API prises en charge**|Lecture/écriture de texte||v||v|v|v<br/>(En lecture seule)|
||Lecture/Écriture de matrice||v|||v||
||Lecture/écriture de tableau||v|||v||
||Lecture/écriture HTML|||||v||
||Lecture/Écriture<br/>Office Open XML|||||v||
||Lecture des propriétés de tâche, de ressource, de vue et de champ||||||v|
||Événements modifiés de sélection||v|||v||
||Obtention de l’ensemble du document||||v|v||
||Liaisons et événements de liaison|v<br/>(Liaisons de tableau complètes et partielles uniquement)|v|||v||
||Lecture/écriture des parties XML personnalisées|||||v||
||Faire persister les données d’état de complément (paramètres)|v<br/>(Par complément hôte)|v<br/>(Par document)|v<br/>(Par boîte aux lettres)|v<br/>(Par document)|v<br/>(Par document)||
||Événements modifiés de paramètres|v|v||v|v||
||Obtention du mode de vue active<br/>et affichage des événements modifiés||||v|||
||Accès à des emplacements<br/>dans le document||v||v|v||
||Activation en fonction du contexte<br/>à l’aide de règles et de RegEx|||v||||
||Lecture des propriétés d’élément|||v||||
||Lecture de profil utilisateur|||v||||
||Obtention des pièces jointes|||v||||
||Obtention du jeton d’identité d’utilisateur|||v||||
||Appel des services web Exchange|||v||||
