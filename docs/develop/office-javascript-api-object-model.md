---
title: Modèle d’objet d’API JavaScript courant
description: En savoir plus sur le modèle Office’objet API courant JavaScript
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: dd30f5e5be70f58fec9eb4c84c0491397792950b
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774209"
---
# <a name="common-javascript-api-object-model"></a>Modèle d’objet d’API JavaScript courant

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Les API JavaScript donnent accès à Office la fonctionnalité sous-jacente de l’application cliente. La majeure partie de cet accès passe par quelques objets importants. L’objet [Context](#context-object) donne accès à l’environnement d’exécution après l’initialisation. L’objet [Document](#document-object)donne à l’utilisateur le contrôle d’un document Excel, PowerPoint ou Word. [L’objet Mailbox](#mailbox-object) donne à un Outlook un accès aux messages, rendez-vous et profils utilisateur. Comprendre les relations entre ces objets de haut niveau constitue le fondement d’un Office de niveau supérieur.

## <a name="context-object"></a>Context, objet

**S’applique à :** tous les types de complément

Lorsqu’un complément est [initialisé](initialize-add-in.md), il peut interagir avec de nombreux objets différents dans l’environnement d’exécution. Le contexte du runtime du complément est indiqué dans l’API par l’objet [Context](/javascript/api/office/office.context). **Context** est l’objet principal qui permet d’accéder aux objets les plus importants de l’API, tels que les objets [Document](/javascript/api/office/office.document) et [Mailbox](/javascript/api/outlook/office.mailbox), qui à leur tour donnent accès au contenu des documents et boîtes aux lettres.

Par exemple, dans les compléments de contenu ou du volet Office, vous pouvez utiliser la propriété document de l’objet Context pour accéder aux propriétés et aux méthodes de l’objet Document afin d’interagir avec le contenu de documents Word, de feuilles de calcul Excel ou de planifications Project. De même, dans les compléments Outlook, vous pouvez utiliser la propriété mailbox de l’objet Context pour accéder aux méthodes et aux propriétés de l’objet Mailbox afin d’interagir avec le contenu des messages, des demandes de réunion ou des rendez-vous.

L’objet **Context** permet également d’accéder aux propriétés [contentLanguage](/javascript/api/office/office.context#contentLanguage) et [displayLanguage](/javascript/api/office/office.context#displayLanguage) qui vous permet de déterminer les paramètres régionaux (langue) utilisés dans le document ou l’élément, ou par l’application Office. La propriété [roamingSettings](/javascript/api/office/office.context#roamingSettings) vous permet d’accéder aux membres de l’objet [RoamingSettings](/javascript/api/office/office.context#roamingSettings) qui stocke les paramètres spécifiques à votre complément pour les boîtes aux lettres individuelles des utilisateurs. Enfin, l’objet **Context** fournit une propriété [ui](/javascript/api/office/office.context#ui) qui permet à votre complément d’ouvrir des boîtes de dialogue contextuelles.

## <a name="document-object"></a>Objet Document

**S’applique à :** types de complément de contenu et du volet Office

Pour permettre l’interaction avec les données de document dans Excel, PowerPoint et Word, l’API fournit l’objet [Document](/javascript/api/office/office.document). Vous pouvez utiliser les membres `Document` de l’objet pour accéder aux données des manières suivantes.

- Lecture et écriture dans les sélections actives sous forme de texte, de cellules contiguës (matrices) ou de tableaux.

- Données tabulaires (matrices ou tableaux).

- Liaisons (créées avec les méthodes « add » de `Bindings` l’objet).

- Parties XML personnalisées (uniquement pour Word).

- Paramètres ou état de complément persistant par complément dans le document.

Vous pouvez également utiliser `Document` l’objet pour interagir avec les données dans Project documents. La fonctionnalité propre à Project de l’API est documentée dans la classe abstraite [ProjectDocument](/javascript/api/office/office.document) des membres. Pour plus d’informations sur la création de compléments du volet Office pour Project, voir [Compléments du volet Office pour Project](../project/project-add-ins.md).

Toutes ces formes d’accès aux données commencent par une instance de l’objet `Document` abstrait.

Vous pouvez accéder à une instance de l’objet lorsque le volet Des tâches ou le module de contenu est initialisé à l’aide de la propriété `Document` [de document](/javascript/api/office/office.context#document) de `Context` l’objet. L’objet définit les fonctions communes d’accès aux données partagées entre les documents Word et Excel documents, et fournit également l’accès à l’objet pour `Document` `CustomXmlParts` les documents Word.

`Document`L’objet prend en charge quatre façons pour les développeurs d’accéder au contenu du document.

- Accès basé sur les sélections

- Accès basé sur les liaisons

- Accès basé sur les parties XML personnalisées (Word uniquement)

- Accès basé sur l’intégralité du document (PowerPoint et Word uniquement)

Pour vous aider à comprendre comment fonctionnent les méthodes d’accès aux données par sélection et par liaison, nous expliquerons tout d’abord comment les API d’accès aux données fournissent un accès aux données cohérent parmi les différentes applications Office.

### <a name="consistent-data-access-across-office-applications"></a>Accès aux données cohérent sur les différentes applications Office

 **S’applique à :** types de complément de contenu et du volet Office

Pour créer des extensions qui fonctionnent en toute transparence sur différents documents Office, l’API JavaScript Office extrait les particularités de chaque application Office par le biais de types de données courants et la possibilité de forcer différents contenus de document en trois types de données courants.

#### <a name="common-data-types"></a>Type de données communs

Dans l’accès aux données basé sur la sélection et basé sur la liaison, les contenus de documents sont exposés via des types de données qui sont communs à toutes les applications Office prises en charge. Dans Office 2013, trois principaux types de données sont pris en charge.

|**Type de données**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Texte|Fournit une représentation sous forme de chaîne des données de la sélection ou de la liaison.|Dans Excel 2013, Project 2013 et PowerPoint 2013, seul le texte brut est pris en charge. Dans Word 2013, trois formats de texte sont pris en charge : texte brut, HTML et Office Open XML (OOXML). Lorsque du texte est sélectionné dans une cellule d’Excel, les méthodes reposant sur la sélection lisent et écrivent le contenu entier de la cellule, même si uniquement une partie du texte est sélectionnée dans la cellule. Lorsque du texte est sélectionné dans Word et PowerPoint, les méthodes reposant sur la sélection lisent et écrivent uniquement la série de caractères sélectionnés. Project 2013 et PowerPoint 2013 prennent uniquement en charge l’accès aux données basé sur les sélections.|
|Matrice|Fournit les données de la sélection ou de la liaison sous forme d’un **tableau** bidimensionnel implémenté dans JavaScript sous forme de tableau de tableaux.Par exemple, deux lignes de valeurs **string** dans deux colonnes donneront ` [['a', 'b'], ['c', 'd']]`, et une seule colonne de trois lignes donnera `[['a'], ['b'], ['c']]`.|L’accès aux données de matrice est pris en charge uniquement dans Excel 2013 et Word 2013.|
|Tableau|Fournit les données dans la sélection ou la liaison sous forme d’objet [TableData](/javascript/api/office/office.tabledata). L’objet  `TableData` expose les données via les propriétés `headers` et `rows`.|L’accès aux données de tableau est pris en charge uniquement dans Excel 2013 et Word 2013.|

#### <a name="data-type-coercion"></a>Contrainte du type de données

Les méthodes d’accès aux données des objets `Document` et [Binding](/javascript/api/office/office.binding) prennent en charge la spécification du type de données voulu à l’aide du paramètre _coercionType_ de ces méthodes, ainsi que les valeurs d’énumération [CoercionType](/javascript/api/office/office.coerciontype) correspondantes. Quelle que soit la forme réelle de la liaison, les différentes applications Office prennent en charge les types de données communs en tentant de forcer le type des données selon le type demandé. Par exemple, si un tableau ou un paragraphe Word est sélectionné, le développeur peut indiquer qu’il souhaite le lire en tant que texte brut, HTML, Office Open XML ou en tant que tableau, et l’implémentation de l’API gère les transformations et conversions de données nécessaires.

> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si vous avez besoin que vos données tabulaires s’développent dynamiquement lorsque des lignes et des colonnes sont ajoutées et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données de table (en spécifiant le paramètre _coercionType_ d’une méthode d’accès aux données objet ou en tant que ou `Document` `Binding` `"table"` `Office.CoercionType.Table` ). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre  _coercionType_ de la méthode d’accès aux données en tant que ou ), ce qui fournit un modèle plus simple d’interaction avec les `"matrix"` `Office.CoercionType.Matrix` données.

Si les données sont d’un type qui ne peut pas être forcé vers le type spécifié, la propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#status) du rappel renvoie `"failed"`. Par ailleurs, vous pouvez utiliser la propriété [AsyncResult.error](/javascript/api/office/office.asyncresult#error) pour accéder à un objet [Error](/javascript/api/office/office.error) incluant des informations sur la raison de l’échec de l’appel de la méthode.

## <a name="work-with-selections-using-the-document-object"></a>Utiliser des sélections à l’aide de l’objet Document

L’objet expose des méthodes qui vous permet de lire et d’écrire dans la sélection actuelle de l’utilisateur de manière « obtenir et `Document` définir ». Pour ce faire, `Document` l’objet fournit les `getSelectedDataAsync` méthodes et les `setSelectedDataAsync` méthodes.

Pour obtenir des exemples de code montrant comment effectuer des tâches avec les sélections, voir [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

## <a name="work-with-bindings-using-the-bindings-and-binding-objects"></a>Utiliser des liaisons à l’aide des objets Bindings et Binding

L’accès aux données basé sur les liaisons permet aux compléments de contenu et du volet Office d’accéder de manière cohérente à une région particulière d’un document ou d’une feuille de calcul par l’intermédiaire d’un identificateur associé à une liaison. Le complément doit d’abord établir la liaison en appelant l’une des méthodes qui associent une partie du document à un identificateur unique : [addFromPromptAsync](/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_), [addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) ou [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_). Une fois la liaison établie, le complément peut utiliser l’identificateur fourni pour accéder aux données contenues dans la région associée du document ou de la feuille de calcul. La création de liaisons fournit la valeur suivante à votre add-in.

- Elle permet l’accès aux structures de données communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (série contiguë de caractères).

- Elle permet les opérations de lecture/écriture sans exiger que l’utilisateur effectue une sélection.

- Elle établit une relation entre le complément et les données du document. Les liaisons persistent dans le document et sont accessibles par la suite.

L’établissement d’une liaison vous permet également de vous abonner aux données et aux événements de changement de sélection qui sont concernés par cette région particulière du document ou de la feuille de calcul. Cela signifie que le complément est seulement notifié des changements qui surviennent dans la région délimitée, par opposition aux changements généraux affectant l’ensemble du document ou de la feuille de calcul.

L’objet [Bindings](/javascript/api/office/office.bindings) expose une méthode [getAllAsync](/javascript/api/office/office.bindings#getAllAsync_options__callback_) qui donne accès à toutes les liaisons établies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID à l’aide de la méthode [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) ou [Office.select](/javascript/api/office). Vous pouvez établir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l’une des méthodes suivantes de l’objet  `Bindings` : [addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_), [addFromPromptAsync](/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_), [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_) ou [releaseByIdAsync](/javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_).

Il existe trois types de liaisons que vous spécifiez avec le paramètre  _bindingType_ lorsque vous créez une liaison avec `addFromSelectionAsync` la ou les `addFromPromptAsync` `addFromNamedItemAsync` méthodes.

|**Type de liaison**|**Description**|**Prise en charge d’application hôte**|
|:-----|:-----|:-----|
|Liaison de texte|Établit une liaison à une zone du document qui est représentée en tant que texte.|Dans Word, la plupart des sélections contiguës sont valides, tandis que dans Excel, seules les sélections de cellules uniques peuvent être la cible d’une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.|
|Matrix binding|Établit une liaison à une zone fixe d’un document qui contient des données tabulaires sans en-tête. Les données dans une liaison de matrice sont écrites ou lues comme un **tableau** bidimensionnel, implémenté dans JavaScript sous forme de tableau de tableaux. Par exemple, deux lignes de valeurs **string** dans deux colonnes peuvent être écrites ou lues comme ` [['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut être écrite ou lue comme `[['a'], ['b'], ['c']]`.|Dans Excel, toute sélection contiguë de cellules peut être utilisée pour établir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.|
|Table binding|Établit une liaison à une zone d’un document qui contient un tableau avec des en-têtes. Les données dans une liaison de tableau sont écrites ou lues comme un objet [TableData](/javascript/api/office/office.tabledata). L’objet `TableData` expose les données via les propriétés **headers** et **rows**.|Tout tableau Excel ou Word peut être la base d’une liaison de tableau. Une fois que vous établissez une liaison de tableau, chaque nouvelle ligne ou colonne qu’un utilisateur ajoute au tableau est automatiquement incluse dans la liaison. |

<br/>

Une fois qu’une liaison est créée à l’aide de l’une des trois méthodes « add » de l’objet, vous pouvez travailler avec les données et propriétés de la liaison à l’aide des méthodes de `Bindings` l’objet correspondant : [MatrixBinding](/javascript/api/office/office.matrixbinding), [TableBinding](/javascript/api/office/office.tablebinding)ou [TextBinding](/javascript/api/office/office.textbinding). Ces trois objets héritent des méthodes [getDataAsync](/javascript/api/office/office.binding#getDataAsync_options__callback_) et [setDataAsync](/javascript/api/office/office.binding#setDataAsync_data__options__callback_) de l’objet qui vous permettent d’interagir avec les données `Binding` liées.

Pour obtenir des exemples de code qui montrent comment effectuer des tâches avec les liaisons, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).

## <a name="work-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Utiliser des parties XML personnalisées à l’aide des objets CustomXmlParts et CustomXmlPart

 **S’applique à :** compléments du volet Office pour Word

Les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts) et [CustomXmlPart](/javascript/api/office/office.customxmlpart) de l’API donnent accès à des parties XML personnalisées dans les documents Word, qui permettent une manipulation orientée XML du contenu du document. Pour des démonstrations de l’utilisation des objets et des éléments, voir l’exemple de `CustomXmlParts` `CustomXmlPart` code [Word-add-in-Work-with-custom-XML-parts.](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts)

## <a name="work-with-the-entire-document-using-the-getfileasync-method"></a>Utiliser l’intégralité du document à l’aide de la méthode getFileAsync

 **S’applique à :** compléments du volet Office pour Word et PowerPoint

La méthode [Document.getFileAsync](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) et les membres des objets [File](/javascript/api/office/office.file) et [Slice](/javascript/api/office/office.slice) fournissent les fonctionnalités permettant d’obtenir l’intégralité des fichiers Word et PowerPoint sous forme de sections (blocs) de 4 Mo maximum à la fois. Pour plus d’informations, reportez-vous à la rubrique [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="mailbox-object"></a>Objet Mailbox

**S’applique à :** compléments Outlook

Les compléments Outlook utilisent principalement un sous-ensemble de l’API exposée via l’objet [Mailbox](/javascript/api/outlook/office.mailbox). Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](/javascript/api/outlook/office.item), utilisez la propriété [mailbox](/javascript/api/office/office.context#mailbox) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

En outre, Outlook compléments peuvent utiliser les objets suivants.

- `Office` objet : pour l’initialisation.

- `Context` objet : pour l’accès au contenu et l’affichage des propriétés de langue.

- `RoamingSettings`objet : pour enregistrer Outlook paramètres personnalisés spécifiques au add-in dans la boîte aux lettres de l’utilisateur où le module est installé.

Pour plus d’informations sur l’utilisation de JavaScript dans les compléments Outlook, reportez-vous à la rubrique [Compléments Outlook](../outlook/outlook-add-ins-overview.md).
