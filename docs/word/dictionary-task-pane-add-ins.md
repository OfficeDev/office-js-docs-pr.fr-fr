---
title: Créer un complément dictionnaire du volet Office
description: Découvrez comment créer un complément de volet office de dictionnaire.
ms.date: 09/26/2019
ms.localizationpriority: medium
ms.openlocfilehash: 7b6df6ec5e3fc90899475e3fd089a8e5c0ca766b
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187321"
---
# <a name="create-a-dictionary-task-pane-add-in"></a>Créer un complément dictionnaire du volet Office

Cet article présente un exemple de complément du volet Office et d’un service web associé qui fournissent des définitions de dictionnaire ou des entrées du dictionnaire des synonymes sur la sélection actuelle de l’utilisateur dans un document Word 2013.

Une Complément Office de dictionnaire est basée sur le complément du volet Office standard avec des fonctionnalités supplémentaires pour prendre en charge l’interrogation et l’affichage de définitions à partir d’un service web XML de dictionnaire à des endroits supplémentaires dans l’interface utilisateur du complément Office.

Dans un complément du volet Office classique, un utilisateur sélectionne un mot ou une expression dans son document, puis la logique JavaScript sous-jacente du complément transmet cette sélection au service web XML du fournisseur de dictionnaire. La page web du fournisseur de dictionnaire s’actualise ensuite pour afficher les définitions de la sélection pour l’utilisateur. Le composant du service web XML renvoie jusqu’à trois définitions dans le format défini par le schéma XML OfficeDefinitions, qui sont ensuite affichées à l’utilisateur à d’autres endroits dans l’interface utilisateur de l’application Office d’hébergement. La figure 1 illustre l’expérience de sélection et d’affichage pour un complément de dictionnaire Bing s’exécutant dans Word 2013.

*Figure 1. Complément de dictionnaire affichant des définitions pour le mot sélectionné*

![Application dictionnaire affichant une définition.](../images/dictionary-agave-01.jpg)

Il vous incombe de déterminer si la sélection du lien **Afficher plus** dans l’interface utilisateur HTML du complément de dictionnaire affiche plus d’informations dans le volet Office ou ouvre une fenêtre de navigateur distincte sur la page web complète du mot ou de l’expression sélectionné.
La figure 2 montre la commande **Définir** dans le menu contextuel qui permet aux utilisateurs de lancer rapidement des dictionnaires installés. Les figures 3 à 5 montrent les endroits dans l’interface utilisateur d’Office où les services XML de dictionnaire sont utilisés pour fournir des définitions dans Word 2013.

*Figure 2. Commande Définir dans le menu contextuel*

![Définir le menu contextuel.](../images/dictionary-agave-02.jpg)


*Figure 3. Définitions dans les volets Orthographe et Grammaire*

![Définitions dans les volets Orthographe et Grammaire.](../images/dictionary-agave-03.jpg)


*Figure 4. Définitions dans le volet Synonymes*

![Définitions dans le volet Dictionnaire des synonymes.](../images/dictionary-agave-04.jpg)


*Figure 5. Définitions dans le mode Lecture*

![Définitions en mode lecture.](../images/dictionary-agave-05.jpg)

Pour créer un complément du volet Office qui fournit une recherche de dictionnaire, vous créez deux composants principaux : 

- Un service web XML qui recherche des définitions dans un service de dictionnaire, puis renvoie ces valeurs dans un format XML pouvant être consommé et affiché par le complément de dictionnaire.
- Un complément du volet Office qui soumet la sélection actuelle de l’utilisateur au service web de dictionnaire, affiche des définitions, puis insère facultativement ces valeurs dans le document.

Les sections suivantes fournissent des exemples sur la création de ces composants.

## <a name="creating-a-dictionary-xml-web-service"></a>Création d’un service web XML de dictionnaire

Le service web XML doit renvoyer des requêtes au service web sous la forme de code XML conforme au schéma XML OfficeDefinitions. Les deux sections suivantes décrivent le schéma XML OfficeDefinitions, et fournissent un exemple illustrant comment coder un service web XML qui renvoie des requêtes dans ce format XML.

### <a name="officedefinitions-xml-schema"></a>Schéma XML OfficeDefinitions

Le code suivant illustre le XSD pour le schéma XML OfficeDefinitions.

```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="https://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

Le code XML retourné qui est conforme au schéma OfficeDefinitions se compose d’un élément racine `Result` qui contient un `Definitions` élément avec de zéro à trois `Definition` éléments enfants, chacun contenant des définitions dont la longueur ne dépasse pas 400 caractères. En outre, l’URL de la page complète sur le site de dictionnaire doit être fournie dans l’élément `SeeMoreURL` . L’exemple suivant illustre la structure du code XML renvoyé conforme au schéma OfficeDefinitions.

```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <SeeMoreURL xmlns="">www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```

### <a name="sample-dictionary-xml-web-service"></a>Exemple de service web XML de dictionnaire

Le code C# suivant fournit un exemple simple d’écriture de code pour un service web XML qui renvoie le résultat d’une interrogation de dictionnaire dans le format XML OfficeDefinitions.

```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default.
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components.
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

                    // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();

                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
}
```

## <a name="creating-the-components-of-a-dictionary-add-in"></a>Création des composants d’un complément de dictionnaire

Un complément de dictionnaire est composé de trois fichiers de composants principaux :

- Un fichier de manifeste XML qui décrit le complément.
- Un fichier HTML qui fournit l’interface utilisateur du complément.
- Un fichier JavaScript qui fournit la logique pour obtenir la sélection de l’utilisateur dans le document, envoie la sélection sous forme de requête au service web, puis affiche les résultats renvoyés dans l’interface utilisateur du complément.

### <a name="creating-a-dictionary-add-ins-manifest-file"></a>Création du fichier de manifeste d’un complément de dictionnaire

L’exemple suivant illustre un fichier de manifeste pour un complément de dictionnaire.

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--Capabilities specifies the kind of Office application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. Do not put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

L’élément `Dictionary` et ses éléments enfants qui sont spécifiques à la création du fichier manifeste d’un complément de dictionnaire sont décrits dans les sections suivantes. Pour plus d’informations sur les autres éléments du fichier de manifeste, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

### <a name="dictionary-element"></a>Élément Dictionary

Spécifie les paramètres pour les compléments de dictionnaire.

**Élément parent**

`<OfficeApp>`

**Éléments enfants**

`<TargetDialects>`, `<QueryUri>`, `<CitationText>`, `<DictionaryName>`, `<DictionaryHomePage>`

**Remarques**

L’élément `Dictionary` et ses éléments enfants sont ajoutés au manifeste d’un complément du volet Office lorsque vous créez un complément de dictionnaire.

#### <a name="targetdialects-element"></a>Élément TargetDialects

Indique les langues régionales prises en charge par ce dictionnaire. Requis pour les compléments de dictionnaire.

**Élément parent**

`<Dictionary>`

**Élément enfant**

`<TargetDialect>`

**Remarques**

L’élément `TargetDialects` et ses éléments enfants spécifient l’ensemble des langues régionales que contient votre dictionnaire. Par exemple, si votre dictionnaire s’applique à l’espagnol (Mexique) et à l’espagnol (Pérou), mais pas à l’espagnol (Espagne), vous pouvez le préciser dans cet élément. N’indiquez pas plus d’une langue (par exemple, espagnol et anglais) dans ce manifeste. Publiez les langues distinctes dans des dictionnaires différents.

**Exemple**

```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```

#### <a name="targetdialect-element"></a>Élément TargetDialect

Spécifie une langue régionale prise en charge par ce dictionnaire. Requis pour les compléments de dictionnaire.

**Élément parent**

`<TargetDialects>`

**Remarques**

Spécifie la valeur pour une langue régionale dans le format de balise RFC1766`language`, comme EN-US.

**Exemple**

```XML
<TargetDialect>EN-US</TargetDialect>
```

#### <a name="queryuri-element"></a>Élément QueryUri

Spécifie le point d’extrémité pour le service de requête de dictionnaire. Requis pour les compléments de dictionnaire.

**Élément parent**

`<Dictionary>`

**Remarques**

C’est l’URI du service web XML pour le fournisseur de dictionnaire. La requête correctement formulée sera ajoutée à cette URI. 

**Exemple**

```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```

#### <a name="citationtext-element"></a>Élément CitationText

Spécifie le texte à utiliser dans les citations. Requis pour les compléments de dictionnaire.

**Élément parent**

`<Dictionary>`

**Remarques**

Cet élément spécifie le début du texte de citation qui sera affiché sur une ligne sous le contenu qui est renvoyé du service web (par exemple, « Résultats par : » ou « Optimisé par : »).

Pour cet élément, vous pouvez spécifier des valeurs pour des paramètres régionaux supplémentaires à l’aide de l’élément `Override` . Par exemple si un utilisateur exécute le SKU espagnol d’Office, mais utilise un dictionnaire anglais, ceci permet à la ligne de citation de prendre la valeur « Resultados por: Bing » et non « Results by: Bing ». Pour plus d’informations sur la spécification de valeurs pour des paramètres régionaux supplémentaires, voir la section « Fourniture de paramètres pour différents paramètres régionaux » dans [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

**Exemple**

```XML
<CitationText DefaultValue="Results by: " />
```

#### <a name="dictionaryname-element"></a>Élément DictionaryName

Spécifie le nom de ce dictionnaire. Requis pour les compléments de dictionnaire.

**Élément parent**

`<Dictionary>`

**Remarques**

Cet élément spécifie le texte du lien dans le texte de citation. Le texte de citation s’affiche sur une ligne sous le contenu qui est renvoyé du service web.

Pour cet élément, vous pouvez spécifier des valeurs pour des paramètres régionaux supplémentaires.

**Exemple**

```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```

#### <a name="dictionaryhomepage-element"></a>Élément DictionaryHomePage

Spécifie l’URL de la page d’accueil pour le dictionnaire. Requis pour les compléments de dictionnaire.

**Élément parent**

`<Dictionary>`

**Remarques**

Cet élément spécifie l’URL du lien dans le texte de citation. Le texte de citation s’affiche sur une ligne sous le contenu qui est renvoyé du service web.

Pour cet élément, vous pouvez spécifier des valeurs pour des paramètres régionaux supplémentaires.

**Exemple**

```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```

### <a name="creating-a-dictionary-add-ins-html-user-interface"></a>Création de l’interface utilisateur HTML du complément de dictionnaire

Les deux exemples suivants montrent les fichiers HTML et CSS pour l’interface utilisateur du complément de dictionnaire de démonstration. Pour découvrir comment l’interface utilisateur s’affiche dans le volet Office du complément, voir la figure 6 à la suite du code. Pour voir comment l’implémentation du JavaScript dans le fichier Dictionary.js fournit la logique de programmation de cette interface utilisateur HTML, voir « Écriture de l’implémentation JavaScript » immédiatement à la suite de cette section.

```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

L’exemple suivant montre le contenu de Style.css.

```CSS
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```

*Figure 6. Interface utilisateur du dictionnaire de démonstration*

![Interface utilisateur du dictionnaire de démonstration.](../images/dictionary-agave-06.jpg)

### <a name="writing-the-javascript-implementation"></a>Écriture de l’implémentation JavaScript

L’exemple suivant montre l’implémentation JavaScript dans le fichier Dictionary.js qui est appelé dans la page HTML du complément pour fournir la logique de programmation du complément de dictionnaire de démonstration. Ce script réutilise le service web XML décrit précédemment. Lorsqu’il est placé dans le même répertoire que l’exemple de service web, le script obtient des définitions de ce service. Il peut être utilisé avec un service web XML conforme au schéma OfficeDefinitions public en modifiant la variable  `xmlServiceURL` en haut du fichier, et en remplaçant ensuite la clé de l’API Bing pour obtenir des prononciations adéquates.

Les membres principaux de l’API JavaScript Office (Office.js) qui sont appelés à partir de cette implémentation sont les suivants :

- Événement [d’initialisation](/javascript/api/office) de l’objet `Office` , qui est déclenché lorsque le contexte de complément est initialisé et fournit l’accès à une instance d’objet [Document](/javascript/api/office/office.document) qui représente le document avec lequel le complément interagit.
- Méthode [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) de l’objet `Document` , appelée dans la `initialize` fonction pour ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) du document afin d’écouter les modifications de sélection de l’utilisateur.
- Méthode [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) de l’objet `Document` , appelée dans la `tryUpdatingSelectedWord()` fonction lorsque le `SelectionChanged` gestionnaire d’événements est déclenché pour obtenir le mot ou l’expression sélectionné par l’utilisateur, le forcer à entrer du texte brut, puis exécuter la `selectedTextCallback` fonction de rappel asynchrone.
- Lorsque la  `selectTextCallback` fonction de rappel asynchrone passée en tant qu’argument de _rappel_ de la `getSelectedDataAsync` méthode s’exécute, elle obtient la valeur du texte sélectionné lorsque le rappel est retourné. Elle obtient cette valeur à partir de l’argument _selectedText_ du rappel (qui est de type [AsyncResult) à](/javascript/api/office/office.asyncresult) l’aide de la propriété [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) de l’objet retourné `AsyncResult` .
- Le reste du code dans la fonction  `selectedTextCallback` interroge le service web XML pour obtenir des définitions. Il appelle également les API de Microsoft Translator pour fournir l’URL d’un fichier .wav produisant la prononciation du mot sélectionné.
- Le reste du code dans Dictionary.js affiche la liste de définitions et le lien de prononciation dans l’interface utilisateur HTML du complément.

```js
// The document the dictionary add-in is interacting with.
var _doc;
// The last looked-up word, which is also the currently displayed word.
var lastLookup;
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
var appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8";

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
var xmlServiceUrl = "WebService.asmx/Define?Word=";

// Initialize the add-in.
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document;
    // Check whether text is already selected.
    tryUpdatingSelectedWord();
    // Add a handler to refresh when the user changes selection.
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback method.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    // Change the "See More" link to direct to the correct URL.
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text());
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}
```
