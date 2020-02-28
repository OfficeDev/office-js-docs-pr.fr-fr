---
title: Extraire des chaînes d’entités d’un élément Outlook
description: Découvrez comment extraire des chaînes d’entités d’un élément Outlook dans un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 0a9a41d0b479420c0754c0e0d283982082a1452f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325453"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a><span data-ttu-id="8db06-103">Extraire des chaînes d’entité d’un élément Outlook</span><span class="sxs-lookup"><span data-stu-id="8db06-103">Extract entity strings from an Outlook item</span></span>

<span data-ttu-id="8db06-p101">Cet article décrit comment créer un complément Outlook pour l’**affichage des entités** qui extrait des instances de chaînes d’entités connues prises en charge dans l’objet et le corps de l’élément Outlook sélectionné. Cet élément peut être un rendez-vous, un message électronique ou encore une demande, une réponse ou une annulation de réunion.</span><span class="sxs-lookup"><span data-stu-id="8db06-p101">This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.</span></span>

<span data-ttu-id="8db06-106">Les entités prises en charge incluent notamment :</span><span class="sxs-lookup"><span data-stu-id="8db06-106">The supported entities include:</span></span>

- <span data-ttu-id="8db06-107">**Adresse** : une adresse postale aux États-Unis, qui a au moins un sous-ensemble des éléments de numéro de rue, nom de rue, ville, état et code postal.</span><span class="sxs-lookup"><span data-stu-id="8db06-107">**Address**: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.</span></span>
    
- <span data-ttu-id="8db06-108">**Contact** : informations de contact d’une personne, dans le contexte d’autres entités, telles qu’une adresse ou un nom commercial.</span><span class="sxs-lookup"><span data-stu-id="8db06-108">**Contact**: A person's contact information, in the context of other entities such as an address or business name.</span></span>
    
- <span data-ttu-id="8db06-109">**Adresse électronique** : une adresse électronique SMTP.</span><span class="sxs-lookup"><span data-stu-id="8db06-109">**Email address**: An SMTP email address.</span></span>
    
- <span data-ttu-id="8db06-p102">**Suggestion de réunion** : une suggestion de réunion, par exemple une référence à un événement. Notez que seuls les messages (pas les rendez-vous) prennent en charge l’extraction de suggestion de réunion.</span><span class="sxs-lookup"><span data-stu-id="8db06-p102">**Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.</span></span>
    
- <span data-ttu-id="8db06-112">**Numéro de téléphone** : numéro de téléphone nord-américain.</span><span class="sxs-lookup"><span data-stu-id="8db06-112">**Phone number**: A North American phone number.</span></span>
    
- <span data-ttu-id="8db06-113">**Suggestion de tâche** : une suggestion de tâche, généralement exprimée dans une phrase associée à une action.</span><span class="sxs-lookup"><span data-stu-id="8db06-113">**Task suggestion**: A task suggestion, typically expressed in an actionable phrase.</span></span>
    
- <span data-ttu-id="8db06-114">**URL**</span><span class="sxs-lookup"><span data-stu-id="8db06-114">**URL**</span></span>
    
<span data-ttu-id="8db06-p103">La plupart de ces entités s’appuient sur la reconnaissance vocale en langage naturel qui est basée sur l’apprentissage machine de grandes quantités de données. Par conséquent, la reconnaissance est non déterministe et dépend parfois du contexte de l’élément Outlook.</span><span class="sxs-lookup"><span data-stu-id="8db06-p103">Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.</span></span>

<span data-ttu-id="8db06-p104">Outlook active le complément des entités lorsque l’utilisateur sélectionne un rendez-vous, un message électronique ou une demande, réponse ou annulation de réunion à afficher. Lors de l’initialisation, l’exemple de complément d’entités lit toutes les instances des entités prises en charge à partir de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="8db06-p104">Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.</span></span> 

<span data-ttu-id="8db06-p105">Le complément comporte des boutons permettant à l’utilisateur de choisir un type d’entité. Quand l’utilisateur sélectionne une entité, le complément affiche les instances de l’entité sélectionnée dans le volet de complément. Les sections suivantes répertorient les manifestes XML, les fichiers HTML et JavaScript du complément des entités, et met en évidence le code qui prend en charge l’extraction de l’entité associée.</span><span class="sxs-lookup"><span data-stu-id="8db06-p105">The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.</span></span>

## <a name="xml-manifest"></a><span data-ttu-id="8db06-122">Manifeste XML</span><span class="sxs-lookup"><span data-stu-id="8db06-122">XML manifest</span></span>

<span data-ttu-id="8db06-123">Le complément pour entités a deux règles d’activation jointes par une opération OR logique.</span><span class="sxs-lookup"><span data-stu-id="8db06-123">The entities add-in has two activation rules joined by a logical OR operation.</span></span> 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

<span data-ttu-id="8db06-124">Ces règles spécifient qu’Outlook doit activer ce complément lorsque l’élément sélectionné dans l’inspecteur ou le volet de lecture est un rendez-vous ou un message (notamment un message électronique, ou une demande, une réponse ou une annulation de réunion).</span><span class="sxs-lookup"><span data-stu-id="8db06-124">These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).</span></span>

<span data-ttu-id="8db06-p106">Voici le manifeste du complément pour entités. Il utilise la version 1.1 du schéma pour les manifestes des Compléments Office.</span><span class="sxs-lookup"><span data-stu-id="8db06-p106">The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## <a name="html-implementation"></a><span data-ttu-id="8db06-127">Implémentation HTML</span><span class="sxs-lookup"><span data-stu-id="8db06-127">HTML implementation</span></span>

<span data-ttu-id="8db06-p107">Le fichier HTML du complément pour entités spécifie les boutons permettant à l’utilisateur de sélectionner chaque type d’entité, et un autre bouton pour effacer les instances affichées d’une entité. Il inclut un fichier JavaScript, default_entities.js, qui est décrit dans la section suivante sous [Implémentation JavaScript](#javascript-implementation). Le fichier JavaScript inclut le gestionnaire d’événements pour chacun des boutons.</span><span class="sxs-lookup"><span data-stu-id="8db06-p107">The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.</span></span>

<span data-ttu-id="8db06-p108">Notez que tous les compléments Outlook doivent comprendre le fichier office.js. Le fichier HTML suivant inclut la version 1.1 du fichier office.js sur le réseau de distribution de contenu (CDN).</span><span class="sxs-lookup"><span data-stu-id="8db06-p108">Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN.</span></span> 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## <a name="style-sheet"></a><span data-ttu-id="8db06-133">Feuille de style</span><span class="sxs-lookup"><span data-stu-id="8db06-133">Style sheet</span></span>


<span data-ttu-id="8db06-p109">Le complément pour entités utilise un fichier CSS facultatif, default_entities.css, pour spécifier la mise en forme de la sortie. Le fichier CSS est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p109">The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.</span></span>


```CSS
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## <a name="javascript-implementation"></a><span data-ttu-id="8db06-136">Implémentation JavaScript</span><span class="sxs-lookup"><span data-stu-id="8db06-136">JavaScript implementation</span></span>

<span data-ttu-id="8db06-137">Les sections suivantes expliquent comment l’exemple suivant (le fichier default_entities.js) extrait des entités connues de l’objet et du corps du message ou du rendez-vous consulté par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8db06-137">The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing.</span></span>

## <a name="extracting-entities-upon-initialization"></a><span data-ttu-id="8db06-138">Extraction d’entités lors de l’initialisation</span><span class="sxs-lookup"><span data-stu-id="8db06-138">Extracting entities upon initialization</span></span>

<span data-ttu-id="8db06-139">Lors de l’événement [Office.initialize](/javascript/api/office#office-initialize-reason-), le complément pour entités appelle la méthode [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8db06-139">Upon the [Office.initialize](/javascript/api/office#office-initialize-reason-) event, the entities add-in calls the [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method of the current item.</span></span> <span data-ttu-id="8db06-140">La `getEntities` méthode renvoie la variable `_MyEntities` globale tableau d’instances des entités prises en charge.</span><span class="sxs-lookup"><span data-stu-id="8db06-140">The `getEntities` method returns the global variable `_MyEntities` an array of instances of supported entities.</span></span> <span data-ttu-id="8db06-141">Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-141">The following is the related JavaScript code.</span></span>


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a><span data-ttu-id="8db06-142">Extraction d’adresses</span><span class="sxs-lookup"><span data-stu-id="8db06-142">Extracting addresses</span></span>


<span data-ttu-id="8db06-143">Lorsque l’utilisateur clique sur le bouton **Obtenir les adresses**, le gestionnaire d’événements `myGetAddresses` obtient un tableau d’adresses à partir de la propriété [adresses](/javascript/api/outlook/office.entities#addresses) de l’objet `_MyEntities`, si une adresse a été extraite.</span><span class="sxs-lookup"><span data-stu-id="8db06-143">When the user clicks the **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](/javascript/api/outlook/office.entities#addresses) property of the `_MyEntities` object, if any address was extracted.</span></span> <span data-ttu-id="8db06-144">Toute adresse extraite est stockée comme chaîne dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="8db06-144">Each extracted address is stored as a string in the array.</span></span> <span data-ttu-id="8db06-145">`myGetAddresses` forme une chaîne HTML locale dans `htmlText` pour afficher la liste des adresses extraites.</span><span class="sxs-lookup"><span data-stu-id="8db06-145">`myGetAddresses` forms a local HTML string in `htmlText` to display the list of extracted addresses.</span></span> <span data-ttu-id="8db06-146">Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-146">The following is the related JavaScript code.</span></span>


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a><span data-ttu-id="8db06-147">Extraction d’informations de contact</span><span class="sxs-lookup"><span data-stu-id="8db06-147">Extracting contact information</span></span>


<span data-ttu-id="8db06-p112">Lorsque l’utilisateur clique sur le bouton **Obtenir des informations de contact**, le gestionnaire d’événements `myGetContacts` obtient un tableau de contacts avec leurs informations à partir de la propriété [contacts](/javascript/api/outlook/office.entities#contacts) de l’objet `_MyEntities`, si des contacts ont été extraits. Chaque contact extrait est stocké sous la forme d’un objet [Contact](/javascript/api/outlook/office.contact) dans le tableau. `myGetContacts` obtient d’autres données sur le contact. Notez que le contexte détermine si Outlook peut extraire un contact à partir d’un élément &mdash; Il doit exister une signature à la fin d’un message électronique ou au moins une partie des informations suivantes à proximité du contact :</span><span class="sxs-lookup"><span data-stu-id="8db06-p112">When the user clicks the **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](/javascript/api/outlook/office.entities#contacts) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](/javascript/api/outlook/office.contact) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item&mdash;a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:</span></span>


- <span data-ttu-id="8db06-152">La chaîne représentant le nom du contact à partir de la propriété [Contact.personName](/javascript/api/outlook/office.contact#personname).</span><span class="sxs-lookup"><span data-stu-id="8db06-152">The string representing the contact's name from the [Contact.personName](/javascript/api/outlook/office.contact#personname) property.</span></span>

- <span data-ttu-id="8db06-153">La chaîne représentant le nom de l’entreprise associée au contact à partir de la propriété [Contact.businessName](/javascript/api/outlook/office.contact#businessname).</span><span class="sxs-lookup"><span data-stu-id="8db06-153">The string representing the company name associated with the contact from the [Contact.businessName](/javascript/api/outlook/office.contact#businessname) property.</span></span>

- <span data-ttu-id="8db06-p113">Le tableau des numéros de téléphone associés au contact à partir de la propriété [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers). Chaque numéro de téléphone est représenté par un objet [PhoneNumber](/javascript/api/outlook/office.phonenumber).</span><span class="sxs-lookup"><span data-stu-id="8db06-p113">The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.</span></span>

- <span data-ttu-id="8db06-156">Pour chaque membre **PhoneNumber** dans le tableau des numéros de téléphone, la chaîne représentant le numéro de téléphone à partir de la propriété [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring).</span><span class="sxs-lookup"><span data-stu-id="8db06-156">For each **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="8db06-p114">Le tableau des URL associées au contact à partir de la propriété [Contact.urls](/javascript/api/outlook/office.contact#urls). Chaque URL est représentée sous la forme d’une chaîne dans un membre de tableau.</span><span class="sxs-lookup"><span data-stu-id="8db06-p114">The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#urls) property. Each URL is represented as a string in an array member.</span></span>

- <span data-ttu-id="8db06-p115">Le tableau des adresses électroniques associées au contact à partir de la propriété [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses). Chaque adresse électronique est représentée sous la forme d’une chaîne dans un membre de tableau.</span><span class="sxs-lookup"><span data-stu-id="8db06-p115">The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) property. Each email address is represented as a string in an array member.</span></span>

- <span data-ttu-id="8db06-p116">Le tableau des adresses postales associées au contact à partir de la propriété [Contact.addresses](/javascript/api/outlook/office.contact#addresses). Chaque adresse postale est représentée sous la forme d’une chaîne dans un membre de tableau.</span><span class="sxs-lookup"><span data-stu-id="8db06-p116">The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#addresses) property. Each postal address is represented as a string in an array member.</span></span>

<span data-ttu-id="8db06-p117">`myGetContacts` forme une chaîne HTML locale dans `htmlText` pour afficher les données pour chaque contact. Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p117">`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a><span data-ttu-id="8db06-165">Extraction des adresses électroniques</span><span class="sxs-lookup"><span data-stu-id="8db06-165">Extracting email addresses</span></span>


<span data-ttu-id="8db06-p118">Lorsque l’utilisateur clique sur le bouton **Obtenir des adresses électroniques**, le gestionnaire d’événements `myGetEmailAddresses` obtient un tableau d’adresses électroniques SMTP à partir de la propriété [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) de l’objet `_MyEntities`, si des adresses ont été extraites. Chaque adresse électronique extraite est stockée sous la forme d’une chaîne dans le tableau. `myGetEmailAddresses` forme une chaîne HTML locale dans `htmlText` pour afficher la liste des adresses électroniques extraites. Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p118">When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.</span></span>


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a><span data-ttu-id="8db06-170">Extraction de suggestions de réunion</span><span class="sxs-lookup"><span data-stu-id="8db06-170">Extracting meeting suggestions</span></span>


<span data-ttu-id="8db06-171">Lorsque l’utilisateur clique sur le bouton **Obtenir des suggestions de réunion**, le gestionnaire d’événements `myGetMeetingSuggestions` obtient un tableau de suggestions de réunion à partir de la propriété [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) de l’objet `_MyEntities`, si des suggestions ont été extraites.</span><span class="sxs-lookup"><span data-stu-id="8db06-171">When the user clicks the **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) property of the `_MyEntities` object, if any was extracted.</span></span>


 > [!NOTE]
 > <span data-ttu-id="8db06-172">Seuls les messages mais pas les rendez-vous prennent en charge le `MeetingSuggestion` type d’entité.</span><span class="sxs-lookup"><span data-stu-id="8db06-172">Only messages but not appointments support the `MeetingSuggestion` entity type.</span></span>

<span data-ttu-id="8db06-p119">Chaque suggestion de réunion extraite est stockée sous la forme d’un objet [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) dans le tableau. `myGetMeetingSuggestions` obtient d’autres données sur chaque suggestion de réunion :</span><span class="sxs-lookup"><span data-stu-id="8db06-p119">Each extracted meeting suggestion is stored as a [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:</span></span>


- <span data-ttu-id="8db06-175">La chaîne identifiée comme suggestion de réunion à partir de la propriété [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring).</span><span class="sxs-lookup"><span data-stu-id="8db06-175">The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) property.</span></span>

- <span data-ttu-id="8db06-p120">Le tableau des participants de la réunion à partir de la propriété [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees). Chaque participant est représenté par un objet [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="8db06-p120">The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="8db06-178">Le nom de chaque participant à partir de la propriété [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname).</span><span class="sxs-lookup"><span data-stu-id="8db06-178">For each attendee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="8db06-179">L’adresse SMTP de chaque participant à partir de la propriété [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress).</span><span class="sxs-lookup"><span data-stu-id="8db06-179">For each attendee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

- <span data-ttu-id="8db06-180">La chaîne représentant l’emplacement de la suggestion de réunion à partir de la propriété [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location).</span><span class="sxs-lookup"><span data-stu-id="8db06-180">The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) property.</span></span>

- <span data-ttu-id="8db06-181">La chaîne représentant l’objet de la suggestion de réunion à partir de la propriété [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject).</span><span class="sxs-lookup"><span data-stu-id="8db06-181">The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) property.</span></span>

- <span data-ttu-id="8db06-182">La chaîne représentant l’heure de début de la suggestion de réunion à partir de la propriété [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start).</span><span class="sxs-lookup"><span data-stu-id="8db06-182">The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) property.</span></span>

- <span data-ttu-id="8db06-183">La chaîne représentant l’heure de fin de la suggestion de réunion à partir de la propriété [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end).</span><span class="sxs-lookup"><span data-stu-id="8db06-183">The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) property.</span></span>

<span data-ttu-id="8db06-p121">`myGetMeetingSuggestions` forme une chaîne HTML locale dans `htmlText` pour afficher les données pour chacune des suggestions de réunion. Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p121">`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-phone-numbers"></a><span data-ttu-id="8db06-186">Extraction de numéros de téléphone</span><span class="sxs-lookup"><span data-stu-id="8db06-186">Extracting phone numbers</span></span>


<span data-ttu-id="8db06-p122">Lorsque l’utilisateur clique sur le bouton **Obtenir des numéros de téléphone**, le gestionnaire d’événements `myGetPhoneNumbers` obtient un tableau de numéros de téléphone à partir de la propriété [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) de l’objet `_MyEntities`, si des numéros de téléphone ont été extraits. Chaque numéro de téléphone extrait est stocké sous la forme d’un objet [PhoneNumber](/javascript/api/outlook/office.phonenumber) dans le tableau. `myGetPhoneNumbers` obtient d’autres données sur chaque numéro de téléphone :</span><span class="sxs-lookup"><span data-stu-id="8db06-p122">When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:</span></span>


- <span data-ttu-id="8db06-190">La chaîne représentant le type de numéro de téléphone (par exemple, numéro de téléphone du domicile) à partir de la propriété [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type).</span><span class="sxs-lookup"><span data-stu-id="8db06-190">The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) property.</span></span>

- <span data-ttu-id="8db06-191">La chaîne représentant le numéro de téléphone réel à partir de la propriété [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring).</span><span class="sxs-lookup"><span data-stu-id="8db06-191">The string representing the actual phone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="8db06-192">La chaîne qui a été initialement identifiée comme le numéro de téléphone à partir de la propriété [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring).</span><span class="sxs-lookup"><span data-stu-id="8db06-192">The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) property.</span></span>

<span data-ttu-id="8db06-p123">`myGetPhoneNumbers` forme une chaîne HTML locale dans `htmlText` pour afficher les données pour chacun des numéros de téléphone. Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p123">`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-task-suggestions"></a><span data-ttu-id="8db06-195">Extraction de suggestions de tâches</span><span class="sxs-lookup"><span data-stu-id="8db06-195">Extracting task suggestions</span></span>


<span data-ttu-id="8db06-p124">Lorsque l’utilisateur clique sur le bouton **Obtenir des suggestions de tâches**, le gestionnaire d’événements `myGetTaskSuggestions` obtient un tableau de suggestions de tâches à partir de la propriété [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) de l’objet `_MyEntities`, si des suggestions ont été extraites. Chaque suggestion de tâche extraite est stockée sous la forme d’un objet [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) dans le tableau. `myGetTaskSuggestions` obtient d’autres données sur chaque suggestion de tâche :</span><span class="sxs-lookup"><span data-stu-id="8db06-p124">When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:</span></span>


- <span data-ttu-id="8db06-199">La chaîne qui a été initialement identifiée comme une suggestion de tâche à partir de la propriété [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring).</span><span class="sxs-lookup"><span data-stu-id="8db06-199">The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) property.</span></span>

- <span data-ttu-id="8db06-p125">Le tableau des cessionnaires de tâches à partir de la propriété [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees). Chaque cessionnaire est représenté par un objet [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="8db06-p125">The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="8db06-202">Le nom de chaque cessionnaire à partir de la propriété [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname).</span><span class="sxs-lookup"><span data-stu-id="8db06-202">For each assignee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="8db06-203">L’adresse SMTP de chaque cessionnaire à partir de la propriété [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress).</span><span class="sxs-lookup"><span data-stu-id="8db06-203">For each assignee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

<span data-ttu-id="8db06-p126">`myGetTaskSuggestions` forme une chaîne HTML locale dans `htmlText` pour afficher les données pour chaque suggestion de tâche. Le code JavaScript associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p126">`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-urls"></a><span data-ttu-id="8db06-206">Extraction d’URL</span><span class="sxs-lookup"><span data-stu-id="8db06-206">Extracting URLs</span></span>


<span data-ttu-id="8db06-p127">Lorsque l’utilisateur clique sur le bouton **Obtenir des URL**, le gestionnaire d’événements `myGetUrls` obtient un tableau d’URL à partir de la propriété [urls](/javascript/api/outlook/office.entities#urls) de l’objet `_MyEntities`, si des URL ont été extraites. Chaque URL extraite est stockée sous la forme d’une chaîne dans le tableau. `myGetUrls` forme une chaîne HTML locale dans `htmlText` pour afficher la liste des URL extraites.</span><span class="sxs-lookup"><span data-stu-id="8db06-p127">When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#urls) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.</span></span>


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a><span data-ttu-id="8db06-210">Effacement des chaînes d’entités affichées</span><span class="sxs-lookup"><span data-stu-id="8db06-210">Clearing displayed entity strings</span></span>


<span data-ttu-id="8db06-p128">Enfin, le complément pour entités spécifie un gestionnaire d’événements  `myClearEntitiesBox` qui efface les chaînes affichées. Le code associé est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-p128">Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.</span></span>


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a><span data-ttu-id="8db06-213">Listing JavaScript</span><span class="sxs-lookup"><span data-stu-id="8db06-213">JavaScript listing</span></span>


<span data-ttu-id="8db06-214">Le listing complet de l’implémentation JavaScript est présenté ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="8db06-214">The following is the complete listing of the JavaScript implementation.</span></span>


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="see-also"></a><span data-ttu-id="8db06-215">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8db06-215">See also</span></span>

- [<span data-ttu-id="8db06-216">Créer des compléments Outlook pour des formulaires de lecture</span><span class="sxs-lookup"><span data-stu-id="8db06-216">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="8db06-217">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="8db06-217">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- <span data-ttu-id="8db06-218">Méthode [item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="8db06-218">[item.getEntities method](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span></span>
