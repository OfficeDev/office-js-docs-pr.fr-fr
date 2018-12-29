---
title: Conservation de l’état et des paramètres des compléments
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: ce2b9ffce97e6338d62cdf07d722ffa384283d28
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458068"
---
# <a name="persisting-add-in-state-and-settings"></a>Conservation de l’état et des paramètres des compléments

Les compléments Office sont essentiellement des applications web exécutées dans l’environnement sans état d’un contrôle de navigateur. En conséquence, votre complément devra peut-être faire persister les données pour assurer la continuité de certaines opérations ou fonctionnalités entre les sessions d’utilisation du complément. Par exemple, votre complément peut disposer de paramètres personnalisés ou d’autres valeurs dont il a besoin pour l’enregistrement et le rechargement à la prochaine initialisation, tels que l’affichage préféré d’un utilisateur ou l’emplacement par défaut. Pour ce faire, vous pouvez procéder comme suit :

- Utilisez les membres de l’API JavaScript pour Office qui stockent les données sous l’une des formes suivantes :
    -  Paires nom/valeur dans un conteneur de propriétés stocké dans un emplacement qui dépend du type de complément.
    -  Éléments XML personnalisés stockés dans le document.
    
- Utilisez des techniques fournies par le contrôle de navigateur sous-jacent : les cookies de navigateur ou le stockage web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    
Cet article se concentre sur l’utilisation de l’interface API JavaScript pour Office afin de faire persister l’état du complément. Pour obtenir des exemples d’utilisation des cookies de navigateur et du stockage web, voir l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistance de l’état et des paramètres d’un complément avec l’interface API JavaScript pour Office

L’interface API JavaScript pour Office fournit les objets [Settings](https://docs.microsoft.com/javascript/api/office/office.settings), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings) et [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties) pour enregistrer l’état du complément dans plusieurs sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id) du complément qui les a créées.

|**Objet**|**Type de complément**|**Emplacement de stockage**|**ôte Office**|
|:-----|:-----|:-----|:-----|
|[Paramètres](https://docs.microsoft.com/javascript/api/office/office.settings)|Contenu et volet de tâches|Document, feuille de calcul ou présentation le complément collabore avec lequel le complément fonctionne. Les paramètres de complément de contenu et de volet Office sont disponibles pour le complément qui les a créés dans le document dans lequel ils sont enregistrés.<br/><br/>**Remarque importante :** ne stockez pas de mots de passe ou autres informations d’identification personnelle (PII) avec l’objet **Settings**. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation de PII de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.|Word, Excel ou PowerPoint<br/><br/> **Remarque :** les compléments du volet Office pour Project 2013 ne prennent pas en charge l’API **Settings** pour le stockage de l’état ou des paramètres du complément. Cependant, pour les compléments exécutés dans Project (et d’autres applications hôtes Office), vous pouvez utiliser des techniques telles que les cookies de navigateur ou le stockage web. Pour plus d’informations sur ces techniques, reportez-vous à l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings)|Outlook|Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé. Comme ces paramètres sont stockés dans la boîte aux lettres de serveur de l’utilisateur, ils sont itinérants et accessibles par le complément lorsqu’il s’exécute dans le contexte d’une application hôte cliente ou d’un navigateur pris en charge accédant à la boîte aux lettres de cet utilisateur.<br/><br/> Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.|Outlook|
|[CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties)|Outlook|Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.|Outlook|
|[CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts)|volet Office|Document, feuille de calcul ou présentation que le complément utilise. Les paramètres de complément de volet Office sont disponibles pour le complément qui les a créés dans le document dans lequel ils sont enregistrés.<br/><br/>**Important :** ne stockez pas de mot de passe ni d’autres informations d’identification personnelle dans une partie XML personnalisée. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation des informations d’identification personnelle de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.|Word (à l’aide de l’API JavaScript courante pour Office) Excel (à l’aide de l’API JavaScript pour Excel propre à l’hôte|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Données de paramètres gérées en mémoire à l’exécution

> [!NOTE]
> Les deux sections suivantes abordent les paramètres dans le contexte de l’API JavaScript courante pour Office. L’API JavaScript pour Excel propre à un hôte propose également un accès aux paramètres personnalisés. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [SettingCollection pour Excel](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection).

En interne, les données dans le conteneur de propriétés accessibles avec les objets**Settings**, **CustomProperties**et**RoamingSettings**sont stockées en tant qu’objet JSON (JavaScript Object Notation) sérialisé contenant des paires nom/valeur. Le nom (clé) pour chaque valeur doit être une **chaîne**, et la valeur stockée peut être une**chaîne**JavaScript, un**nombre**, une**date**, ou un**objet**, mais pas une **fonction**.

Cet exemple de structure de conteneur des propriétés contient trois valeurs**string** définies nommées `firstName`,  `location` et `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Une fois le sac de propriété paramètres enregistré durant la session de complément précédente, il peut être chargé lorsque le complément est initialisé ou à tout moment après pendant la session en cours du complément. Lors de la session, les paramètres sont gérés dans entièrement en mémoire à l’aide des méthodes**obtenir**, **définir**, et **supprimer**de l’objet qui correspond aux paramètres de type que vous créez (**Paramètres**, **CustomProperties**, ou **RoamingSettings**). 


> [!IMPORTANT]
> Pour conserver les ajouts, les mises à jour ou les suppressions apportées pendant la session actuelle du complément à l’emplacement de stockage, vous devez appeler la méthode**saveAsync**de l’objet correspondant utilisé pour le fonctionnement de ce genre de paramètres. Les méthodes**obtenir**, **définir**, et **supprimer** fonctionnent uniquement sur la copie en mémoire du sac de propriété de paramètres. Si votre complément est fermé sans appels**saveAsync**, les modifications apportées aux paramètres pendant cette session seront perdues. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office


Pour conserver l’état ou les paramètres personnalisés d’un complément de contenu ou du volet Office pour Word, Excel ou PowerPoint, utilisez l’objet [Settings](https://docs.microsoft.com/javascript/api/office/office.settings) et ses méthodes. Le conteneur de propriétés créé à l’aide des méthodes de l’objet **Settings** est accessible uniquement par l’instance du complément de contenu ou du volet Office qui l’a créé, et uniquement à partir du document dans lequel il est enregistré.

L’objet**paramètres**est automatiquement chargé dans le cadre de l’objet de[Document](https://docs.microsoft.com/javascript/api/office/office.document)et est disponible lorsque le volet Office ou un complément de contenu est activé. Après l’objet de**Document**instancié, vous pouvez accéder à l’objet**paramètres**avec la propriété des[paramètres](https://docs.microsoft.com/javascript/api/office/office.document#settings)de l’objet du**Document**. Pendant la durée de vie de cette session, vous pouvez simplement utiliser les méthodes**Settings.get**, **Settings.set**, et **Settings.remove**qui permettent de lire, écrire ou supprimer des paramètres persistantes et l’état de complément de la copie en mémoire du sac de propriété.

Étant donné que les méthodes de définition (set) et de suppression (remove) fonctionnent uniquement par rapport à la copie en mémoire du conteneur des propriétés de paramètres, pour enregistrer de nouveaux paramètres ou des paramètres modifiés dans le document auquel le complément est associé, vous devez appeler la méthode [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings#saveasync-options--callback-).


### <a name="creating-or-updating-a-setting-value"></a>Création ou mise à jour d’une valeur de paramètre

L’exemple de code suivant montre comment utiliser la méthode [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings#set-name--value-) pour créer un paramètre appelé `'themeColor'` avec la valeur `'green'`. Le premier paramètre de la méthode set est le _name_ (ID) respectant la casse du paramètre à définir ou à créer. Le second paramètre est la _value_ du paramètre.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 Le paramètre avec le nom spécifié est créé s’il n’existe pas déjà ou sa valeur est mise à jour s’il existe. Utilisez la méthode **Settings.saveAsync** pour rendre persistants les paramètres (nouveaux ou mis à jour) du document.


### <a name="getting-the-value-of-a-setting"></a>Obtention de la valeur d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings#get-name-) pour obtenir la valeur d’un paramètre nommé « themeColor ». Le seul paramètre de la méthode **get** est le _name_ respectant la casse du paramètre.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 La méthode **get** retourne la valeur qui a été précédemment enregistrée pour le _name_ du paramètre qui a été passé. Si le paramètre n’existe pas, la méthode retourne **null**.


### <a name="removing-a-setting"></a>Suppression d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings#remove-name-) pour supprimer un paramètre portant le nom « themeColor ». Le seul paramètre de la méthode **remove** est le _name_ respectant la casse du paramètre.


```js
Office.context.document.settings.remove('themeColor');
```

Rien ne se produit si le paramètre n’existe pas. Utilisez la méthode**Settings.saveAsync** pour faire persister la suppression du paramètre du document.


### <a name="saving-your-settings"></a>Enregistrement de vos paramètres

Pour enregistrer les ajouts, modifications ou suppressions que votre complément a effectués sur la copie en mémoire du conteneur de propriétés des paramètres pendant la session en cours, vous devez appeler la méthode [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings#saveasync-options--callback-) pour les stocker dans le document. L’unique paramètre de la méthode **saveAsync** est _callback_, lequel est une fonction de rappel avec un paramètre unique. 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

La fonction anonyme qui est passée dans la méthode**saveAsync**comme paramètre_callback_est exécutée lorsque l’opération est terminée. Le paramètre_asyncResult_du rappel permet d’accéder à un objet**AsyncResult**qui contient l’état de l’opération. Dans l’exemple, la fonction vérifie la propriété**AsyncResult.status** pour déterminer si l’enregistrement opération a réussi ou échoué et puis affiche le résultat dans la page du complément.

## <a name="how-to-save-custom-xml-to-the-document"></a>Enregistrement des parties XML personnalisées dans le document

> [!NOTE]
> Cette section décrit les parties XML personnalisées dans le contexte de l’API JavaScript courante pour Office qui est prise en charge dans Word. L’API JavaScript pour Excel propre à un hôte permet également d’accéder aux parties XML personnalisées. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [CustomXmlPart pour Excel](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart).

Une option de stockage supplémentaire est disponible lorsque vous avez besoin de stocker des informations dépassant les limites de taille des paramètres du document ou comportant un caractère structuré. Vous pouvez conserver le balisage XML personnalisé dans un complément de volet Office pour Word (et pour Excel, mais reportez-vous à la remarque en haut de cette section). Dans Word, utilisez l’objet[CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart) et ses méthodes (nouveau, voir la note ci-dessus pour Excel). Le code suivant crée une partie XML personnalisée et affiche son ID et son contenu dans des balises div sur la page. Un attribut`xmlns` doit figurer dans la chaîne XML.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```

Pour récupérer une partie XML personnalisée, vous utilisez la méthode [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mais l’identifiant correspond à un GUID généré lorsque la partie XML est créée. Vous ne pouvez donc pas connaître l’identifiant lors du codage. Pour cette raison, il est recommandé de stocker immédiatement l’identifiant de la partie XML en tant que paramètre et de lui donner une clé facilement mémorisable lorsque vous créez une partie XML. L’exemple de méthode suivant montre comment procéder. (Toutefois, reportez-vous aux sections précédentes de cet article pour obtenir des détails et des meilleures pratiques lorsque vous utilisez des paramètres personnalisés).

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

Le code suivant montre comment récupérer la partie XML en obtenant d’abord son identifiant partir d’un paramètre.

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID'));
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId, 
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Enregistrement des paramètres dans la boîte aux lettres de l’utilisateur pour les compléments Outlook en tant que paramètres d’itinérance


Un complément Outlook peut utiliser l’objet[RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings)pour enregistrer les données de paramètres et d’état qui vous sont propres à la boîte aux lettres de l’utilisateur. Ces données ne sont pas accessibles uniquement par ce complément Outlook au nom de l’utilisateur qui exécute le complément. Les données stockées dans la boîte aux lettres Exchange Server de l’utilisateur et sont accessibles lorsque cet utilisateur se connecte à son compte et exécute le complément Outlook.


### <a name="loading-roaming-settings"></a>Chargement des paramètres d’itinérance


Un complément Outlook charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](https://docs.microsoft.com/javascript/api/office). L’exemple de code JavaScript suivant explique comment charger des paramètres d’itinérance existants.


```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>Création ou affectation d’un paramètre d’itinérance


Pour faire suite à l’exemple précédent, la fonction  `setAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#set-name--value-) pour définir ou mettre à jour un paramètre nommé `cookie` avec la date du jour. Elle réenregistre ensuite tous les paramètres d’itinérance sur le serveur Exchange avec la méthode [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#saveasync-callback-).


```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

La méthode**saveAsync**enregistre les paramètres d’itinérance façon asynchrone par rapport et prend une fonction de rappel facultatif. Cet exemple de code transmet une fonction de rappel nommée `saveMyAppSettingsCallback` à la méthode **saveAsync**. Lorsque l’appel asynchrone renvoie, le paramètre_asyncResult_ de la fonction`saveMyAppSettingsCallback`permet d’accéder à un objet[AsyncResult](https://docs.microsoft.com/javascript/api/outlook)que vous pouvez utiliser pour déterminer la réussite ou l’échec de la opération avec la propriété**AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Suppression d’un paramètre d’itinérance


Toujours dans le prolongement des exemples précédents, la fonction  `removeAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#remove-name-) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Enregistrement des paramètres par élément pour les compléments Outlook en tant que propriétés personnalisées


Les propriétés personnalisées permettent à votre complément Outlook de stocker des informations sur un élément qu’il utilise. Par exemple, si votre complément Outlook crée un rendez-vous à partir d’une suggestion de réunion dans un message, vous pouvez utiliser des propriétés personnalisées pour stocker le fait que la réunion a été créée. Cela garantit que si le message est rouvert, votre complément Outlook ne propose pas de recréer le rendez-vous.

Pour pouvoir utiliser des propriétés personnalisées pour un élément de message, de rendez-vous ou de demande de réunion particulier, vous devez charger les propriétés en mémoire en appelant la méthode [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) de l’objet **Item**. Si des propriétés personnalisées sont déjà définies pour l’élément actuel, elles sont chargées à ce moment à partir du serveur Exchange. Après avoir chargé les propriétés, vous pouvez utiliser les méthodes [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties#set-name--value-) et [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings) de l’objet **CustomProperties** pour ajouter, mettre à jour et récupérer des propriétés en mémoire. Pour enregistrer les modifications que vous avez apportées aux propriétés personnalisées de l’élément, vous devez utiliser la méthode [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) pour conserver les modifications de l’élément sur le serveur Exchange.


### <a name="custom-properties-example"></a>Exemple de propriétés personnalisées

L’exemple suivant illustre un ensemble simplifié des fonctions pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément Outlook qui utilise des propriétés personnalisées. 

Un complément Outlook qui utilise ces fonctions récupère des propriétés personnalisées en appelant la méthode**get** sur la variable`_customProps`, comme le montre l’exemple suivant.




```js
var property = _customProps.get("propertyName");
```

Cet exemple inclut les fonctions suivantes:



|**Nom de la fonction**|**Description**|
|:-----|:-----|
| `Office.initialize`|Initialise le complément et charge les propriétés personnalisées pour l’élément actuel à partir du serveur Exchange.|
| `customPropsCallback`|Obtient les propriétés personnalisées retournées du serveur Exchange et les enregistre pour une utilisation ultérieure.|
| `updateProperty`|Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur Exchange.|
| `removeProperty`|Supprime une propriété spécifique, puis fait persister la suppression sur le serveur Exchange.|
| `saveCallback`|Rappel pour les appels à la méthode**saveAsync** dans les fonctions`updateProperty` et `removeProperty`.|



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
