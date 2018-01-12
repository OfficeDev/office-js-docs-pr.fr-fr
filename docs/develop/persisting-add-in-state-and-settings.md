
# <a name="persisting-add-in-state-and-settings"></a>Conservation de l’état et des paramètres des compléments

Les Compléments Office sont essentiellement des applications web exécutées dans l’environnement sans état d’un contrôle de navigateur. En conséquence, votre complément devra peut-être faire persister les données pour assurer la continuité de certaines opérations ou fonctionnalités entre les sessions d’utilisation du complément. Par exemple, votre complément peut disposer de paramètres personnalisés ou d’autres valeurs dont il a besoin pour l’enregistrement et le rechargement à la prochaine initialisation, tels que l’affichage préféré d’un utilisateur ou l’emplacement par défaut.

Pour ce faire, vous pouvez :


- Utilisez des membres de l’interface API JavaScript pour Office qui stockent les données en tant que paires nom/valeur dans un conteneur des propriétés stocké à un emplacement qui dépend du type de complément.
    
- Utiliser des techniques fournies par le contrôle de navigateur sous-jacent : les cookies de navigateur ou le stockage Web HTML5 ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) ou [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)).
    
Cet article se concentre sur l’utilisation de l’interface API JavaScript pour Office afin de faire persister l’état du complément. Pour obtenir des exemples d’utilisation des cookies de navigateur et du stockage web, voir l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistance de l’état et des paramètres d’un complément avec l’interface API JavaScript pour Office


L’interface API JavaScript pour Office fournit les objets [Settings](../../reference/shared/settings.md), [RoamingSettings](../../reference/outlook/RoamingSettings.md) et [CustomProperties](../../reference/outlook/CustomProperties.md) pour enregistrer l’état du complément dans plusieurs sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) du complément qui les a créées.



|**Objet**|**Type de complément**|**Emplacement de stockage**|**ôte Office**|
|:-----|:-----|:-----|:-----|
|[Settings](../../reference/shared/settings.md)|Contenu et volet de tâches|Document, feuille de calcul ou présentation qu’utilise le complément.Seul le complément qui a créé les paramètres de complément de contenu et du volet Office peut y accéder à partir du document où ils sont enregistrés. **Important :** ne stockez pas de mots de passe ou autres informations d’identification personnelle (PII) avec l’objet **Settings**. Les données enregistrées ne sont pas visibles par les utilisateurs finaux. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation de PII de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.|Word, Excel ou PowerPoint **Remarque :** les compléments du volet Office pour Project 2013 ne prennent pas en charge l’API **Settings** pour le stockage de l’état ou des paramètres du complément. Cependant, pour les compléments exécutés dans Project (et d’autres applications hôtes Office), vous pouvez utiliser des techniques telles que les cookies de navigateur ou le stockage web. Pour plus d’informations sur ces techniques, voir l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé.Comme ces paramètres sont stockés dans la boîte aux lettres de serveur de l’utilisateur, ils sont itinérants et accessibles par le complément lorsqu’il s’exécute dans le contexte d’une application hôte cliente ou d’un navigateur pris en charge accédant à la boîte aux lettres de cet utilisateur. Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.|Outlook|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Données de paramètres gérées en mémoire à l’exécution


En interne, les données dans le conteneur de propriétés accessibles avec les objets  **Settings**,  **CustomProperties** et **RoamingSettings** sont stockées en tant qu’objet JSON (JavaScript Object Notation) sérialisé contenant des paires nom/valeur. Le nom (clé) de chaque valeur doit être une **string** et la valeur stockée peut être un élément JavaScript **string**,  **number**,  **date** ou **object**, mais pas  **function**.

Cet exemple de structure de conteneur des propriétés contient trois valeurs  **string** définies nommées `firstName`,  `location` et `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

Une fois le conteneur de propriétés des paramètres enregistré lors de la session de complément précédente, il peut être chargé lorsque le complément est initialisé ou à tout moment par la suite pendant la session active du complément. Pendant cette session, les paramètres sont gérés entièrement en mémoire à l’aide des méthodes  **get**,  **set** et **remove** de l’objet qui correspond aux paramètres de type créés ( **Settings**,  **CustomProperties** ou **RoamingSettings**). 


 >**Importante**  Pour rendre persistants les ajouts, mises à jour ou suppressions pendant la session active du complément dans l’emplacement de stockage, vous devez appeler la méthode  **saveAsync** de l’objet correspondant utilisé pour travailler avec ce type de paramètres. Les méthodes **get**,  **set** et **remove** fonctionnent uniquement sur la copie en mémoire du conteneur des propriétés des paramètres. Si votre complément est fermé sans appel à **saveAsync**, les modifications apportées aux paramètres pendant la session sont perdues. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office


Pour conserver l’état ou les paramètres personnalisés d’un complément de contenu ou du volet Office pour Word, Excel ou PowerPoint, utilisez l’objet [Settings](../../reference/shared/settings.md) et ses méthodes. Le conteneur de propriétés créé à l’aide des méthodes de l’objet **Settings** est accessible uniquement par l’instance du complément de contenu ou du volet Office qui l’a créé, et uniquement à partir du document dans lequel il est enregistré.

L’objet  **Settings** est automatiquement chargé comme partie intégrante de l’objet [Document](../../reference/shared/document.md) et il est disponible lorsque le complément du volet Office ou de contenu est activé. Une fois que l’objet **Document** est instancié, vous pouvez accéder à l’objet **Settings** en utilisant la propriété [settings](../../reference/shared/document.settings.md) de l’objet **Document**. Pendant la durée de vie de la session, vous ne pouvez utiliser que les méthodes  **Settings.get**,  **Settings.set** et **Settings.remove** pour lire, écrire et supprimer les paramètres et l’état du complément conservés dans la copie en mémoire du conteneur de propriétés.

Étant donné que les méthodes de définition (set) et de suppression (remove) fonctionnent uniquement par rapport à la copie en mémoire du conteneur des propriétés de paramètres, pour enregistrer de nouveaux paramètres ou des paramètres modifiés dans le document auquel le complément est associé, vous devez appeler la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


### <a name="creating-or-updating-a-setting-value"></a>Création ou mise à jour d’une valeur de paramètre

L’exemple de code suivant montre comment utiliser la méthode [Settings.set](../../reference/shared/settings.set.md) pour créer un paramètre appelé `'themeColor'` avec la valeur `'green'`. Le premier paramètre de la méthode set est le _name_ (ID) respectant la casse du paramètre à définir ou à créer. Le second paramètre est la _value_ du paramètre.


```
Office.context.document.settings.set('themeColor', 'green');
```

 Le paramètre avec le nom spécifié est créé s’il n’existe pas déjà ou sa valeur est mise à jour s’il existe. Utilisez la méthode **Settings.saveAsync** pour rendre persistants les paramètres (nouveaux ou mis à jour) du document.


### <a name="getting-the-value-of-a-setting"></a>Obtention de la valeur d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.get](../../reference/shared/settings.get.md) pour obtenir la valeur d’un paramètre nommé « themeColor ». Le seul paramètre de la méthode **get** est le _name_ respectant la casse du paramètre.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 La méthode **get** retourne la valeur qui a été précédemment enregistrée pour le _name_ du paramètre qui a été passé. Si le paramètre n’existe pas, la méthode retourne **null**.


### <a name="removing-a-setting"></a>Suppression d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.remove](../../reference/shared/settings.removehandlerasync.md) pour supprimer un paramètre portant le nom « themeColor ». Le seul paramètre de la méthode **remove** est le _name_ respectant la casse du paramètre.


```
Office.context.document.settings.remove('themeColor');
```

Rien ne se produit si le paramètre n’existe pas. Utilisez la méthode  **Settings.saveAsync** pour faire persister la suppression du paramètre du document.


### <a name="saving-your-settings"></a>Enregistrement de vos paramètres

Pour enregistrer les ajouts, modifications ou suppressions que votre complément a effectués sur la copie en mémoire du conteneur de propriétés des paramètres pendant la session en cours, vous devez appeler la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md) pour les stocker dans le document. L’unique paramètre de la méthode **saveAsync** est _callback_, lequel est une fonction de rappel avec un paramètre unique. 


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

La fonction anonyme passée dans la méthode  **saveAsync** comme paramètre _callback_ est exécutée lorsque l’opération est terminée. Le paramètre _asyncResult_ du rappel permet d’accéder à un objet **AsyncResult** contenant le statut de l’opération. Dans l’exemple, la fonction vérifie la propriété **AsyncResult.status** pour savoir si l’opération d’enregistrement a réussi ou échoué, puis affiche le statut dans la page du complément.


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Enregistrement des paramètres dans la boîte aux lettres de l’utilisateur pour les compléments Outlook en tant que paramètres d’itinérance


Un complément Outlook peut utiliser l’objet [RoamingSettings ](../../reference/outlook/RoamingSettings.md) pour enregistrer les données d’état du complément et de paramètres propres à la boîte aux lettres de l’utilisateur. Ces données sont accessibles uniquement par ce complément Outlook au nom de l’utilisateur qui l’exécute. Les données sont stockées sur la boîte aux lettres Exchange Server de l’utilisateur et accessibles lorsque l’utilisateur se connecte à son compte et exécute le complément Outlook.


### <a name="loading-roaming-settings"></a>Chargement des paramètres d’itinérance


Un complément Outlook charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](../../reference/shared/office.initialize.md). L’exemple de code JavaScript suivant explique comment charger des paramètres d’itinérance existants.


```
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


Pour faire suite à l’exemple précédent, la fonction  `setAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) pour définir ou mettre à jour un paramètre nommé `cookie` avec la date du jour. Elle réenregistre ensuite tous les paramètres d’itinérance sur le serveur Exchange avec la méthode [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md).


```
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

La méthode  **saveAsync** enregistre les paramètres d’itinérance de manière asynchrone et admet une fonction de rappel facultative. Cet exemple de code transmet une fonction de rappel nommée `saveMyAppSettingsCallback` à la méthode **saveAsync**. Lors du renvoi de l’appel asynchrone, le paramètre  _asyncResult_ de la fonction `saveMyAppSettingsCallback` fournit un accès à un objet [AsyncResult](../../reference/outlook/simple-types.md) que vous pouvez utiliser pour déterminer le succès ou l’échec de l’opération avec la propriété **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Suppression d’un paramètre d’itinérance


Toujours dans le prolongement des exemples précédents, la fonction  `removeAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Enregistrement des paramètres par élément pour les compléments Outlook en tant que propriétés personnalisées


Les propriétés personnalisées permettent à votre complément Outlook de stocker des informations sur un élément qu’il utilise. Par exemple, si votre complément Outlook crée un rendez-vous à partir d’une suggestion de réunion dans un message, vous pouvez utiliser des propriétés personnalisées pour stocker le fait que la réunion a été créée. Cela garantit que si le message est rouvert, votre complément Outlook ne propose pas de recréer le rendez-vous.

Pour pouvoir utiliser des propriétés personnalisées pour un élément de message, de rendez-vous ou de demande de réunion particulier, vous devez charger les propriétés en mémoire en appelant la méthode [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) de l’objet **Item**. Si des propriétés personnalisées sont déjà définies pour l’élément actuel, elles sont chargées à ce moment à partir du serveur Exchange. Après avoir chargé les propriétés, vous pouvez utiliser les méthodes [set](../../reference/outlook/CustomProperties.md) et [get](../../reference/outlook/RoamingSettings.md) de l’objet **CustomProperties** pour ajouter, mettre à jour et récupérer des propriétés en mémoire. Pour enregistrer les modifications que vous avez apportées aux propriétés personnalisées de l’élément, vous devez utiliser la méthode [saveAsync](../../reference/outlook/CustomProperties.md) pour conserver les modifications de l’élément sur le serveur Exchange.


### <a name="custom-properties-example"></a>Exemple de propriétés personnalisées

L’exemple suivant illustre un ensemble simplifié des fonctions pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément Outlook qui utilise des propriétés personnalisées. 

Un complément Outlook qui utilise ces fonctions récupère des propriétés personnalisées en appelant la méthode  **get** sur la variable `_customProps`, comme le montre l’exemple suivant.




```
var property = _customProps.get("propertyName");
```

Cet exemple inclut les fonctions suivantes :



|**Nom de la fonction**|**Description**|
|:-----|:-----|
| `Office.initialize`|Initialise le complément et charge les propriétés personnalisées pour l’élément actuel à partir du serveur Exchange.|
| `customPropsCallback`|Obtient les propriétés personnalisées retournées du serveur Exchange et les enregistre pour une utilisation ultérieure.|
| `updateProperty`|Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur Exchange.|
| `removeProperty`|Supprime une propriété spécifique, puis fait persister la suppression sur le serveur Exchange.|
| `saveCallback`|Rappel pour les appels à la méthode  **saveAsync** dans les fonctions `updateProperty` et `removeProperty`.|



```
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


## <a name="additional-resources"></a>Ressources supplémentaires



- [Présentation de l’API JavaScript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
