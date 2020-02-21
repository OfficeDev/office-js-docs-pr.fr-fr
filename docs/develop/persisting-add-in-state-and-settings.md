---
title: Conservation de l’état et des paramètres des compléments
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 69fc0b1316a1a4eb0dfe0ebea01ffdbfe88dcd8c
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163506"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="b9935-102">Conservation de l’état et des paramètres des compléments</span><span class="sxs-lookup"><span data-stu-id="b9935-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="b9935-p101">Les compléments Office sont essentiellement des applications web exécutées dans l’environnement sans état d’un contrôle de navigateur. En conséquence, votre complément devra peut-être faire persister les données pour assurer la continuité de certaines opérations ou fonctionnalités entre les sessions d’utilisation du complément. Par exemple, votre complément peut disposer de paramètres personnalisés ou d’autres valeurs dont il a besoin pour l’enregistrement et le rechargement à la prochaine initialisation, tels que l’affichage préféré d’un utilisateur ou l’emplacement par défaut. Pour ce faire, vous pouvez procéder comme suit :</span><span class="sxs-lookup"><span data-stu-id="b9935-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="b9935-107">Utilisez les membres de l’API JavaScript pour Office qui stockent les données sous l’une des formes suivantes :</span><span class="sxs-lookup"><span data-stu-id="b9935-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="b9935-108">Paires nom/valeur dans un conteneur de propriétés stocké dans un emplacement qui dépend du type de complément.</span><span class="sxs-lookup"><span data-stu-id="b9935-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="b9935-109">Éléments XML personnalisés stockés dans le document.</span><span class="sxs-lookup"><span data-stu-id="b9935-109">Custom XML stored in the document.</span></span>

- <span data-ttu-id="b9935-110">Utilisez des techniques fournies par le contrôle de navigateur sous-jacent : les cookies de navigateur ou le stockage web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="b9935-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="b9935-p102">Cet article se concentre sur l’utilisation de l’interface API JavaScript pour Office afin de faire persister l’état du complément. Pour obtenir des exemples d’utilisation des cookies de navigateur et du stockage web, voir l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="b9935-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="b9935-113">Persistance de l’état et des paramètres d’un complément avec l’interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="b9935-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="b9935-p103">L’interface API JavaScript pour Office fournit les objets [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [CustomProperties](/javascript/api/outlook/office.customproperties) pour enregistrer l’état du complément dans plusieurs sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](/office/dev/add-ins/reference/manifest/id) du complément qui les a créées.</span><span class="sxs-lookup"><span data-stu-id="b9935-p103">The JavaScript API for Office provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](/office/dev/add-ins/reference/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="b9935-116">**Objet**</span><span class="sxs-lookup"><span data-stu-id="b9935-116">**Object**</span></span>|<span data-ttu-id="b9935-117">**Type de complément**</span><span class="sxs-lookup"><span data-stu-id="b9935-117">**Add-in type support**</span></span>|<span data-ttu-id="b9935-118">**Emplacement de stockage**</span><span class="sxs-lookup"><span data-stu-id="b9935-118">**Storage location**</span></span>|<span data-ttu-id="b9935-119">**ôte Office**</span><span class="sxs-lookup"><span data-stu-id="b9935-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b9935-120">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b9935-120">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="b9935-121">Contenu et volet de tâches</span><span class="sxs-lookup"><span data-stu-id="b9935-121">content and task pane</span></span>|<span data-ttu-id="b9935-122">Document, feuille de calcul ou présentation le complément collabore avec lequel le complément fonctionne. Les paramètres de complément de contenu et de volet Office sont disponibles pour le complément qui les a créés dans le document dans lequel ils sont enregistrés.</span><span class="sxs-lookup"><span data-stu-id="b9935-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="b9935-p104">**Remarque importante :** ne stockez pas de mots de passe ou autres informations d’identification personnelle (PII) avec l’objet **Settings**. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation de PII de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9935-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="b9935-126">Word, Excel ou PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b9935-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="b9935-p105">**Remarque :** les compléments du volet Office pour Project 2013 ne prennent pas en charge l’API **Settings** pour le stockage de l’état ou des paramètres du complément. Cependant, pour les compléments exécutés dans Project (et d’autres applications hôtes Office), vous pouvez utiliser des techniques telles que les cookies de navigateur ou le stockage web. Pour plus d’informations sur ces techniques, reportez-vous à l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="b9935-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="b9935-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9935-130">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="b9935-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-131">Outlook</span></span>|<span data-ttu-id="b9935-132">Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé. Comme ces paramètres sont stockés dans la boîte aux lettres de serveur de l’utilisateur, ils sont itinérants et accessibles par le complément lorsqu’il s’exécute dans le contexte d’une application hôte cliente ou d’un navigateur pris en charge accédant à la boîte aux lettres de cet utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9935-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="b9935-133">Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.</span><span class="sxs-lookup"><span data-stu-id="b9935-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="b9935-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-134">Outlook</span></span>|
|[<span data-ttu-id="b9935-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="b9935-135">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="b9935-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-136">Outlook</span></span>|<span data-ttu-id="b9935-p106">Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.</span><span class="sxs-lookup"><span data-stu-id="b9935-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="b9935-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-139">Outlook</span></span>|
|[<span data-ttu-id="b9935-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9935-140">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="b9935-141">volet Office</span><span class="sxs-lookup"><span data-stu-id="b9935-141">task pane</span></span>|<span data-ttu-id="b9935-p107">Document, feuille de calcul ou présentation que le complément utilise. Les paramètres de complément de volet Office sont disponibles pour le complément qui les a créés dans le document dans lequel ils sont enregistrés.</span><span class="sxs-lookup"><span data-stu-id="b9935-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="b9935-p108">**Important :** ne stockez pas de mot de passe ni d’autres informations d’identification personnelle dans une partie XML personnalisée. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation des informations d’identification personnelle de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9935-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="b9935-147">Word (à l’aide de l’API JavaScript courante pour Office) Excel (à l’aide de l’API JavaScript pour Excel propre à l’hôte</span><span class="sxs-lookup"><span data-stu-id="b9935-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="b9935-148">Données de paramètres gérées en mémoire à l’exécution</span><span class="sxs-lookup"><span data-stu-id="b9935-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="b9935-p109">Les deux sections suivantes abordent les paramètres dans le contexte de l’API JavaScript courante pour Office. L’API JavaScript pour Excel propre à un hôte propose également un accès aux paramètres personnalisés. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [SettingCollection pour Excel](/javascript/api/excel/excel.settingcollection).</span><span class="sxs-lookup"><span data-stu-id="b9935-p109">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="b9935-153">En interne, les données du conteneur des propriétés accessibles avec les objets **Settings**, **CustomProperties** ou **RoamingSettings** sont stockées en tant qu’objet JSON (JavaScript Object Notation) sérialisé contenant des paires nom/valeur.</span><span class="sxs-lookup"><span data-stu-id="b9935-153">Internally, the data in the property bag accessed with the **Settings**, **CustomProperties**, or **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="b9935-154">Le nom (clé) de chaque valeur doit être de type **string** (chaîne), et la valeur stockée peut être de type **string** (chaîne), **number** (nombre), **date** ou **object** (objet) JavaScript, mais pas de type **function** (fonction).</span><span class="sxs-lookup"><span data-stu-id="b9935-154">The name (key) for each value must be a **string**, and the stored value can be a JavaScript **string**, **number**, **date**, or **object**, but not a **function**.</span></span>

<span data-ttu-id="b9935-155">Cet exemple de structure de conteneur des propriétés contient trois valeurs de type **string** (chaîne) définies, nommées `firstName`, `location` et `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="b9935-155">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="b9935-156">Après avoir enregistré le conteneur des propriétés de paramètres durant la session de complément précédente, vous pouvez le charger pendant ou après l’initialisation du complément, durant la session actuelle du complément.</span><span class="sxs-lookup"><span data-stu-id="b9935-156">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="b9935-157">Lors de la session, les paramètres sont gérés intégralement en mémoire à l’aide des méthodes **get**, **set**, et **remove** de l’objet correspondant au type de paramètre que vous créez (**Settings**, **CustomProperties**, ou **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="b9935-157">During the session, the settings are managed in entirely in memory using the **get**, **set**, and **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**, **CustomProperties**, or **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="b9935-158">Pour conserver les mises à jour, suppressions ou ajouts effectués pendant la session actuelle du complément dans l’emplacement de stockage, vous devez appeler la méthode **saveAsync** de l’objet correspondant utilisé pour le fonctionnement de ce genre de paramètres.</span><span class="sxs-lookup"><span data-stu-id="b9935-158">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the **saveAsync** method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="b9935-159">Les méthodes **get**, **set** et **remove** s’appliquent uniquement à la copie en mémoire du conteneur des propriétés de paramètres.</span><span class="sxs-lookup"><span data-stu-id="b9935-159">The **get**, **set**, and **remove** methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="b9935-160">Si vous fermez votre complément sans appeler la méthode **saveAsync**, vous perdrez les modifications apportées aux paramètres pendant cette session.</span><span class="sxs-lookup"><span data-stu-id="b9935-160">If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="b9935-161">Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office</span><span class="sxs-lookup"><span data-stu-id="b9935-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="b9935-p113">Pour conserver l’état ou les paramètres personnalisés d’un complément de contenu ou du volet Office pour Word, Excel ou PowerPoint, utilisez l’objet [Settings](/javascript/api/office/office.settings) et ses méthodes. Le conteneur de propriétés créé à l’aide des méthodes de l’objet **Settings** est accessible uniquement par l’instance du complément de contenu ou du volet Office qui l’a créé, et uniquement à partir du document dans lequel il est enregistré.</span><span class="sxs-lookup"><span data-stu-id="b9935-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="b9935-164">L’objet **Settings** est automatiquement chargé en tant que partie de l’objet [Document](/javascript/api/office/office.document). Il est disponible lors de l’activation du volet Office ou d’un complément de contenu.</span><span class="sxs-lookup"><span data-stu-id="b9935-164">The **Settings** object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="b9935-165">Après l’instanciation de l’objet **Document**, vous pouvez accéder à l’objet **Settings** avec la propriété [settings](/javascript/api/office/office.document#settings) de l’objet **Document**.</span><span class="sxs-lookup"><span data-stu-id="b9935-165">After the **Document** object is instantiated, you can access the **Settings** object with the [settings](/javascript/api/office/office.document#settings) property of the **Document** object.</span></span> <span data-ttu-id="b9935-166">Pendant cette session, vous pouvez uniquement utiliser les méthodes **Settings.get**, **Settings.set** et **Settings.remove**. Elles permettent de lire, écrire ou supprimer les paramètres et l’état du complément conservés dans la copie en mémoire du conteneur des propriétés.</span><span class="sxs-lookup"><span data-stu-id="b9935-166">During the lifetime of the session, you can just use the **Settings.get**, **Settings.set**, and **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="b9935-167">Étant donné que les méthodes de définition (set) et de suppression (remove) fonctionnent uniquement par rapport à la copie en mémoire du conteneur des propriétés de paramètres, pour enregistrer de nouveaux paramètres ou des paramètres modifiés dans le document auquel le complément est associé, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="b9935-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="b9935-168">Création ou mise à jour d’une valeur de paramètre</span><span class="sxs-lookup"><span data-stu-id="b9935-168">Creating or updating a setting value</span></span>

<span data-ttu-id="b9935-p115">L’exemple de code suivant montre comment utiliser la méthode [Settings.set](/javascript/api/office/office.settings#set-name--value-) pour créer un paramètre appelé `'themeColor'` avec la valeur `'green'`. Le premier paramètre de la méthode set est le _name_ (ID) respectant la casse du paramètre à définir ou à créer. Le second paramètre est la _value_ du paramètre.</span><span class="sxs-lookup"><span data-stu-id="b9935-p115">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="b9935-p116">Le paramètre avec le nom spécifié est créé s’il n’existe pas déjà ou sa valeur est mise à jour s’il existe. Utilisez la méthode **Settings.saveAsync** pour rendre persistants les paramètres (nouveaux ou mis à jour) du document.</span><span class="sxs-lookup"><span data-stu-id="b9935-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="b9935-174">Obtention de la valeur d’un paramètre</span><span class="sxs-lookup"><span data-stu-id="b9935-174">Getting the value of a setting</span></span>

<span data-ttu-id="b9935-p117">L’exemple suivant illustre comment utiliser la méthode [Settings.get](/javascript/api/office/office.settings#get-name-) pour obtenir la valeur d’un paramètre nommé « themeColor ». Le seul paramètre de la méthode **get** est le _name_ respectant la casse du paramètre.</span><span class="sxs-lookup"><span data-stu-id="b9935-p117">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="b9935-p118">La méthode **get** retourne la valeur qui a été précédemment enregistrée pour le _name_ du paramètre qui a été passé. Si le paramètre n’existe pas, la méthode retourne **null**.</span><span class="sxs-lookup"><span data-stu-id="b9935-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="b9935-179">Suppression d’un paramètre</span><span class="sxs-lookup"><span data-stu-id="b9935-179">Removing a setting</span></span>

<span data-ttu-id="b9935-p119">L’exemple suivant montre comment utiliser la méthode [Settings.remove](/javascript/api/office/office.settings#remove-name-) pour supprimer un paramètre portant le nom « themeColor ». Le seul paramètre de la méthode **remove** est le _name_ respectant la casse du paramètre.</span><span class="sxs-lookup"><span data-stu-id="b9935-p119">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="b9935-182">Rien ne se produit si le paramètre n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="b9935-182">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="b9935-183">Utilisez la méthode **Settings.saveAsync** pour faire persister la suppression du paramètre du document.</span><span class="sxs-lookup"><span data-stu-id="b9935-183">Use the **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="b9935-184">Enregistrement de vos paramètres</span><span class="sxs-lookup"><span data-stu-id="b9935-184">Saving your settings</span></span>

<span data-ttu-id="b9935-p121">Pour enregistrer les modifications, suppressions ou ajouts que votre complément a effectués sur la copie en mémoire du conteneur de propriétés des paramètres pendant la session en cours, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) pour les stocker dans le document. L’unique paramètre de la méthode **saveAsync** est _callback_, lequel est une fonction de rappel avec un paramètre unique.</span><span class="sxs-lookup"><span data-stu-id="b9935-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="b9935-187">Le système exécute la fonction anonyme transmise dans la méthode **saveAsync** en tant que paramètre _callback_ lorsque l’opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="b9935-187">The anonymous function passed into the **saveAsync** method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="b9935-188">Le paramètre _asyncResult_ du rappel permet d’accéder à un objet **AsyncResult** qui contient l’état de l’opération.</span><span class="sxs-lookup"><span data-stu-id="b9935-188">The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation.</span></span> <span data-ttu-id="b9935-189">Dans l’exemple, la fonction vérifie la propriété **AsyncResult.status** pour déterminer si l’opération d’enregistrement a réussi ou échoué, puis affiche le résultat sur la page du complément.</span><span class="sxs-lookup"><span data-stu-id="b9935-189">In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="b9935-190">Enregistrement du XML personnalisé dans le document</span><span class="sxs-lookup"><span data-stu-id="b9935-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="b9935-p123">Cette section décrit les parties XML personnalisées dans le contexte de l’API JavaScript courante pour Office qui est prise en charge dans Word. L’API JavaScript pour Excel propre à un hôte permet également d’accéder aux parties XML personnalisées. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [CustomXmlPart pour Excel](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="b9935-p123">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="b9935-195">Une option de stockage supplémentaire est disponible lorsque vous avez besoin de stocker des informations dépassant les limites de taille des paramètres du document ou comportant un caractère structuré.</span><span class="sxs-lookup"><span data-stu-id="b9935-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="b9935-196">Vous pouvez conserver le balisage XML personnalisé dans un complément de volet Office pour Word (et pour Excel, mais reportez-vous à la remarque en haut de cette section).</span><span class="sxs-lookup"><span data-stu-id="b9935-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="b9935-197">Dans Word, utilisez l’objet [CustomXmlPart](/javascript/api/office/office.customxmlpart) et ses méthodes (rappel : pour Excel, consultez la note précédente).</span><span class="sxs-lookup"><span data-stu-id="b9935-197">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="b9935-198">Le code suivant crée une partie XML personnalisée, puis affiche son ID et son contenu dans des balises div sur la page.</span><span class="sxs-lookup"><span data-stu-id="b9935-198">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="b9935-199">Un attribut `xmlns` doit figurer dans la chaîne XML.</span><span class="sxs-lookup"><span data-stu-id="b9935-199">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

<span data-ttu-id="b9935-p125">Pour récupérer une partie XML personnalisée, vous utilisez la méthode [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mais l’identifiant correspond à un GUID généré lorsque la partie XML est créée. Vous ne pouvez donc pas connaître l’identifiant lors du codage. Pour cette raison, il est recommandé de stocker immédiatement l’identifiant de la partie XML en tant que paramètre et de lui donner une clé facilement mémorisable lorsque vous créez une partie XML. L’exemple de méthode suivant montre comment procéder. (Toutefois, reportez-vous aux sections précédentes de cet article pour obtenir des détails et des meilleures pratiques lorsque vous utilisez des paramètres personnalisés).</span><span class="sxs-lookup"><span data-stu-id="b9935-p125">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="b9935-204">Le code suivant montre comment récupérer la partie XML en obtenant d’abord son identifiant partir d’un paramètre.</span><span class="sxs-lookup"><span data-stu-id="b9935-204">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="b9935-205">Enregistrement des paramètres en tant que paramètres d’itinérance dans la boîte aux lettres de l’utilisateur pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-205">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="b9935-206">Un complément Outlook peut utiliser l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) pour enregistrer les données de paramètres et d’état du complément propres à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9935-206">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="b9935-207">Seul ce complément Outlook peut accéder aux données pour le compte de l’utilisateur qui exécute le complément.</span><span class="sxs-lookup"><span data-stu-id="b9935-207">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="b9935-208">Les données sont stockées dans la boîte aux lettres Exchange Server de l’utilisateur et sont accessibles lorsque cet utilisateur se connecte à son compte et exécute le complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9935-208">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="b9935-209">Chargement des paramètres d’itinérance</span><span class="sxs-lookup"><span data-stu-id="b9935-209">Loading roaming settings</span></span>


<span data-ttu-id="b9935-p127">Un complément Outlook charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](/javascript/api/office). L’exemple de code JavaScript suivant explique comment charger des paramètres d’itinérance existants.</span><span class="sxs-lookup"><span data-stu-id="b9935-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="b9935-212">Création ou affectation d’un paramètre d’itinérance</span><span class="sxs-lookup"><span data-stu-id="b9935-212">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="b9935-p128">Pour faire suite à l’exemple précédent, la fonction `setAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) pour définir ou mettre à jour un paramètre nommé `cookie` avec la date du jour. Elle réenregistre ensuite tous les paramètres d’itinérance sur le serveur Exchange avec la méthode [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="b9935-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>


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

<span data-ttu-id="b9935-215">La méthode **saveAsync** enregistre les paramètres d’itinérance de manière asynchrone et admet une fonction de rappel facultative.</span><span class="sxs-lookup"><span data-stu-id="b9935-215">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="b9935-216">Cet exemple de code transmet une fonction de rappel nommée `saveMyAppSettingsCallback` à la méthode **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="b9935-216">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="b9935-217">Lors du renvoi de l’appel asynchrone, le paramètre _asyncResult_ de la fonction `saveMyAppSettingsCallback` fournit un accès à un objet [AsyncResult](/javascript/api/outlook) que vous pouvez utiliser pour déterminer le succès ou l’échec de l’opération avec la propriété**AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="b9935-217">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/outlook) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="b9935-218">Suppression d’un paramètre d’itinérance</span><span class="sxs-lookup"><span data-stu-id="b9935-218">Removing a roaming setting</span></span>


<span data-ttu-id="b9935-219">Toujours dans le prolongement des exemples précédents, la fonction  `removeAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9935-219">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="b9935-220">Enregistrement des paramètres par élément pour les compléments Outlook en tant que propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="b9935-220">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="b9935-p130">Les propriétés personnalisées permettent à votre complément Outlook de stocker des informations sur un élément qu’il utilise. Par exemple, si votre complément Outlook crée un rendez-vous à partir d’une suggestion de réunion dans un message, vous pouvez utiliser des propriétés personnalisées pour stocker le fait que la réunion a été créée. Cela garantit que si le message est rouvert, votre complément Outlook ne propose pas de recréer le rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b9935-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="b9935-p131">Pour pouvoir utiliser des propriétés personnalisées pour un élément de message, de rendez-vous ou de demande de réunion particulier, vous devez charger les propriétés en mémoire en appelant la méthode [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) de l’objet **Item**. Si des propriétés personnalisées sont déjà définies pour l’élément actuel, elles sont chargées à ce moment à partir du serveur Exchange. Après avoir chargé les propriétés, vous pouvez utiliser les méthodes [set](/javascript/api/outlook/office.customproperties#set-name--value-) et [get](/javascript/api/outlook/office.roamingsettings) de l’objet **CustomProperties** pour ajouter, mettre à jour et récupérer des propriétés en mémoire. Pour enregistrer les modifications que vous avez apportées aux propriétés personnalisées de l’élément, vous devez utiliser la méthode [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) pour conserver les modifications de l’élément sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9935-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="b9935-228">Exemple de propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="b9935-228">Custom properties example</span></span>

<span data-ttu-id="b9935-p132">L’exemple suivant illustre un ensemble simplifié des fonctions pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément Outlook qui utilise des propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b9935-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="b9935-231">Un complément Outlook qui utilise ces fonctions récupère toutes les propriétés personnalisées en appelant la méthode **get** sur la variable `_customProps`, comme le montre l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="b9935-231">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="b9935-232">Cet exemple inclut les fonctions suivantes :</span><span class="sxs-lookup"><span data-stu-id="b9935-232">This example includes the following functions:</span></span>



|<span data-ttu-id="b9935-233">**Nom de la fonction**</span><span class="sxs-lookup"><span data-stu-id="b9935-233">**Function name**</span></span>|<span data-ttu-id="b9935-234">**Description**</span><span class="sxs-lookup"><span data-stu-id="b9935-234">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="b9935-235">Initialise le complément et charge les propriétés personnalisées pour l’élément actuel à partir du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9935-235">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="b9935-236">Obtient les propriétés personnalisées retournées du serveur Exchange et les enregistre pour une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9935-236">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="b9935-237">Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9935-237">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="b9935-238">Supprime une propriété spécifique, puis fait persister la suppression sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b9935-238">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="b9935-239">Rappel pour les appels à la méthode**saveAsync** dans les fonctions`updateProperty` et `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="b9935-239">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="b9935-240">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b9935-240">See also</span></span>

- [<span data-ttu-id="b9935-241">Présentation de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="b9935-241">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="b9935-242">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="b9935-242">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="b9935-243">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="b9935-243">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
