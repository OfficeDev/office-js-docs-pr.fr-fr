---
title: Gérer l’état et les paramètres d’un Outlook de gestion
description: Découvrez comment faire persister l’état et les paramètres d’un Outlook un autre.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 69c22ab912d5099c42d6c69b364465a585cba1d4
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592009"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="43403-103">Gérer l’état et les paramètres d’un Outlook de gestion</span><span class="sxs-lookup"><span data-stu-id="43403-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="43403-104">Veuillez consulter [l’état et les paramètres persistants](../develop/persisting-add-in-state-and-settings.md) du module de mise en place dans la section **Concepts** de base de cette documentation avant de lire cet article.</span><span class="sxs-lookup"><span data-stu-id="43403-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="43403-105">Pour les Outlook, l’API JavaScript Office fournit des objets [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [CustomProperties](/javascript/api/outlook/office.customproperties) pour enregistrer l’état du add-in entre les sessions, comme décrit dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="43403-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="43403-106">Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](../reference/manifest/id.md) du complément qui les a créées.</span><span class="sxs-lookup"><span data-stu-id="43403-106">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="43403-107">**Objet**</span><span class="sxs-lookup"><span data-stu-id="43403-107">**Object**</span></span>|<span data-ttu-id="43403-108">**Emplacement de stockage**</span><span class="sxs-lookup"><span data-stu-id="43403-108">**Storage location**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="43403-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="43403-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="43403-110">Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé.</span><span class="sxs-lookup"><span data-stu-id="43403-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="43403-111">Étant donné que ces paramètres sont stockés dans la boîte aux lettres du serveur de l’utilisateur, ils peuvent « se déplacer » avec l’utilisateur et sont disponibles pour le module lorsqu’il est en cours d’exécution dans le contexte d’un client pris en charge accédant à la boîte aux lettres de cet utilisateur.</span><span class="sxs-lookup"><span data-stu-id="43403-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="43403-112">Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.</span><span class="sxs-lookup"><span data-stu-id="43403-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="43403-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="43403-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="43403-p103">Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.</span><span class="sxs-lookup"><span data-stu-id="43403-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="43403-116">Enregistrement des paramètres en tant que paramètres d’itinérance dans la boîte aux lettres de l’utilisateur pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="43403-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="43403-117">Un complément Outlook peut utiliser l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) pour enregistrer les données de paramètres et d’état du complément propres à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="43403-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="43403-118">Seul ce complément Outlook peut accéder aux données pour le compte de l’utilisateur qui exécute le complément.</span><span class="sxs-lookup"><span data-stu-id="43403-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="43403-119">Les données sont stockées dans la boîte aux lettres Exchange Server de l’utilisateur et sont accessibles lorsque cet utilisateur se connecte à son compte et exécute le complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="43403-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="43403-120">Chargement des paramètres d’itinérance</span><span class="sxs-lookup"><span data-stu-id="43403-120">Loading roaming settings</span></span>

<span data-ttu-id="43403-121">L’exemple de code JavaScript suivant explique comment charger des paramètres d’itinérance existants.</span><span class="sxs-lookup"><span data-stu-id="43403-121">The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="43403-122">Création ou affectation d’un paramètre d’itinérance</span><span class="sxs-lookup"><span data-stu-id="43403-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="43403-p105">Pour faire suite à l’exemple précédent, la fonction `setAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) pour définir ou mettre à jour un paramètre nommé `cookie` avec la date du jour. Elle réenregistre ensuite tous les paramètres d’itinérance sur le serveur Exchange avec la méthode [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="43403-p105">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

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

<span data-ttu-id="43403-125">La méthode **saveAsync** enregistre les paramètres d’itinérance de manière asynchrone et admet une fonction de rappel facultative.</span><span class="sxs-lookup"><span data-stu-id="43403-125">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="43403-126">Cet exemple de code transmet une fonction de rappel nommée `saveMyAppSettingsCallback` à la méthode **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="43403-126">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="43403-127">Lors du renvoi de l’appel asynchrone, le paramètre _asyncResult_ de la fonction `saveMyAppSettingsCallback` fournit un accès à un objet [AsyncResult](/javascript/api/office/office.asyncresult) que vous pouvez utiliser pour déterminer le succès ou l’échec de l’opération avec la propriété **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="43403-127">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/office/office.asyncresult) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="43403-128">Suppression d’un paramètre d’itinérance</span><span class="sxs-lookup"><span data-stu-id="43403-128">Removing a roaming setting</span></span>

<span data-ttu-id="43403-129">Toujours dans le prolongement des exemples précédents, la fonction  `removeAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="43403-129">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="43403-130">Enregistrement des paramètres par élément pour les compléments Outlook en tant que propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="43403-130">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="43403-p107">Les propriétés personnalisées permettent à votre complément Outlook de stocker des informations sur un élément qu’il utilise. Par exemple, si votre complément Outlook crée un rendez-vous à partir d’une suggestion de réunion dans un message, vous pouvez utiliser des propriétés personnalisées pour stocker le fait que la réunion a été créée. Cela garantit que si le message est rouvert, votre complément Outlook ne propose pas de recréer le rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="43403-p107">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="43403-p108">Pour pouvoir utiliser des propriétés personnalisées pour un élément de message, de rendez-vous ou de demande de réunion particulier, vous devez charger les propriétés en mémoire en appelant la méthode [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) de l’objet **Item**. Si des propriétés personnalisées sont déjà définies pour l’élément actuel, elles sont chargées à ce moment à partir du serveur Exchange. Après avoir chargé les propriétés, vous pouvez utiliser les méthodes [set](/javascript/api/outlook/office.customproperties#set-name--value-) et [get](/javascript/api/outlook/office.roamingsettings) de l’objet **CustomProperties** pour ajouter, mettre à jour et récupérer des propriétés en mémoire. Pour enregistrer les modifications que vous avez apportées aux propriétés personnalisées de l’élément, vous devez utiliser la méthode [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) pour conserver les modifications de l’élément sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="43403-p108">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="43403-138">Exemple de propriétés personnalisées</span><span class="sxs-lookup"><span data-stu-id="43403-138">Custom properties example</span></span>

<span data-ttu-id="43403-p109">L’exemple suivant illustre un ensemble simplifié des fonctions pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément Outlook qui utilise des propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="43403-p109">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span>

<span data-ttu-id="43403-141">Un complément Outlook qui utilise ces fonctions récupère toutes les propriétés personnalisées en appelant la méthode **get** sur la variable `_customProps`, comme le montre l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="43403-141">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="43403-142">Cet exemple inclut les fonctions suivantes :</span><span class="sxs-lookup"><span data-stu-id="43403-142">This example includes the following functions:</span></span>

|<span data-ttu-id="43403-143">**Nom de la fonction**</span><span class="sxs-lookup"><span data-stu-id="43403-143">**Function name**</span></span>|<span data-ttu-id="43403-144">**Description**</span><span class="sxs-lookup"><span data-stu-id="43403-144">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="43403-145">Initialise le complément et charge les propriétés personnalisées pour l’élément actuel à partir du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="43403-145">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="43403-146">Obtient les propriétés personnalisées retournées du serveur Exchange et les enregistre pour une utilisation ultérieure.</span><span class="sxs-lookup"><span data-stu-id="43403-146">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="43403-147">Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="43403-147">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="43403-148">Supprime une propriété spécifique, puis fait persister la suppression sur le serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="43403-148">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="43403-149">Rappel pour les appels à la méthode **saveAsync** dans les fonctions`updateProperty` et `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="43403-149">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

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

### <a name="platform-behavior-in-emails"></a><span data-ttu-id="43403-150">Comportement de la plateforme dans les e-mails</span><span class="sxs-lookup"><span data-stu-id="43403-150">Platform behavior in emails</span></span>

<span data-ttu-id="43403-151">Le tableau suivant récapitule le comportement des propriétés personnalisées enregistrées dans les e-mails pour Outlook clients.</span><span class="sxs-lookup"><span data-stu-id="43403-151">The following table summarizes saved custom properties behavior in emails for various Outlook clients.</span></span>

|<span data-ttu-id="43403-152">Scénario</span><span class="sxs-lookup"><span data-stu-id="43403-152">Scenario</span></span>|<span data-ttu-id="43403-153">Windows</span><span class="sxs-lookup"><span data-stu-id="43403-153">Windows</span></span>|<span data-ttu-id="43403-154">Web</span><span class="sxs-lookup"><span data-stu-id="43403-154">Web</span></span>|<span data-ttu-id="43403-155">Mac</span><span class="sxs-lookup"><span data-stu-id="43403-155">Mac</span></span>|
|---|---|---|---|
|<span data-ttu-id="43403-156">Nouvelle composition</span><span class="sxs-lookup"><span data-stu-id="43403-156">New compose</span></span>|<span data-ttu-id="43403-157">null</span><span class="sxs-lookup"><span data-stu-id="43403-157">null</span></span>|<span data-ttu-id="43403-158">null</span><span class="sxs-lookup"><span data-stu-id="43403-158">null</span></span>|<span data-ttu-id="43403-159">null</span><span class="sxs-lookup"><span data-stu-id="43403-159">null</span></span>|
|<span data-ttu-id="43403-160">Répondre, répondre à tous</span><span class="sxs-lookup"><span data-stu-id="43403-160">Reply, reply all</span></span>|<span data-ttu-id="43403-161">null</span><span class="sxs-lookup"><span data-stu-id="43403-161">null</span></span>|<span data-ttu-id="43403-162">null</span><span class="sxs-lookup"><span data-stu-id="43403-162">null</span></span>|<span data-ttu-id="43403-163">null</span><span class="sxs-lookup"><span data-stu-id="43403-163">null</span></span>|
|<span data-ttu-id="43403-164">Transférer</span><span class="sxs-lookup"><span data-stu-id="43403-164">Forward</span></span>|<span data-ttu-id="43403-165">Charge les propriétés du parent</span><span class="sxs-lookup"><span data-stu-id="43403-165">Loads parent's properties</span></span>|<span data-ttu-id="43403-166">null</span><span class="sxs-lookup"><span data-stu-id="43403-166">null</span></span>|<span data-ttu-id="43403-167">null</span><span class="sxs-lookup"><span data-stu-id="43403-167">null</span></span>|
|<span data-ttu-id="43403-168">Élément envoyé à partir d’une nouvelle composition</span><span class="sxs-lookup"><span data-stu-id="43403-168">Sent item from new compose</span></span>|<span data-ttu-id="43403-169">null</span><span class="sxs-lookup"><span data-stu-id="43403-169">null</span></span>|<span data-ttu-id="43403-170">null</span><span class="sxs-lookup"><span data-stu-id="43403-170">null</span></span>|<span data-ttu-id="43403-171">null</span><span class="sxs-lookup"><span data-stu-id="43403-171">null</span></span>|
|<span data-ttu-id="43403-172">Élément envoyé à partir de la réponse ou de la réponse à tous</span><span class="sxs-lookup"><span data-stu-id="43403-172">Sent item from reply or reply all</span></span>|<span data-ttu-id="43403-173">null</span><span class="sxs-lookup"><span data-stu-id="43403-173">null</span></span>|<span data-ttu-id="43403-174">null</span><span class="sxs-lookup"><span data-stu-id="43403-174">null</span></span>|<span data-ttu-id="43403-175">null</span><span class="sxs-lookup"><span data-stu-id="43403-175">null</span></span>|
|<span data-ttu-id="43403-176">Élément envoyé de l’avant</span><span class="sxs-lookup"><span data-stu-id="43403-176">Sent item from forward</span></span>|<span data-ttu-id="43403-177">Supprime les propriétés du parent s’il n’est pas enregistré</span><span class="sxs-lookup"><span data-stu-id="43403-177">Removes parent's properties if not saved</span></span>|<span data-ttu-id="43403-178">null</span><span class="sxs-lookup"><span data-stu-id="43403-178">null</span></span>|<span data-ttu-id="43403-179">null</span><span class="sxs-lookup"><span data-stu-id="43403-179">null</span></span>|

<span data-ttu-id="43403-180">Pour gérer la situation sur les Windows :</span><span class="sxs-lookup"><span data-stu-id="43403-180">To handle the situation on Windows:</span></span>

1. <span data-ttu-id="43403-181">Recherchez les propriétés existantes lors de l’initialisation de votre add-in, et conservez-les ou déséchantez-les selon vos besoins.</span><span class="sxs-lookup"><span data-stu-id="43403-181">Check for existing properties on initializing your add-in, and keep them or clear them as needed.</span></span>
1. <span data-ttu-id="43403-182">Lorsque vous définirez des propriétés personnalisées, incluez une propriété supplémentaire pour indiquer si les propriétés personnalisées ont été ajoutées lors de la lecture du message ou par mode lecture du complément.</span><span class="sxs-lookup"><span data-stu-id="43403-182">When setting custom properties, include an additional property to indicate whether the custom properties were added during message read or by Read mode of the add-in.</span></span> <span data-ttu-id="43403-183">Cela vous aidera à différencier si la propriété a été créée au cours de la composition ou héritée du parent.</span><span class="sxs-lookup"><span data-stu-id="43403-183">This will help you differentiate if the property was created during compose or inherited from the parent.</span></span>
1. <span data-ttu-id="43403-184">Pour vérifier si l’utilisateur envoie un e-mail ou répond, vous pouvez utiliser [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (disponible à partir de l’ensemble de conditions requises 1.10).</span><span class="sxs-lookup"><span data-stu-id="43403-184">To check if the user is forwarding an email or replying, you can use [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (available from requirement set 1.10).</span></span>

## <a name="see-also"></a><span data-ttu-id="43403-185">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="43403-185">See also</span></span>

- [<span data-ttu-id="43403-186">Conservation de l’état et des paramètres des compléments</span><span class="sxs-lookup"><span data-stu-id="43403-186">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="43403-187">Initialiser votre complément Office</span><span class="sxs-lookup"><span data-stu-id="43403-187">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)
