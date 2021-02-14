---
title: Activer les scénarios d’accès délégué dans un add-in Outlook
description: Décrit brièvement l’accès délégué et explique comment configurer la prise en charge des add-ins.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234239"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="5c843-103">Activer les scénarios d’accès délégué dans un add-in Outlook</span><span class="sxs-lookup"><span data-stu-id="5c843-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="5c843-104">Un propriétaire de boîte aux lettres peut utiliser la fonctionnalité d’accès délégué pour permettre à quelqu’un d’autre de [gérer son courrier et son calendrier.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)</span><span class="sxs-lookup"><span data-stu-id="5c843-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="5c843-105">Cet article spécifie les autorisations déléguées que l’API JavaScript Office prend en charge et décrit comment activer les scénarios d’accès délégué dans votre application Outlook.</span><span class="sxs-lookup"><span data-stu-id="5c843-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5c843-106">L’accès délégué n’est pas disponible actuellement dans Outlook sur Android et iOS.</span><span class="sxs-lookup"><span data-stu-id="5c843-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="5c843-107">En outre, cette fonctionnalité n’est pas disponible actuellement avec les boîtes aux lettres [partagées de](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) groupe dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="5c843-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="5c843-108">Cette fonctionnalité peut être rendue disponible à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="5c843-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="5c843-109">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.8.</span><span class="sxs-lookup"><span data-stu-id="5c843-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="5c843-110">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="5c843-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="5c843-111">Autorisations prise en charge pour l’accès délégué</span><span class="sxs-lookup"><span data-stu-id="5c843-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="5c843-112">Le tableau suivant décrit les autorisations déléguées que l’API JavaScript Office prend en charge.</span><span class="sxs-lookup"><span data-stu-id="5c843-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="5c843-113">Autorisation</span><span class="sxs-lookup"><span data-stu-id="5c843-113">Permission</span></span>|<span data-ttu-id="5c843-114">Valeur</span><span class="sxs-lookup"><span data-stu-id="5c843-114">Value</span></span>|<span data-ttu-id="5c843-115">Description</span><span class="sxs-lookup"><span data-stu-id="5c843-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="5c843-116">Lire</span><span class="sxs-lookup"><span data-stu-id="5c843-116">Read</span></span>|<span data-ttu-id="5c843-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="5c843-117">1 (000001)</span></span>|<span data-ttu-id="5c843-118">Peut lire des éléments.</span><span class="sxs-lookup"><span data-stu-id="5c843-118">Can read items.</span></span>|
|<span data-ttu-id="5c843-119">Write</span><span class="sxs-lookup"><span data-stu-id="5c843-119">Write</span></span>|<span data-ttu-id="5c843-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="5c843-120">2 (000010)</span></span>|<span data-ttu-id="5c843-121">Peut créer des éléments.</span><span class="sxs-lookup"><span data-stu-id="5c843-121">Can create items.</span></span>|
|<span data-ttu-id="5c843-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="5c843-122">DeleteOwn</span></span>|<span data-ttu-id="5c843-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="5c843-123">4 (000100)</span></span>|<span data-ttu-id="5c843-124">Peut supprimer uniquement les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="5c843-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="5c843-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="5c843-125">DeleteAll</span></span>|<span data-ttu-id="5c843-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="5c843-126">8 (001000)</span></span>|<span data-ttu-id="5c843-127">Peut supprimer tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="5c843-127">Can delete any items.</span></span>|
|<span data-ttu-id="5c843-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="5c843-128">EditOwn</span></span>|<span data-ttu-id="5c843-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="5c843-129">16 (010000)</span></span>|<span data-ttu-id="5c843-130">Peut modifier uniquement les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="5c843-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="5c843-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="5c843-131">EditAll</span></span>|<span data-ttu-id="5c843-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="5c843-132">32 (100000)</span></span>|<span data-ttu-id="5c843-133">Peut modifier n’importe quel objet.</span><span class="sxs-lookup"><span data-stu-id="5c843-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="5c843-134">Actuellement, l’API prend en charge l’obtention d’autorisations déléguées existantes, mais pas la définition d’autorisations déléguées.</span><span class="sxs-lookup"><span data-stu-id="5c843-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="5c843-135">[L’objet DelegatePermissions est](/javascript/api/outlook/office.mailboxenums.delegatepermissions) implémenté à l’aide d’un masque de bits pour indiquer les autorisations du délégué.</span><span class="sxs-lookup"><span data-stu-id="5c843-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="5c843-136">Chaque position dans le masque de bits représente une autorisation particulière et, si elle est définie sur, le délégué dispose de `1` l’autorisation respective.</span><span class="sxs-lookup"><span data-stu-id="5c843-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="5c843-137">Par exemple, si le deuxième bit à partir de la droite est `1` , le délégué dispose de **l’autorisation d’écriture.**</span><span class="sxs-lookup"><span data-stu-id="5c843-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="5c843-138">Vous pouvez voir un exemple de vérification d’une autorisation spécifique dans la [section](#perform-an-operation-as-delegate) Effectuer une opération en tant que délégué plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="5c843-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="5c843-139">Synchronisation entre les clients de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5c843-139">Sync across mailbox clients</span></span>

<span data-ttu-id="5c843-140">Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées immédiatement entre les boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="5c843-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="5c843-141">Toutefois, si des opérations REST ou services web Exchange (EWS) ont été utilisées pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons plutôt d’utiliser [l’objet CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter ce délai.</span><span class="sxs-lookup"><span data-stu-id="5c843-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="5c843-142">Pour en savoir plus, consultez la [section des](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) propriétés personnalisées de l’article « Obtenir et définir des métadonnées dans un add-in Outlook ».</span><span class="sxs-lookup"><span data-stu-id="5c843-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5c843-143">Dans un scénario de délégué, vous ne pouvez pas utiliser EWS avec les jetons actuellement fournis par office.js API.</span><span class="sxs-lookup"><span data-stu-id="5c843-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="5c843-144">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="5c843-144">Configure the manifest</span></span>

<span data-ttu-id="5c843-145">Pour activer les scénarios d’accès délégué dans votre add-in, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) dans le manifeste sous `true` l’élément `DesktopFormFactor` parent.</span><span class="sxs-lookup"><span data-stu-id="5c843-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="5c843-146">Pour l’instant, les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="5c843-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="5c843-147">Pour prendre en charge les appels REST d’un délégué, définissez le nœud [Autorisations](../reference/manifest/permissions.md) dans le manifeste sur `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="5c843-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="5c843-148">L’exemple suivant montre `SupportsSharedFolders` l’ensemble `true` d’éléments dans une section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="5c843-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="5c843-149">Effectuer une opération en tant que délégué</span><span class="sxs-lookup"><span data-stu-id="5c843-149">Perform an operation as delegate</span></span>

<span data-ttu-id="5c843-150">Vous pouvez obtenir les propriétés partagées d’un élément en mode Composition ou Lecture en appelant la méthode [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="5c843-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="5c843-151">Cela renvoie un [objet SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations du délégué, l’adresse e-mail du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.</span><span class="sxs-lookup"><span data-stu-id="5c843-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="5c843-152">L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué dispose d’une autorisation d’écriture et passer un appel REST. </span><span class="sxs-lookup"><span data-stu-id="5c843-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="5c843-153">En tant que délégué, vous pouvez utiliser REST pour obtenir le contenu d’un message Outlook joint à un élément Outlook ou à un [billet de groupe.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="5c843-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="5c843-154">Gérer l’appel de REST sur les éléments partagés et non partagés</span><span class="sxs-lookup"><span data-stu-id="5c843-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="5c843-155">Si vous souhaitez appeler une opération REST sur un élément, que l’élément soit partagé ou non, vous pouvez utiliser l’API pour déterminer si l’élément `getSharedPropertiesAsync` est partagé.</span><span class="sxs-lookup"><span data-stu-id="5c843-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="5c843-156">Après cela, vous pouvez construire l’URL REST pour l’opération à l’aide de l’objet approprié.</span><span class="sxs-lookup"><span data-stu-id="5c843-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="5c843-157">Limites</span><span class="sxs-lookup"><span data-stu-id="5c843-157">Limitations</span></span>

<span data-ttu-id="5c843-158">Selon les scénarios de votre add-in, vous devez tenir compte de quelques limitations lors de la gestion des situations de délégué.</span><span class="sxs-lookup"><span data-stu-id="5c843-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="5c843-159">REST et EWS</span><span class="sxs-lookup"><span data-stu-id="5c843-159">REST and EWS</span></span>

<span data-ttu-id="5c843-160">Votre add-in peut utiliser REST, mais pas EWS, et l’autorisation du module doit être définie pour activer l’accès REST à la boîte aux lettres `ReadWriteMailbox` du propriétaire.</span><span class="sxs-lookup"><span data-stu-id="5c843-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="5c843-161">Mode composition de message</span><span class="sxs-lookup"><span data-stu-id="5c843-161">Message Compose mode</span></span>

<span data-ttu-id="5c843-162">En mode composition de message, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) n’est pas pris en charge dans Outlook sur le web ou Windows, sauf si les conditions suivantes sont remplies.</span><span class="sxs-lookup"><span data-stu-id="5c843-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="5c843-163">Le propriétaire partage au moins un dossier de boîte aux lettres avec le délégué.</span><span class="sxs-lookup"><span data-stu-id="5c843-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="5c843-164">Le délégué rédige un message dans le dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="5c843-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="5c843-165">Exemples :</span><span class="sxs-lookup"><span data-stu-id="5c843-165">Examples:</span></span>

    - <span data-ttu-id="5c843-166">Le délégué répond ou envoie un message électronique dans le dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="5c843-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="5c843-167">Le délégué enregistre un brouillon, puis le déplace de son propre dossier **Brouillons** vers le dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="5c843-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="5c843-168">Le délégué ouvre le brouillon à partir du dossier partagé, puis continue la composition.</span><span class="sxs-lookup"><span data-stu-id="5c843-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="5c843-169">Une fois le message envoyé, il se trouve généralement dans le dossier Éléments envoyés **du** délégué.</span><span class="sxs-lookup"><span data-stu-id="5c843-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="5c843-170">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5c843-170">See also</span></span>

- [<span data-ttu-id="5c843-171">Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier</span><span class="sxs-lookup"><span data-stu-id="5c843-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="5c843-172">Partage de calendrier dans Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="5c843-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="5c843-173">Comment commander des éléments de manifeste</span><span class="sxs-lookup"><span data-stu-id="5c843-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="5c843-174">[Masque (calcul)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="5c843-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="5c843-175">Opérateurs de bits JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c843-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)