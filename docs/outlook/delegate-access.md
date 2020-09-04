---
title: Activer les scénarios d’accès délégué dans un complément Outlook
description: Décrit brièvement l’accès délégué et explique comment configurer la prise en charge des compléments.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 965798995638c9ec518d1e69ed23763d845fb7ec
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47363827"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="f11b9-103">Activer les scénarios d’accès délégué dans un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="f11b9-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="f11b9-104">Un propriétaire de boîte aux lettres peut utiliser la fonctionnalité accès délégué pour [permettre à quelqu’un d’autre de gérer son courrier et son calendrier](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="f11b9-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="f11b9-105">Cet article indique les autorisations déléguées prises en charge par l’API JavaScript pour Office et explique comment activer les scénarios d’accès délégué dans votre complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="f11b9-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f11b9-106">L’accès délégué n’est pas disponible actuellement dans Outlook sur Android et iOS.</span><span class="sxs-lookup"><span data-stu-id="f11b9-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="f11b9-107">En outre, cette fonctionnalité n’est pas disponible actuellement avec les [boîtes aux lettres partagées de groupe](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide#shared-mailboxes) dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="f11b9-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="f11b9-108">Cette fonctionnalité peut être rendue disponible à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="f11b9-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="f11b9-109">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="f11b9-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="f11b9-110">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="f11b9-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="f11b9-111">Autorisations prises en charge pour l’accès délégué</span><span class="sxs-lookup"><span data-stu-id="f11b9-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="f11b9-112">Le tableau suivant décrit les autorisations déléguées prises en charge par l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="f11b9-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="f11b9-113">Autorisation</span><span class="sxs-lookup"><span data-stu-id="f11b9-113">Permission</span></span>|<span data-ttu-id="f11b9-114">Valeur</span><span class="sxs-lookup"><span data-stu-id="f11b9-114">Value</span></span>|<span data-ttu-id="f11b9-115">Description</span><span class="sxs-lookup"><span data-stu-id="f11b9-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="f11b9-116">Lire</span><span class="sxs-lookup"><span data-stu-id="f11b9-116">Read</span></span>|<span data-ttu-id="f11b9-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="f11b9-117">1 (000001)</span></span>|<span data-ttu-id="f11b9-118">Peut lire des éléments.</span><span class="sxs-lookup"><span data-stu-id="f11b9-118">Can read items.</span></span>|
|<span data-ttu-id="f11b9-119">Write</span><span class="sxs-lookup"><span data-stu-id="f11b9-119">Write</span></span>|<span data-ttu-id="f11b9-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="f11b9-120">2 (000010)</span></span>|<span data-ttu-id="f11b9-121">Peut créer des éléments.</span><span class="sxs-lookup"><span data-stu-id="f11b9-121">Can create items.</span></span>|
|<span data-ttu-id="f11b9-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="f11b9-122">DeleteOwn</span></span>|<span data-ttu-id="f11b9-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="f11b9-123">4 (000100)</span></span>|<span data-ttu-id="f11b9-124">Peut uniquement supprimer les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="f11b9-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="f11b9-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="f11b9-125">DeleteAll</span></span>|<span data-ttu-id="f11b9-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="f11b9-126">8 (001000)</span></span>|<span data-ttu-id="f11b9-127">Peut supprimer tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="f11b9-127">Can delete any items.</span></span>|
|<span data-ttu-id="f11b9-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="f11b9-128">EditOwn</span></span>|<span data-ttu-id="f11b9-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="f11b9-129">16 (010000)</span></span>|<span data-ttu-id="f11b9-130">Ne peut modifier que les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="f11b9-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="f11b9-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="f11b9-131">EditAll</span></span>|<span data-ttu-id="f11b9-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="f11b9-132">32 (100000)</span></span>|<span data-ttu-id="f11b9-133">Peut modifier tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="f11b9-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="f11b9-134">Actuellement, l’API prend en charge l’obtention des autorisations de délégué existantes, mais pas la définition des autorisations de délégué.</span><span class="sxs-lookup"><span data-stu-id="f11b9-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="f11b9-135">L’objet [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) est implémenté à l’aide d’un masque de masques pour indiquer les autorisations du délégué.</span><span class="sxs-lookup"><span data-stu-id="f11b9-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="f11b9-136">Chaque position dans le masque de données représente une autorisation particulière et si elle est définie sur `1` Then, le délégué dispose de l’autorisation correspondante.</span><span class="sxs-lookup"><span data-stu-id="f11b9-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="f11b9-137">Par exemple, si le deuxième bit à partir de la droite est `1` , le délégué dispose alors d’une autorisation en **écriture** .</span><span class="sxs-lookup"><span data-stu-id="f11b9-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="f11b9-138">Vous pouvez voir un exemple sur la façon de vérifier une autorisation spécifique dans la section [effectuer une opération en tant que délégué](#perform-an-operation-as-delegate) plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="f11b9-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="f11b9-139">Synchronisation entre les clients de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f11b9-139">Sync across mailbox clients</span></span>

<span data-ttu-id="f11b9-140">Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées entre les boîtes aux lettres immédiatement.</span><span class="sxs-lookup"><span data-stu-id="f11b9-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="f11b9-141">Toutefois, si les opérations REST ou services Web Exchange (EWS) ont été utilisées pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons d’utiliser à la place l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter ce délai.</span><span class="sxs-lookup"><span data-stu-id="f11b9-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="f11b9-142">Pour plus d’informations, reportez-vous à la [section Propriétés personnalisées](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) de l’article « obtenir et définir des métadonnées dans un complément Outlook ».</span><span class="sxs-lookup"><span data-stu-id="f11b9-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f11b9-143">Dans un scénario de délégué, vous ne pouvez pas utiliser EWS avec les jetons actuellement fournis par office.js API.</span><span class="sxs-lookup"><span data-stu-id="f11b9-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f11b9-144">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="f11b9-144">Configure the manifest</span></span>

<span data-ttu-id="f11b9-145">Pour activer les scénarios d’accès délégué dans votre complément, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` dans le manifeste sous l’élément parent `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="f11b9-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="f11b9-146">Actuellement, les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f11b9-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="f11b9-147">Pour prendre en charge les appels REST à partir d’un délégué, définissez le nœud [autorisations](../reference/manifest/permissions.md) dans le manifeste sur `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="f11b9-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="f11b9-148">L’exemple suivant montre l' `SupportsSharedFolders` élément défini `true` dans une section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="f11b9-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="f11b9-149">Effectuer une opération en tant que délégué</span><span class="sxs-lookup"><span data-stu-id="f11b9-149">Perform an operation as delegate</span></span>

<span data-ttu-id="f11b9-150">Vous pouvez obtenir les propriétés partagées d’un élément en mode de composition ou de lecture en appelant la méthode [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="f11b9-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f11b9-151">Cela renvoie un objet [SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations du délégué, l’adresse de messagerie du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.</span><span class="sxs-lookup"><span data-stu-id="f11b9-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f11b9-152">Dans un scénario de délégué, votre complément peut utiliser REST mais pas EWS, et l’autorisation du complément doit être définie sur `ReadWriteMailbox` pour permettre l’accès Rest à la boîte aux lettres du propriétaire.</span><span class="sxs-lookup"><span data-stu-id="f11b9-152">In a delegate scenario, your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

<span data-ttu-id="f11b9-153">L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué dispose d’une autorisation en **écriture** et passer un appel Rest.</span><span class="sxs-lookup"><span data-stu-id="f11b9-153">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="f11b9-154">En tant que délégué, vous pouvez utiliser REST pour [obtenir le contenu d’un message Outlook attaché à un élément ou un billet de groupe Outlook](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span><span class="sxs-lookup"><span data-stu-id="f11b9-154">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="see-also"></a><span data-ttu-id="f11b9-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f11b9-155">See also</span></span>

- [<span data-ttu-id="f11b9-156">Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier</span><span class="sxs-lookup"><span data-stu-id="f11b9-156">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="f11b9-157">Partage de calendriers dans Office 365</span><span class="sxs-lookup"><span data-stu-id="f11b9-157">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="f11b9-158">Procédure de tri des éléments de manifeste</span><span class="sxs-lookup"><span data-stu-id="f11b9-158">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="f11b9-159">[Mask (Computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="f11b9-159">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="f11b9-160">Opérateurs de bits JavaScript</span><span class="sxs-lookup"><span data-stu-id="f11b9-160">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)