---
title: Activer les scénarios d’accès délégué dans un complément Outlook
description: Décrit brièvement l’accès délégué et explique comment configurer la prise en charge des compléments.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6cee68af9efc02bbb474effaba1a898511aea531
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166206"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="5d708-103">Activer les scénarios d’accès délégué dans un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="5d708-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="5d708-104">Un propriétaire de boîte aux lettres peut utiliser la fonctionnalité accès délégué pour [permettre à quelqu’un d’autre de gérer son courrier et son calendrier](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="5d708-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="5d708-105">Cet article indique les autorisations déléguées prises en charge par l’API JavaScript pour Office et explique comment activer les scénarios d’accès délégué dans votre complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="5d708-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5d708-106">L’accès délégué n’est pas disponible actuellement dans Outlook sur Mac, Android et iOS.</span><span class="sxs-lookup"><span data-stu-id="5d708-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="5d708-107">Cette fonctionnalité peut être rendue disponible à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="5d708-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="5d708-108">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="5d708-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="5d708-109">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="5d708-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="5d708-110">Autorisations prises en charge pour l’accès délégué</span><span class="sxs-lookup"><span data-stu-id="5d708-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="5d708-111">Le tableau suivant décrit les autorisations déléguées prises en charge par l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="5d708-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="5d708-112">Autorisation</span><span class="sxs-lookup"><span data-stu-id="5d708-112">Permission</span></span>|<span data-ttu-id="5d708-113">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d708-113">Value</span></span>|<span data-ttu-id="5d708-114">Description</span><span class="sxs-lookup"><span data-stu-id="5d708-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="5d708-115">Lecture</span><span class="sxs-lookup"><span data-stu-id="5d708-115">Read</span></span>|<span data-ttu-id="5d708-116">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="5d708-116">1 (000001)</span></span>|<span data-ttu-id="5d708-117">Peut lire des éléments.</span><span class="sxs-lookup"><span data-stu-id="5d708-117">Can read items.</span></span>|
|<span data-ttu-id="5d708-118">Write</span><span class="sxs-lookup"><span data-stu-id="5d708-118">Write</span></span>|<span data-ttu-id="5d708-119">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="5d708-119">2 (000010)</span></span>|<span data-ttu-id="5d708-120">Peut créer des éléments.</span><span class="sxs-lookup"><span data-stu-id="5d708-120">Can create items.</span></span>|
|<span data-ttu-id="5d708-121">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="5d708-121">DeleteOwn</span></span>|<span data-ttu-id="5d708-122">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="5d708-122">4 (000100)</span></span>|<span data-ttu-id="5d708-123">Peut uniquement supprimer les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="5d708-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="5d708-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="5d708-124">DeleteAll</span></span>|<span data-ttu-id="5d708-125">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="5d708-125">8 (001000)</span></span>|<span data-ttu-id="5d708-126">Peut supprimer tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="5d708-126">Can delete any items.</span></span>|
|<span data-ttu-id="5d708-127">EditOwn</span><span class="sxs-lookup"><span data-stu-id="5d708-127">EditOwn</span></span>|<span data-ttu-id="5d708-128">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="5d708-128">16 (010000)</span></span>|<span data-ttu-id="5d708-129">Ne peut modifier que les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="5d708-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="5d708-130">EditAll</span><span class="sxs-lookup"><span data-stu-id="5d708-130">EditAll</span></span>|<span data-ttu-id="5d708-131">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="5d708-131">32 (100000)</span></span>|<span data-ttu-id="5d708-132">Peut modifier tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="5d708-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="5d708-133">Actuellement, l’API prend en charge l’obtention des autorisations de délégué existantes, mais pas la définition des autorisations de délégué.</span><span class="sxs-lookup"><span data-stu-id="5d708-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="5d708-134">L’objet [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) est implémenté à l’aide d’un masque de masques pour indiquer les autorisations du délégué.</span><span class="sxs-lookup"><span data-stu-id="5d708-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="5d708-135">Chaque position dans le masque de données représente une autorisation particulière et si elle est `1` définie sur Then, le délégué dispose de l’autorisation correspondante.</span><span class="sxs-lookup"><span data-stu-id="5d708-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="5d708-136">Par exemple, si le deuxième bit à partir de la `1`droite est, le délégué dispose alors d’une autorisation en **écriture** .</span><span class="sxs-lookup"><span data-stu-id="5d708-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="5d708-137">Vous pouvez voir un exemple sur la façon de vérifier une autorisation spécifique dans la section [effectuer une opération en tant que délégué](#perform-an-operation-as-delegate) plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="5d708-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="5d708-138">Synchronisation entre les clients de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d708-138">Sync across mailbox clients</span></span>

<span data-ttu-id="5d708-139">Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées entre les boîtes aux lettres immédiatement.</span><span class="sxs-lookup"><span data-stu-id="5d708-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="5d708-140">Toutefois, si le complément utilise les opérations REST ou EWS pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons d’utiliser à la place l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter ce délai.</span><span class="sxs-lookup"><span data-stu-id="5d708-140">However, if the add-in uses REST or EWS operations to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="5d708-141">Pour plus d’informations, reportez-vous à la [section Propriétés personnalisées](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) de l’article « obtenir et définir des métadonnées dans un complément Outlook ».</span><span class="sxs-lookup"><span data-stu-id="5d708-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="5d708-142">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="5d708-142">Configure the manifest</span></span>

<span data-ttu-id="5d708-143">Pour activer les scénarios d’accès délégué dans votre complément, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` dans le manifeste sous l’élément `DesktopFormFactor`parent.</span><span class="sxs-lookup"><span data-stu-id="5d708-143">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="5d708-144">Actuellement, les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="5d708-144">At present, other form factors are not supported.</span></span>

<span data-ttu-id="5d708-145">L’exemple suivant montre l' `SupportsSharedFolders` élément défini `true` dans une section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="5d708-145">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="5d708-146">Effectuer une opération en tant que délégué</span><span class="sxs-lookup"><span data-stu-id="5d708-146">Perform an operation as delegate</span></span>

<span data-ttu-id="5d708-147">Vous pouvez obtenir les propriétés partagées d’un élément en mode de composition ou de lecture en appelant la méthode [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="5d708-147">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="5d708-148">Cela renvoie un objet [SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations du délégué, l’adresse de messagerie du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.</span><span class="sxs-lookup"><span data-stu-id="5d708-148">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="5d708-149">L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué dispose d’une autorisation en **écriture** et passer un appel Rest.</span><span class="sxs-lookup"><span data-stu-id="5d708-149">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5d708-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5d708-150">See also</span></span>

- [<span data-ttu-id="5d708-151">Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier</span><span class="sxs-lookup"><span data-stu-id="5d708-151">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="5d708-152">Partage de calendriers dans Office 365</span><span class="sxs-lookup"><span data-stu-id="5d708-152">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="5d708-153">Procédure de tri des éléments de manifeste</span><span class="sxs-lookup"><span data-stu-id="5d708-153">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="5d708-154">[Mask (Computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="5d708-154">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="5d708-155">Opérateurs de bits JavaScript</span><span class="sxs-lookup"><span data-stu-id="5d708-155">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)