---
title: Activer les dossiers partagés et les scénarios de boîtes aux lettres partagées dans un Outlook de messagerie
description: Explique comment configurer la prise en charge de la prise en charge des dossiers partagés (c’est-à-dire. accès délégué) et boîtes aux lettres partagées.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 70578f2c78a9dd88efc9ba70d5599a13e121df53
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290711"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="69bf0-104">Activer les dossiers partagés et les scénarios de boîtes aux lettres partagées dans un Outlook de messagerie</span><span class="sxs-lookup"><span data-stu-id="69bf0-104">Enable shared folders and shared mailbox scenarios in an Outlook add-in</span></span>

<span data-ttu-id="69bf0-105">Cet article explique comment activer les scénarios de dossiers partagés [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)(également appelés accès délégué) et de boîtes aux lettres partagées (désormais en prévisualisation) dans votre application Outlook, y compris les autorisations que l’API JavaScript Office prend en charge.</span><span class="sxs-lookup"><span data-stu-id="69bf0-105">This article describes how to enable shared folders (also known as delegate access) and shared mailbox (now in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69bf0-106">La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="69bf0-106">Support for this feature was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="69bf0-107">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="69bf0-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-setups"></a><span data-ttu-id="69bf0-108">Configurations prise en charge</span><span class="sxs-lookup"><span data-stu-id="69bf0-108">Supported setups</span></span>

<span data-ttu-id="69bf0-109">Les sections suivantes décrivent les configurations prise en charge pour les boîtes aux lettres partagées (désormais en prévisualisation) et les dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="69bf0-109">The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders.</span></span> <span data-ttu-id="69bf0-110">Les API de fonctionnalité peuvent ne pas fonctionner comme prévu dans d’autres configurations.</span><span class="sxs-lookup"><span data-stu-id="69bf0-110">The feature APIs may not work as expected in other configurations.</span></span> <span data-ttu-id="69bf0-111">Sélectionnez la plateforme que vous souhaitez apprendre à configurer.</span><span class="sxs-lookup"><span data-stu-id="69bf0-111">Select the platform you'd like to learn how to configure.</span></span>

### <a name="windows"></a>[<span data-ttu-id="69bf0-112">Windows</span><span class="sxs-lookup"><span data-stu-id="69bf0-112">Windows</span></span>](#tab/windows)

#### <a name="shared-folders"></a><span data-ttu-id="69bf0-113">Dossiers partagés</span><span class="sxs-lookup"><span data-stu-id="69bf0-113">Shared folders</span></span>

<span data-ttu-id="69bf0-114">Le propriétaire de la boîte aux lettres [doit d’abord fournir l’accès à un délégué.](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)</span><span class="sxs-lookup"><span data-stu-id="69bf0-114">The mailbox owner must first [provide access to a delegate](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="69bf0-115">Le délégué doit ensuite suivre les instructions décrites dans la section « Ajouter la boîte aux lettres d’une autre personne à votre profil » de l’article Gérer les éléments de courrier et de calendrier [d’une autre personne.](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)</span><span class="sxs-lookup"><span data-stu-id="69bf0-115">The delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="69bf0-116">Boîtes aux lettres partagées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="69bf0-116">Shared mailboxes (preview)</span></span>

<span data-ttu-id="69bf0-117">Exchange administrateurs de serveur peuvent créer et gérer des boîtes aux lettres partagées pour des ensembles d’utilisateurs à accéder.</span><span class="sxs-lookup"><span data-stu-id="69bf0-117">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="69bf0-118">Actuellement, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) est la seule version de serveur prise en charge pour cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="69bf0-118">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="69bf0-119">Une fonctionnalité Exchange Server appelée « mappage automatique » est mise en [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) place par défaut, ce qui signifie que la boîte aux lettres partagée doit apparaître automatiquement dans l’application Outlook d’un utilisateur après la fermeture et la réouverture de Outlook.</span><span class="sxs-lookup"><span data-stu-id="69bf0-119">An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened.</span></span> <span data-ttu-id="69bf0-120">Toutefois, si un administrateur a désactivé le mappage automatique, l’utilisateur doit suivre les étapes manuelles décrites dans la section « Ajouter une boîte aux lettres partagée à Outlook » de l’article Ouvrir et utiliser une boîte aux lettres partagée dans [Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span><span class="sxs-lookup"><span data-stu-id="69bf0-120">However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span></span>

> [!WARNING]
> <span data-ttu-id="69bf0-121">Ne **vous connectez** PAS à la boîte aux lettres partagée avec un mot de passe.</span><span class="sxs-lookup"><span data-stu-id="69bf0-121">Do **NOT** sign into the shared mailbox with a password.</span></span> <span data-ttu-id="69bf0-122">Les API de fonctionnalité ne fonctionneront pas dans ce cas.</span><span class="sxs-lookup"><span data-stu-id="69bf0-122">The feature APIs won't work in that case.</span></span>

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="69bf0-123">Navigateur web – Outlook moderne</span><span class="sxs-lookup"><span data-stu-id="69bf0-123">Web browser - modern Outlook</span></span>](#tab/modern)

#### <a name="shared-folders"></a><span data-ttu-id="69bf0-124">Dossiers partagés</span><span class="sxs-lookup"><span data-stu-id="69bf0-124">Shared folders</span></span>

<span data-ttu-id="69bf0-125">Le propriétaire de la boîte aux lettres doit [d’abord fournir l’accès à un délégué](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) en mettant à jour les autorisations du dossier de boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="69bf0-125">The mailbox owner must first [provide access to a delegate](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) by updating the mailbox folder permissions.</span></span> <span data-ttu-id="69bf0-126">Le délégué doit ensuite suivre les instructions décrites dans la section « Ajouter la boîte aux lettres d’une autre personne à votre liste de dossiers dans Outlook Web App » de l’article Accéder à la boîte aux lettres [d’une](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081)autre personne.</span><span class="sxs-lookup"><span data-stu-id="69bf0-126">The delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="69bf0-127">Boîtes aux lettres partagées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="69bf0-127">Shared mailboxes (preview)</span></span>

<span data-ttu-id="69bf0-128">Exchange administrateurs de serveur peuvent créer et gérer des boîtes aux lettres partagées pour des ensembles d’utilisateurs à accéder.</span><span class="sxs-lookup"><span data-stu-id="69bf0-128">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="69bf0-129">Actuellement, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) est la seule version de serveur prise en charge pour cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="69bf0-129">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="69bf0-130">Après avoir reçu l’accès, un utilisateur de boîte aux lettres partagée doit suivre les étapes décrites dans la section « Ajouter la boîte aux lettres partagée afin qu’elle s’affiche sous votre boîte aux lettres principale » de l’article Ouvrir et utiliser une boîte aux lettres partagée dans [Outlook sur le web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span><span class="sxs-lookup"><span data-stu-id="69bf0-130">After receiving access, a shared mailbox user must follow the steps outlined in the "Add the shared mailbox so it displays under your primary mailbox" section of the article [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span></span>

> [!WARNING]
> <span data-ttu-id="69bf0-131">**N’utilisez PAS** d’autres options telles que « Ouvrir une autre boîte aux lettres ».</span><span class="sxs-lookup"><span data-stu-id="69bf0-131">Do **NOT** use other options like "Open another mailbox".</span></span> <span data-ttu-id="69bf0-132">Il se peut que les API de fonctionnalité ne fonctionnent pas correctement.</span><span class="sxs-lookup"><span data-stu-id="69bf0-132">The feature APIs may not work properly then.</span></span>

---

<span data-ttu-id="69bf0-133">Pour en savoir plus sur l’endroit où les modules sont activés et non activés en général, reportez-vous à la section Éléments de boîte aux lettres disponibles pour les [add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) de la page de vue d’ensemble des Outlook.</span><span class="sxs-lookup"><span data-stu-id="69bf0-133">To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.</span></span>

## <a name="supported-permissions"></a><span data-ttu-id="69bf0-134">Autorisations prise en charge</span><span class="sxs-lookup"><span data-stu-id="69bf0-134">Supported permissions</span></span>

<span data-ttu-id="69bf0-135">Le tableau suivant décrit les autorisations que l’API JavaScript Office prend en charge pour les délégués et les utilisateurs de boîtes aux lettres partagées.</span><span class="sxs-lookup"><span data-stu-id="69bf0-135">The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.</span></span>

|<span data-ttu-id="69bf0-136">Autorisation</span><span class="sxs-lookup"><span data-stu-id="69bf0-136">Permission</span></span>|<span data-ttu-id="69bf0-137">Valeur</span><span class="sxs-lookup"><span data-stu-id="69bf0-137">Value</span></span>|<span data-ttu-id="69bf0-138">Description</span><span class="sxs-lookup"><span data-stu-id="69bf0-138">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="69bf0-139">Lecture</span><span class="sxs-lookup"><span data-stu-id="69bf0-139">Read</span></span>|<span data-ttu-id="69bf0-140">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="69bf0-140">1 (000001)</span></span>|<span data-ttu-id="69bf0-141">Peut lire des éléments.</span><span class="sxs-lookup"><span data-stu-id="69bf0-141">Can read items.</span></span>|
|<span data-ttu-id="69bf0-142">Write</span><span class="sxs-lookup"><span data-stu-id="69bf0-142">Write</span></span>|<span data-ttu-id="69bf0-143">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="69bf0-143">2 (000010)</span></span>|<span data-ttu-id="69bf0-144">Peut créer des éléments.</span><span class="sxs-lookup"><span data-stu-id="69bf0-144">Can create items.</span></span>|
|<span data-ttu-id="69bf0-145">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="69bf0-145">DeleteOwn</span></span>|<span data-ttu-id="69bf0-146">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="69bf0-146">4 (000100)</span></span>|<span data-ttu-id="69bf0-147">Peut supprimer uniquement les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="69bf0-147">Can delete only the items they created.</span></span>|
|<span data-ttu-id="69bf0-148">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="69bf0-148">DeleteAll</span></span>|<span data-ttu-id="69bf0-149">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="69bf0-149">8 (001000)</span></span>|<span data-ttu-id="69bf0-150">Peut supprimer tous les éléments.</span><span class="sxs-lookup"><span data-stu-id="69bf0-150">Can delete any items.</span></span>|
|<span data-ttu-id="69bf0-151">EditOwn</span><span class="sxs-lookup"><span data-stu-id="69bf0-151">EditOwn</span></span>|<span data-ttu-id="69bf0-152">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="69bf0-152">16 (010000)</span></span>|<span data-ttu-id="69bf0-153">Peut modifier uniquement les éléments qu’ils ont créés.</span><span class="sxs-lookup"><span data-stu-id="69bf0-153">Can edit only the items they created.</span></span>|
|<span data-ttu-id="69bf0-154">EditAll</span><span class="sxs-lookup"><span data-stu-id="69bf0-154">EditAll</span></span>|<span data-ttu-id="69bf0-155">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="69bf0-155">32 (100000)</span></span>|<span data-ttu-id="69bf0-156">Peut modifier n’importe quel objet.</span><span class="sxs-lookup"><span data-stu-id="69bf0-156">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="69bf0-157">Actuellement, l’API prend en charge l’obtention d’autorisations existantes, mais pas la définition d’autorisations.</span><span class="sxs-lookup"><span data-stu-id="69bf0-157">Currently the API supports getting existing permissions, but not setting permissions.</span></span>

<span data-ttu-id="69bf0-158">[L’objet DelegatePermissions est](/javascript/api/outlook/office.mailboxenums.delegatepermissions) implémenté à l’aide d’un masque de bits pour indiquer les autorisations.</span><span class="sxs-lookup"><span data-stu-id="69bf0-158">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions.</span></span> <span data-ttu-id="69bf0-159">Chaque position dans le masque de bits représente une autorisation particulière et si elle est définie sur, l’utilisateur dispose de `1` l’autorisation respective.</span><span class="sxs-lookup"><span data-stu-id="69bf0-159">Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission.</span></span> <span data-ttu-id="69bf0-160">Par exemple, si le deuxième bit à partir de la droite est `1` , l’utilisateur dispose de **l’autorisation d’écriture.**</span><span class="sxs-lookup"><span data-stu-id="69bf0-160">For example, if the second bit from the right is `1`, then the user has **Write** permission.</span></span> <span data-ttu-id="69bf0-161">Vous pouvez voir un exemple de vérification d’une autorisation spécifique dans la [section](#perform-an-operation-as-delegate-or-shared-mailbox-user) Effectuer une opération en tant que délégué ou utilisateur de boîte aux lettres partagée plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="69bf0-161">You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.</span></span>

## <a name="sync-across-shared-folder-clients"></a><span data-ttu-id="69bf0-162">Synchronisation entre les clients de dossiers partagés</span><span class="sxs-lookup"><span data-stu-id="69bf0-162">Sync across shared folder clients</span></span>

<span data-ttu-id="69bf0-163">Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées immédiatement entre les boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="69bf0-163">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="69bf0-164">Toutefois, si des opérations REST ou Exchange Web Services (EWS) ont été utilisées pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons plutôt d’utiliser [l’objet CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter un tel délai.</span><span class="sxs-lookup"><span data-stu-id="69bf0-164">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="69bf0-165">Pour en savoir plus, consultez la [section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) des propriétés personnalisées de l’article « Obtenir et définir des métadonnées dans un Outlook de données ».</span><span class="sxs-lookup"><span data-stu-id="69bf0-165">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69bf0-166">Dans un scénario de délégué, vous ne pouvez pas utiliser EWS avec les jetons actuellement fournis par office.js API.</span><span class="sxs-lookup"><span data-stu-id="69bf0-166">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="69bf0-167">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="69bf0-167">Configure the manifest</span></span>

<span data-ttu-id="69bf0-168">Pour activer les dossiers partagés et les scénarios de boîtes aux lettres partagées dans votre add-in, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) dans le manifeste sous `true` l’élément `DesktopFormFactor` parent.</span><span class="sxs-lookup"><span data-stu-id="69bf0-168">To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="69bf0-169">Pour l’instant, les autres facteurs de forme ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="69bf0-169">At present, other form factors are not supported.</span></span>

<span data-ttu-id="69bf0-170">Pour prendre en charge les appels REST d’un délégué, définissez le nœud [Autorisations](../reference/manifest/permissions.md) dans le manifeste sur `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="69bf0-170">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="69bf0-171">L’exemple suivant montre `SupportsSharedFolders` l’ensemble `true` d’éléments dans une section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="69bf0-171">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a><span data-ttu-id="69bf0-172">Effectuer une opération en tant qu’utilisateur de boîte aux lettres déléguée ou partagée</span><span class="sxs-lookup"><span data-stu-id="69bf0-172">Perform an operation as delegate or shared mailbox user</span></span>

<span data-ttu-id="69bf0-173">Vous pouvez obtenir les propriétés partagées d’un élément en mode Composition ou Lecture en appelant la méthode [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="69bf0-173">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="69bf0-174">Cela renvoie un [objet SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations de l’utilisateur, l’adresse e-mail du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.</span><span class="sxs-lookup"><span data-stu-id="69bf0-174">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="69bf0-175">L’exemple suivant montre comment obtenir les propriétés partagées d’un message  ou d’un rendez-vous, vérifier si le délégué ou l’utilisateur de boîte aux lettres partagée dispose d’une autorisation d’écriture et passer un appel REST.</span><span class="sxs-lookup"><span data-stu-id="69bf0-175">The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="69bf0-176">En tant que délégué, vous pouvez utiliser REST pour obtenir le contenu d’un message Outlook joint à un élément Outlook ou un [billet de groupe.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="69bf0-176">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="69bf0-177">Gérer l’appel de REST sur les éléments partagés et non partagés</span><span class="sxs-lookup"><span data-stu-id="69bf0-177">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="69bf0-178">Si vous souhaitez appeler une opération REST sur un élément, que l’élément soit partagé ou non, vous pouvez utiliser l’API pour déterminer si l’élément `getSharedPropertiesAsync` est partagé.</span><span class="sxs-lookup"><span data-stu-id="69bf0-178">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="69bf0-179">Après cela, vous pouvez construire l’URL REST pour l’opération à l’aide de l’objet approprié.</span><span class="sxs-lookup"><span data-stu-id="69bf0-179">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="69bf0-180">Limites</span><span class="sxs-lookup"><span data-stu-id="69bf0-180">Limitations</span></span>

<span data-ttu-id="69bf0-181">Selon les scénarios de votre add-in, il existe quelques limitations à prendre en compte lors de la gestion des situations de dossier partagé ou de boîte aux lettres partagée.</span><span class="sxs-lookup"><span data-stu-id="69bf0-181">Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="69bf0-182">Mode composition de message</span><span class="sxs-lookup"><span data-stu-id="69bf0-182">Message Compose mode</span></span>

<span data-ttu-id="69bf0-183">En mode composition de message, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) n’est pas pris en charge dans Outlook sur le web ou sur Windows à moins que les conditions suivantes ne soient remplies.</span><span class="sxs-lookup"><span data-stu-id="69bf0-183">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.</span></span>

<span data-ttu-id="69bf0-184">a.</span><span class="sxs-lookup"><span data-stu-id="69bf0-184">a.</span></span> <span data-ttu-id="69bf0-185">**Accès délégué/Dossiers partagés**</span><span class="sxs-lookup"><span data-stu-id="69bf0-185">**Delegate access/Shared folders**</span></span>

1. <span data-ttu-id="69bf0-186">Le propriétaire de la boîte aux lettres démarre un message.</span><span class="sxs-lookup"><span data-stu-id="69bf0-186">The mailbox owner starts a message.</span></span> <span data-ttu-id="69bf0-187">Il peut s’agit d’un nouveau message, d’une réponse ou d’un forward.</span><span class="sxs-lookup"><span data-stu-id="69bf0-187">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="69bf0-188">Ils enregistrent le message, puis le déplacent de leur propre dossier **Brouillons** vers un dossier partagé avec le délégué.</span><span class="sxs-lookup"><span data-stu-id="69bf0-188">They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.</span></span>
1. <span data-ttu-id="69bf0-189">Le délégué ouvre le brouillon à partir du dossier partagé, puis continue la composition.</span><span class="sxs-lookup"><span data-stu-id="69bf0-189">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="69bf0-190">b.</span><span class="sxs-lookup"><span data-stu-id="69bf0-190">b.</span></span> <span data-ttu-id="69bf0-191">**Boîte aux lettres partagée**</span><span class="sxs-lookup"><span data-stu-id="69bf0-191">**Shared mailbox**</span></span>

1. <span data-ttu-id="69bf0-192">Un utilisateur de boîte aux lettres partagée démarre un message.</span><span class="sxs-lookup"><span data-stu-id="69bf0-192">A shared mailbox user starts a message.</span></span> <span data-ttu-id="69bf0-193">Il peut s’agit d’un nouveau message, d’une réponse ou d’un forward.</span><span class="sxs-lookup"><span data-stu-id="69bf0-193">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="69bf0-194">Ils enregistrent le message, puis le déplacent de leur propre dossier **Brouillons** vers un dossier de la boîte aux lettres partagée.</span><span class="sxs-lookup"><span data-stu-id="69bf0-194">They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.</span></span>
1. <span data-ttu-id="69bf0-195">Un autre utilisateur de boîte aux lettres partagée ouvre le brouillon à partir de la boîte aux lettres partagée, puis continue la composition.</span><span class="sxs-lookup"><span data-stu-id="69bf0-195">Another shared mailbox user opens the draft from the shared mailbox then continues composing.</span></span>

<span data-ttu-id="69bf0-196">Le message se trouve maintenant dans un contexte partagé et les modules qui la prisent en charge de ces scénarios partagés peuvent obtenir les propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="69bf0-196">The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.</span></span> <span data-ttu-id="69bf0-197">Une fois le message envoyé, il se trouve généralement  dans le dossier Éléments envoyés de l’expéditeur.</span><span class="sxs-lookup"><span data-stu-id="69bf0-197">After the message has been sent, it's usually found in the sender's **Sent Items** folder.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="69bf0-198">REST et EWS</span><span class="sxs-lookup"><span data-stu-id="69bf0-198">REST and EWS</span></span>

<span data-ttu-id="69bf0-199">Votre application peut utiliser REST et son autorisation doit être définie pour activer l’accès REST à la boîte aux lettres du propriétaire ou à la boîte aux lettres partagée, le `ReadWriteMailbox` cas échéant.</span><span class="sxs-lookup"><span data-stu-id="69bf0-199">Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox or to the shared mailbox as applicable.</span></span> <span data-ttu-id="69bf0-200">EWS n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="69bf0-200">EWS is not supported.</span></span>

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a><span data-ttu-id="69bf0-201">Boîte aux lettres utilisateur ou partagée masquée dans une liste d’adresses</span><span class="sxs-lookup"><span data-stu-id="69bf0-201">User or shared mailbox hidden from an address list</span></span>

<span data-ttu-id="69bf0-202">Si un administrateur a caché un utilisateur ou une adresse de boîte aux lettres partagée à partir d’une liste d’adresses telle que la liste d’adresses globale ,les éléments de courrier affectés ouverts dans le rapport de boîte aux lettres sont `Office.context.mailbox.item` null.</span><span class="sxs-lookup"><span data-stu-id="69bf0-202">If an admin hid a user or shared mailbox address from an address list like the global address list (GAL), affected mail items opened in the mailbox report `Office.context.mailbox.item` as null.</span></span> <span data-ttu-id="69bf0-203">Par exemple, si l’utilisateur ouvre un élément de courrier dans une boîte aux lettres partagée qui est masquée dans la liste d’adresses gal, représentant cet élément de `Office.context.mailbox.item` courrier est null.</span><span class="sxs-lookup"><span data-stu-id="69bf0-203">For example, if the user opens a mail item in a shared mailbox that's hidden from the GAL, `Office.context.mailbox.item` representing that mail item is null.</span></span>

## <a name="see-also"></a><span data-ttu-id="69bf0-204">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="69bf0-204">See also</span></span>

- [<span data-ttu-id="69bf0-205">Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier</span><span class="sxs-lookup"><span data-stu-id="69bf0-205">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="69bf0-206">Partage de calendrier dans Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="69bf0-206">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="69bf0-207">Ajouter une boîte aux lettres partagée à Outlook</span><span class="sxs-lookup"><span data-stu-id="69bf0-207">Add a shared mailbox to Outlook</span></span>](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [<span data-ttu-id="69bf0-208">Comment commander des éléments de manifeste</span><span class="sxs-lookup"><span data-stu-id="69bf0-208">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="69bf0-209">[Masque (calcul)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="69bf0-209">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="69bf0-210">Opérateurs de bits JavaScript</span><span class="sxs-lookup"><span data-stu-id="69bf0-210">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)