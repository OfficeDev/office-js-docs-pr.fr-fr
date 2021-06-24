---
title: Confidentialité, autorisations et sécurité pour les compléments Outlook
description: Découvrez comment gérer la confidentialité, les autorisations et la sécurité dans un complément Outlook.
ms.date: 04/07/2021
localization_priority: Priority
ms.openlocfilehash: 1c8c5420593b31f403cf8f5fa28659fc130db402
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076993"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a><span data-ttu-id="2d0a9-103">Confidentialité, autorisations et sécurité pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2d0a9-103">Privacy, permissions, and security for Outlook add-ins</span></span>

<span data-ttu-id="2d0a9-104">Les utilisateurs finaux, les développeurs et les administrateurs peuvent appliquer les niveaux d’autorisation hiérarchisés du modèle de sécurité pour les compléments Outlook afin de contrôler les performances et la confidentialité.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-104">End users, developers, and administrators can use the tiered permission levels of the security model for Outlook add-ins to control privacy and performance.</span></span>

<span data-ttu-id="2d0a9-105">Cet article décrit les autorisations que les compléments Outlook peuvent demander, et examine le modèle de sécurité selon les perspectives suivantes.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-105">This article describes the possible permissions that Outlook add-ins can request, and examines the security model from the following perspectives.</span></span>

- <span data-ttu-id="2d0a9-106">**AppSource** : intégrité de complément</span><span class="sxs-lookup"><span data-stu-id="2d0a9-106">**AppSource**: Add-in integrity</span></span>

- <span data-ttu-id="2d0a9-107">**Utilisateurs** : problèmes de confidentialité et de performance</span><span class="sxs-lookup"><span data-stu-id="2d0a9-107">**End-users**: Privacy and performance concerns</span></span>

- <span data-ttu-id="2d0a9-108">**Développeurs** : choix d’autorisations et limites d’utilisation des ressources</span><span class="sxs-lookup"><span data-stu-id="2d0a9-108">**Developers**: Permissions choices and resource usage limits</span></span>

- <span data-ttu-id="2d0a9-109">**Administrateurs**: privilèges pour définir des seuils de performances</span><span class="sxs-lookup"><span data-stu-id="2d0a9-109">**Administrators**: Privileges to set performance thresholds</span></span>

## <a name="permissions-model"></a><span data-ttu-id="2d0a9-110">Modèle d’autorisations</span><span class="sxs-lookup"><span data-stu-id="2d0a9-110">Permissions model</span></span>

<span data-ttu-id="2d0a9-p101">Comme la façon dont les clients perçoivent la sécurité des compléments peut avoir une incidence sur l’adoption de ces derniers, la sécurité des compléments Outlook repose sur un modèle d’autorisations à plusieurs niveaux. Un complément Outlook indique le niveau d’autorisations dont il a besoin, identifiant ainsi l’accès dont il peut disposer et les actions qu’il peut effectuer sur les données de la boîte aux lettres du client.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p101">Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.</span></span>

<span data-ttu-id="2d0a9-113">Le schéma de manifeste version 1.1 comprend quatre niveaux d’autorisation.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-113">Manifest schema version 1.1 includes four levels of permissions.</span></span>

<span data-ttu-id="2d0a9-114">**Tableau 1. Niveaux d’autorisation d’un complément**</span><span class="sxs-lookup"><span data-stu-id="2d0a9-114">**Table 1. Add-in permission levels**</span></span>

|<span data-ttu-id="2d0a9-115">**Niveau d’autorisation**</span><span class="sxs-lookup"><span data-stu-id="2d0a9-115">**Permission level**</span></span>|<span data-ttu-id="2d0a9-116">**Valeur dans le manifeste du complément Outlook**</span><span class="sxs-lookup"><span data-stu-id="2d0a9-116">**Value in Outlook add-in manifest**</span></span>|
|:-----|:-----|
|<span data-ttu-id="2d0a9-117">Restricted</span><span class="sxs-lookup"><span data-stu-id="2d0a9-117">Restricted</span></span>|<span data-ttu-id="2d0a9-118">Restreint</span><span class="sxs-lookup"><span data-stu-id="2d0a9-118">Restricted</span></span>|
|<span data-ttu-id="2d0a9-119">Lire l’élément</span><span class="sxs-lookup"><span data-stu-id="2d0a9-119">Read item</span></span>|<span data-ttu-id="2d0a9-120">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d0a9-120">ReadItem</span></span>|
|<span data-ttu-id="2d0a9-121">Lire/écrire dans l’élément</span><span class="sxs-lookup"><span data-stu-id="2d0a9-121">Read/write item</span></span>|<span data-ttu-id="2d0a9-122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2d0a9-122">ReadWriteItem</span></span>|
|<span data-ttu-id="2d0a9-123">Lire/écrire dans la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d0a9-123">Read/write mailbox</span></span>|<span data-ttu-id="2d0a9-124">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2d0a9-124">ReadWriteMailbox</span></span>|

<span data-ttu-id="2d0a9-125">Les quatre niveaux d’autorisations sont cumulatifs : l’autorisation **boîte aux lettres en lecture/écriture** inclut les autorisations de **élément en lecture/écriture**, **lire élément** et **restreint**, l’autorisation **élément en lecture/écriture** inclut **lire élément** et **restreint** et l’autorisation **lire élément** inclut **restreint**.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-125">The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.</span></span>

<span data-ttu-id="2d0a9-126">L’illustration suivante affiche les quatre niveaux d’autorisations et décrit les fonctionnalités proposées aux utilisateurs finaux, développeur et administrateur par chaque niveau.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-126">The following figure shows the four levels of permissions and describes the capabilities offered to the end user, developer, and administrator by each tier.</span></span> <span data-ttu-id="2d0a9-127">Pour plus d’informations sur ces autorisations, voir [utilisateurs : problèmes de performances et de confidentialité](#end-users-privacy-and-performance-concerns), [développeurs : choix d’autorisation et les limites de l’utilisation de ressources](#developers-permission-choices-and-resource-usage-limits), et [comprendre les autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-127">For more information about these permissions, see [End users: privacy and performance concerns](#end-users-privacy-and-performance-concerns), [Developers: permission choices and resource usage limits](#developers-permission-choices-and-resource-usage-limits), and [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="2d0a9-128">**Association du modèle d’autorisation à quatre niveaux à l’utilisateur final, au développeur et à l’administrateur**</span><span class="sxs-lookup"><span data-stu-id="2d0a9-128">**Relating the four-tier permission model to the end user, developer, and administrator**</span></span>

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1.](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a><span data-ttu-id="2d0a9-130">AppSource : intégrité de complément</span><span class="sxs-lookup"><span data-stu-id="2d0a9-130">AppSource: Add-in integrity</span></span>

<span data-ttu-id="2d0a9-131">[AppSource](https://appsource.microsoft.com) héberge des compléments pouvant être installés par les utilisateurs finals et les administrateurs.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-131">[AppSource](https://appsource.microsoft.com) hosts add-ins that can be installed by end users and administrators.</span></span> <span data-ttu-id="2d0a9-132">AppSource applique les mesures suivantes pour maintenir l’intégrité de ces compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-132">AppSource enforces the following measures to maintain the integrity of these Outlook add-ins.</span></span>

- <span data-ttu-id="2d0a9-133">Oblige le serveur hôte d’un complément à toujours utiliser SSL (Secure Socket Layer) pour communiquer.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-133">Requires the host server of an add-in to always use Secure Socket Layer (SSL) to communicate.</span></span>

- <span data-ttu-id="2d0a9-134">Oblige un développeur à fournir une preuve d’identité, un accord contractuel et une politique de confidentialité conforme pour soumettre les compléments.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-134">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span>

- <span data-ttu-id="2d0a9-135">Archive les compléments en mode lecture seule.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-135">Archives add-ins in read-only mode.</span></span>

- <span data-ttu-id="2d0a9-136">Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-136">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>

## <a name="optional-connected-experiences"></a><span data-ttu-id="2d0a9-137">Expériences connectées facultatives</span><span class="sxs-lookup"><span data-stu-id="2d0a9-137">Optional connected experiences</span></span>

<span data-ttu-id="2d0a9-138">Les utilisateurs finaux et les administrateurs informatiques peuvent désactiver [expériences connectées facultatives dans ](/deployoffice/privacy/optional-connected-experiences) les clients de bureau et mobiles Office.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-138">End users and IT admins can turn off [optional connected experiences in Office](/deployoffice/privacy/optional-connected-experiences) desktop and mobile clients.</span></span> <span data-ttu-id="2d0a9-139">Pour les compléments Outlook, l’impact de la désactivation du paramètres **Expériences connectées optionnelles** dépend du client, mais les compléments installés par l’utilisateur et l’accès à Office Store ne sont généralement pas autorisés.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-139">For Outlook add-ins, the impact of disabling the **Optional connected experiences** setting depends on the client but usually means that user-installed add-ins and access to the Office Store are not allowed.</span></span> <span data-ttu-id="2d0a9-140">Certains compléments Microsoft sont considérés comme essentiels ou stratégiques, et les compléments déployés par l’administrateur informatique d’une organisation via [Déploiement centralisé](../publish/centralized-deployment.md) restent disponibles.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-140">Add-ins deployed by an organization's IT admin through [Centralized Deployment](../publish/centralized-deployment.md) will still be available.</span></span>

- <span data-ttu-id="2d0a9-141">Windows\*, Mac : le bouton **Obtenir des compléments** ne s’affiche pas afin que les utilisateurs ne puissent plus gérer leurs compléments ni accéder à Office Store.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-141">Windows\*, Mac: The **Get Add-ins** button is not displayed so users can no longer manage their add-ins or access the Office Store.</span></span>
- <span data-ttu-id="2d0a9-142">Android, iOS : la boîte de dialogue **Obtenir des compléments** affiche uniquement les compléments déployés par l’administrateur.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-142">Android, iOS: The **Get Add-ins** dialog shows only admin-deployed add-ins.</span></span>
- <span data-ttu-id="2d0a9-143">Navigateur : la disponibilité des compléments et l’accès au Store ne sont pas affectés de sorte que les utilisateurs puissent continuer à [gérer leurs compléments](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), y compris ceux déployés par l’administrateur.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-143">Browser: Availability of add-ins and access to the Store are unaffected so users can continue to [manage their add-ins](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), including admin-deployed ones.</span></span>

  > [!NOTE]
  > <span data-ttu-id="2d0a9-144">\* Pour Windows, la prise en charge de cette expérience/ce comportement est disponible à partir de la version 2008 (build 13127.20296).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-144">\* For Windows, support for this experience/behavior is available from version 2008 (build 13127.20296).</span></span> <span data-ttu-id="2d0a9-145">Pour plus d’informations en fonction de votre version, consultez la page de l’historique des mises à jour de [Miicrosoft 365](/officeupdates/update-history-office365-proplus-by-date) et [comment trouver la version du client et le canal de mise à jour Office que vous utilisez](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-145">For more details according to your version, see the update history page for [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).</span></span>

<span data-ttu-id="2d0a9-146">Pour obtenir des informations générales sur le comportement des compléments, consultez [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md#optional-connected-experiences).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-146">For general add-in behavior, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md#optional-connected-experiences).</span></span>

## <a name="end-users-privacy-and-performance-concerns"></a><span data-ttu-id="2d0a9-147">Utilisateurs : problèmes de confidentialité et de performance</span><span class="sxs-lookup"><span data-stu-id="2d0a9-147">End users: Privacy and performance concerns</span></span>

<span data-ttu-id="2d0a9-148">Le modèle de sécurité résout les problèmes de sécurité, de confidentialité et de performance des utilisateurs des manières suivantes.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-148">The security model addresses security, privacy, and performance concerns of end users in the following ways.</span></span>

- <span data-ttu-id="2d0a9-149">Les messages des utilisateurs qui sont protégés par la Gestion des droits relatifs à l’information (IRM) d’Outlook n’ont pas d’interaction avec les compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-149">End user's messages that are protected by Outlook's Information Rights Management (IRM) do not interact with Outlook add-ins.</span></span>

  > [!IMPORTANT]
  > - <span data-ttu-id="2d0a9-150">Les compléments s’activent sur les messages signés numériquement dans Outlook avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-150">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="2d0a9-151">Dans Windows, cette prise en charge a été introduite avec le build 8711.1000.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-151">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="2d0a9-152">Démarrer avec Outlook build 13229.10000 sur Windows, les compléments peuvent désormais activer les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-152">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="2d0a9-153">Pour plus d’informations sur cette fonctionnalité en mode aperçu, voir [Activation de complément sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-153">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="2d0a9-p108">Avant d’installer un complément de AppSource, les utilisateurs finals peuvent voir l’accès dont peut disposer le complément, ainsi que les actions qu’il peut effectuer sur leurs données, et doivent explicitement confirmer qu’ils veulent poursuivre. Aucun complément Outlook n’est automatiquement transmis sur un ordinateur client sans une validation manuelle par l’utilisateur ou l’administrateur.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p108">Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed. No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.</span></span>

- <span data-ttu-id="2d0a9-p109">L’octroi de l’autorisation **Restreint** permet au complément Outlook d’avoir un accès limité uniquement sur l’élément actuel. L’octroi de l’autorisation **Lire l’élément** permet au complément Outlook d’accéder à des informations d’identification personnelle, par exemple les noms et les adresses électroniques des expéditeurs et des destinataires, uniquement sur l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p109">Granting the **restricted** permission allows the Outlook add-in to have limited access on only the current item. Granting the **read item** permission allows the Outlook add-in to access personal identifiable information, such as sender and recipient names and email addresses, on only the current item,.</span></span>

- <span data-ttu-id="2d0a9-p110">Un utilisateur final peut installer un complément Outlook uniquement pour lui-même. Les compléments de messagerie ayant une incidence sur l’organisation sont installés par un administrateur.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p110">An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.</span></span>

- <span data-ttu-id="2d0a9-160">Les utilisateurs peuvent installer des compléments Outlook qui activent des scénarios contextuels prisés par les utilisateurs tout en minimisant les risques de sécurité pour ces derniers.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-160">End users can install Outlook add-ins that enable context-sensitive scenarios that are compelling to users while minimizing the users' security risks.</span></span>

- <span data-ttu-id="2d0a9-161">Les fichiers manifeste de compléments Outlook installés sont sécurisés dans le compte de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-161">Manifest files of installed Outlook add-ins are secured in the user's email account.</span></span>

- <span data-ttu-id="2d0a9-162">Les données échangées avec des serveurs hébergeant des Compléments Office sont toujours chiffrées conformément au protocole SSL (Secure Socket Layer).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-162">Data communicated with servers hosting Office Add-ins is always encrypted according to the Secure Socket Layer (SSL) protocol.</span></span>

- <span data-ttu-id="2d0a9-163">Applicable uniquement aux clients riches Outlook : les clients riches Outlook surveillent la performance des compléments Outlook installés, exercent un contrôle de gouvernance et désactivent les compléments Outlook qui dépassent les limites pour les aspects suivants.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-163">Applicable to only the Outlook rich clients: The Outlook rich clients monitor the performance of installed Outlook add-ins, exercise governance control, and disable those Outlook add-ins that exceed limits in the following areas.</span></span>

  - <span data-ttu-id="2d0a9-164">Temps de réponse d’activation</span><span class="sxs-lookup"><span data-stu-id="2d0a9-164">Response time to activate</span></span>

  - <span data-ttu-id="2d0a9-165">Nombre de défaillances d’activation ou de réactivation</span><span class="sxs-lookup"><span data-stu-id="2d0a9-165">Number of failures to activate or reactivate</span></span>

  - <span data-ttu-id="2d0a9-166">Utilisation de la mémoire</span><span class="sxs-lookup"><span data-stu-id="2d0a9-166">Memory usage</span></span>

  - <span data-ttu-id="2d0a9-167">Utilisation du processeur</span><span class="sxs-lookup"><span data-stu-id="2d0a9-167">CPU usage</span></span>  

  <span data-ttu-id="2d0a9-p111">La gouvernance dissuade les attaques par déni de service et maintient les performances des compléments à un niveau raisonnable. La barre Entreprise indique aux utilisateurs les compléments Outlook que le client riche Outlook a désactivés sur la base d’un tel contrôle de gouvernance.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p111">Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.</span></span>

- <span data-ttu-id="2d0a9-170">À tout moment, les utilisateurs finals peuvent vérifier les autorisations demandées par les compléments Outlook installés, et désactiver ou activer ultérieurement tout complément Outlook dans le Centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-170">At any time, end users can verify the permissions requested by installed Outlook add-ins, and disable or subsequently enable any Outlook add-in in the Exchange Admin Center.</span></span>

## <a name="developers-permission-choices-and-resource-usage-limits"></a><span data-ttu-id="2d0a9-171">Développeurs : choix d’autorisations et limites d’utilisation des ressources.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-171">Developers: Permission choices and resource usage limits</span></span>

<span data-ttu-id="2d0a9-172">Le modèle de sécurité fournit aux développeurs des niveaux précis d’autorisations à choisir, et de strictes directives de performance à observer.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-172">The security model provides developers granular levels of permissions to choose from, and strict performance guidelines to observe.</span></span>

### <a name="tiered-permissions-increases-transparency"></a><span data-ttu-id="2d0a9-173">Les autorisations à plusieurs niveaux augmentent la transparence</span><span class="sxs-lookup"><span data-stu-id="2d0a9-173">Tiered permissions increases transparency</span></span>

<span data-ttu-id="2d0a9-174">Les développeurs doivent suivre le modèle d’autorisations à plusieurs niveaux pour assurer la transparence et apaiser les inquiétudes des utilisateurs concernant ce que les compléments peuvent faire à leurs données et leur boîte aux lettres, en faisant la promotion indirecte de l’adoption du complément.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-174">Developers should follow the tiered permissions model to provide transparency and alleviate users' concern about what add-ins can do to their data and mailbox, indirectly promoting add-in adoption.</span></span>

- <span data-ttu-id="2d0a9-175">Les développeurs demandent un niveau approprié d’autorisation pour un complément Outlook en fonction de la manière dont il doit être activé, et de son besoin de lire ou d’écrire certaines propriétés d’un élément, ou de créer et d’envoyer un élément.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-175">Developers request an appropriate level of permission for an Outlook add-in, based on how the Outlook add-in should be activated, and its need to read or write certain properties of an item, or to create and send an item.</span></span>

- <span data-ttu-id="2d0a9-176">Les développeurs demandent une autorisation en utilisant l’élément [Permissions](../reference/manifest/permissions.md) dans le manifeste du complément Outlook, en affectant une valeur **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**, selon le cas.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-176">Developers request permission by using the [Permissions](../reference/manifest/permissions.md) element in the manifest of the Outlook add-in, by assigning a value of **Restricted**, **ReadItem**, **ReadWriteItem** or **ReadWriteMailbox**, as appropriate.</span></span>

  > [!NOTE]
  > <span data-ttu-id="2d0a9-177">Notez que l’autorisation **ReadWriteItem** est disponible à partir du schéma de manifeste version 1.1.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-177">Note that the **ReadWriteItem** permission is available starting in manifest schema v1.1.</span></span>

  <span data-ttu-id="2d0a9-178">L’exemple suivant demande l’autorisation **Lire l’élément**.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-178">The following example requests the **read item** permission.</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- <span data-ttu-id="2d0a9-p112">Les développeurs peuvent demander l’autorisation **Restreint** si le complément Outlook est activé lorsqu’un type spécifique d’élément Outlook (rendez-vous ou message) ou des entités extraites spécifiques (numéro de téléphone, adresse, URL) sont présents dans l’objet ou le corps de l’élément. Par exemple, la règle suivante active le complément Outlook si une ou plusieurs des trois entités (numéro de téléphone, adresse postale ou URL) se trouvent dans l’objet ou le corps du message actuel.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p112">Developers can request the **restricted** permission if the Outlook add-in activates on a specific type of Outlook items (appointment or message), or on specific extracted entities (phone number, address, URL) being present in the item's subject or body. For example, the following rule activates the Outlook add-in if one or more of three entities - phone number, postal address, or URL - are found in the subject or body of the current message.</span></span>

  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- <span data-ttu-id="2d0a9-p113">Les développeurs doivent demander l’autorisation **Lire l’élément** si le complément Outlook doit lire les propriétés de l’élément actuel autres que les entités extraites par défaut, ou écrire des propriétés personnalisées définies par le complément sur l’élément actuel, mais ne nécessite pas de lire ou d’écrire d’autres éléments, ou de créer ou d’envoyer un message dans la boîte aux lettres de l’utilisateur. Par exemple, un développeur doit demander l’autorisation **Lire l’élément** si un complément Outlook doit rechercher une entité telle qu’une suggestion de réunion, une suggestion de tâche, une adresse électronique, ou un nom de contact dans l’objet ou le corps de l’élément, ou utilise une expression régulière pour s’activer.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p113">Developers should request the **read item** permission if the Outlook add-in needs to read properties of the current item other than the default extracted entities, or write custom properties set by the add-in on the current item, but does not require reading or writing to other items, or creating or sending a message in the user's mailbox. For example, a developer should request **read item** permission if an Outlook add-in needs to look for an entity like a meeting suggestion, task suggestion, email address, or contact name in the item's subject or body, or uses a regular expression to activate.</span></span>

- <span data-ttu-id="2d0a9-183">Les développeurs doivent demander l’autorisation **Lire/écrire dans l’élément** si le complément Outlook doit écrire dans les propriétés de l’élément composé, comme les noms des destinataires, les adresses de messagerie, le corps et l’objet, ou s’il a besoin d’ajouter ou de supprimer des pièces jointes d’élément.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-183">Developers should request the **read/write item** permission if the Outlook add-in needs to write to properties of the composed item, such as recipient names, email addresses, body, and subject, or needs to add or remove item attachments.</span></span>

- <span data-ttu-id="2d0a9-184">Les développeurs demandent l’autorisation **Lire/écrire dans la boîte aux lettres** uniquement si le complément Outlook doit effectuer une ou plusieurs des actions suivantes à l’aide de la méthode [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-184">Developers request the **read/write mailbox** permission only if the Outlook add-in needs to do one or more of the following actions by using the [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>

  - <span data-ttu-id="2d0a9-185">Lire ou écrire des propriétés d’éléments dans la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-185">Read or write to properties of items in the mailbox.</span></span>
  - <span data-ttu-id="2d0a9-186">Créer, lire, écrire ou envoyer des éléments dans la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-186">Create, read, write, or send items in the mailbox.</span></span>
  - <span data-ttu-id="2d0a9-187">Créer, lire ou écrire dans des dossiers de la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-187">Create, read, or write to folders in the mailbox.</span></span>

### <a name="resource-usage-tuning"></a><span data-ttu-id="2d0a9-188">Réglage de l’utilisation des ressources</span><span class="sxs-lookup"><span data-stu-id="2d0a9-188">Resource usage tuning</span></span>

<span data-ttu-id="2d0a9-p114">Les développeurs doivent connaître les limites de l’utilisation des ressources pour l’activation, incorporer le réglage des performances dans leur flux de travail de développement, afin de réduire le risque d’un complément peu performant refusant le service de l’hôte. Les développeurs doivent suivre les directives concernant la conception des règles d’activation telles que décrites dans [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Si un complément Outlook est destiné à être exécuté sur un client riche Outlook, les développeurs doivent vérifier que les performances du complément se situent dans les limites d’utilisation des ressources.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-p114">Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.</span></span>

### <a name="other-measures-to-promote-user-security"></a><span data-ttu-id="2d0a9-191">Autres mesures visant à promouvoir la sécurité de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="2d0a9-191">Other measures to promote user security</span></span>

<span data-ttu-id="2d0a9-192">Les développeurs doivent connaître et planifier les éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-192">Developers should be aware of and plan for the following as well.</span></span>

- <span data-ttu-id="2d0a9-193">Les développeurs ne peuvent pas utiliser de contrôles ActiveX dans les compléments car ils ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-193">Developers cannot use ActiveX controls in add-ins because they are not supported.</span></span>

- <span data-ttu-id="2d0a9-194">Les développeurs doivent procéder comme suit lorsqu’ils envoient un complément Outlook à AppSource.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-194">Developers should do the following when submitting an Outlook add-in to AppSource.</span></span>

  - <span data-ttu-id="2d0a9-195">Produire un certificat SSL EV (Extended Validation) comme preuve d’identité.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-195">Produce an Extended Validation (EV) SSL certificate as a proof of identity.</span></span>

  - <span data-ttu-id="2d0a9-196">Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-196">Host the add-in they are submitting on a web server that supports SSL.</span></span>

  - <span data-ttu-id="2d0a9-197">Produire une stratégie de confidentialité conforme.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-197">Produce a compliant privacy policy.</span></span>

  - <span data-ttu-id="2d0a9-198">Être prêts à signer un accord contractuel lors de la soumission du complément.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-198">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>

## <a name="administrators-privileges"></a><span data-ttu-id="2d0a9-199">Administrateurs : privilèges</span><span class="sxs-lookup"><span data-stu-id="2d0a9-199">Administrators: Privileges</span></span>

<span data-ttu-id="2d0a9-200">Le modèle de sécurité fournit les droits et les responsabilités suivants aux administrateurs.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-200">The security model provides the following rights and responsibilities to administrators.</span></span>

- <span data-ttu-id="2d0a9-201">Peut empêcher les utilisateurs d’installer un complément Outlook, notamment les compléments sur AppSource.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-201">Can prevent end users from installing any Outlook add-in, including add-ins from AppSource.</span></span>

- <span data-ttu-id="2d0a9-202">Peut désactiver ou activer tout complément Outlook sur le Centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="2d0a9-202">Can disable or enable any Outlook add-in on the Exchange Admin Center.</span></span>

- <span data-ttu-id="2d0a9-203">Applicable uniquement à Outlook sur Windows : peut remplacer les paramètres de seuil de performance par des paramètres du Registre Objet de stratégie de groupe (GPO).</span><span class="sxs-lookup"><span data-stu-id="2d0a9-203">Applicable to only Outlook on Windows: Can override performance threshold settings by GPO registry settings.</span></span>

## <a name="see-also"></a><span data-ttu-id="2d0a9-204">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2d0a9-204">See also</span></span>

- [<span data-ttu-id="2d0a9-205">Confidentialité et sécurité pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="2d0a9-205">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
- [<span data-ttu-id="2d0a9-206">Contrôles de confidentialité pour Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="2d0a9-206">Privacy controls for Microsoft 365 Apps</span></span>](/deployoffice/privacy/overview-privacy-controls)
- [<span data-ttu-id="2d0a9-207">API de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="2d0a9-207">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="2d0a9-208">Limites pour l’activation et l’API JavaScript pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2d0a9-208">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
