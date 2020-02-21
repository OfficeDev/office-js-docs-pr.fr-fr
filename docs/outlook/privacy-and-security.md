---
title: Confidentialité, autorisations et sécurité pour les compléments Outlook
description: Découvrez comment gérer la confidentialité, les autorisations et la sécurité dans un complément Outlook.
ms.date: 10/31/2019
localization_priority: Priority
ms.openlocfilehash: e35b5d2328e7be8e32b3bd093c44eb6846bc759f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166091"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a><span data-ttu-id="1158b-103">Confidentialité, autorisations et sécurité pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="1158b-103">Privacy, permissions, and security for Outlook add-ins</span></span>

<span data-ttu-id="1158b-104">Les utilisateurs finaux, les développeurs et les administrateurs peuvent appliquer les niveaux d’autorisation hiérarchisés du modèle de sécurité pour les compléments Outlook afin de contrôler les performances et la confidentialité.</span><span class="sxs-lookup"><span data-stu-id="1158b-104">End users, developers, and administrators can use the tiered permission levels of the security model for Outlook add-ins to control privacy and performance.</span></span>

<span data-ttu-id="1158b-105">Cet article décrit les autorisations que les compléments Outlook peuvent demander, et examine le modèle de sécurité selon les perspectives suivantes :</span><span class="sxs-lookup"><span data-stu-id="1158b-105">This article describes the possible permissions that Outlook add-ins can request, and examines the security model from the following perspectives:</span></span>

- <span data-ttu-id="1158b-106">**AppSource** : intégrité de complément</span><span class="sxs-lookup"><span data-stu-id="1158b-106">**AppSource**: add-in integrity</span></span>
    
- <span data-ttu-id="1158b-107">**Utilisateurs** : problèmes de confidentialité et de performance</span><span class="sxs-lookup"><span data-stu-id="1158b-107">**End-users**: privacy and performance concerns</span></span>
    
- <span data-ttu-id="1158b-108">**Développeurs** : choix d’autorisations et limites d’utilisation des ressources</span><span class="sxs-lookup"><span data-stu-id="1158b-108">**Developers**: permissions choices and resource usage limits</span></span>
    
- <span data-ttu-id="1158b-109">**Administrateurs**: privilèges pour définir des seuils de performances</span><span class="sxs-lookup"><span data-stu-id="1158b-109">**Administrators**: privileges to set performance thresholds</span></span>
    

## <a name="permissions-model"></a><span data-ttu-id="1158b-110">Modèle d’autorisations</span><span class="sxs-lookup"><span data-stu-id="1158b-110">Permissions model</span></span>

<span data-ttu-id="1158b-p101">Comme la façon dont les clients perçoivent la sécurité des compléments peut avoir une incidence sur l’adoption de ces derniers, la sécurité des compléments Outlook repose sur un modèle d’autorisations à plusieurs niveaux. Un complément Outlook indique le niveau d’autorisations dont il a besoin, identifiant ainsi l’accès dont il peut disposer et les actions qu’il peut effectuer sur les données de la boîte aux lettres du client.</span><span class="sxs-lookup"><span data-stu-id="1158b-p101">Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.</span></span> 

<span data-ttu-id="1158b-113">Le schéma de manifeste version 1.1 comprend quatre niveaux d’autorisation.</span><span class="sxs-lookup"><span data-stu-id="1158b-113">Manifest schema version 1.1 includes four levels of permissions.</span></span> 


<span data-ttu-id="1158b-114">**Tableau 1. Niveaux d’autorisation d’un complément**</span><span class="sxs-lookup"><span data-stu-id="1158b-114">**Table 1. Add-in permission levels**</span></span>

|<span data-ttu-id="1158b-115">**Niveau d’autorisation**</span><span class="sxs-lookup"><span data-stu-id="1158b-115">**Permission level**</span></span>|<span data-ttu-id="1158b-116">**Valeur dans le manifeste du complément Outlook**</span><span class="sxs-lookup"><span data-stu-id="1158b-116">**Value in Outlook add-in manifest**</span></span>|
|:-----|:-----|
|<span data-ttu-id="1158b-117">Restricted</span><span class="sxs-lookup"><span data-stu-id="1158b-117">Restricted</span></span>|<span data-ttu-id="1158b-118">Restreint</span><span class="sxs-lookup"><span data-stu-id="1158b-118">Restricted</span></span>|
|<span data-ttu-id="1158b-119">Lire l’élément</span><span class="sxs-lookup"><span data-stu-id="1158b-119">Read item</span></span>|<span data-ttu-id="1158b-120">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1158b-120">ReadItem</span></span>|
|<span data-ttu-id="1158b-121">Lire/écrire dans l’élément</span><span class="sxs-lookup"><span data-stu-id="1158b-121">Read/write item</span></span>|<span data-ttu-id="1158b-122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1158b-122">ReadWriteItem</span></span>|
|<span data-ttu-id="1158b-123">Lire/écrire dans la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1158b-123">Read/write mailbox</span></span>|<span data-ttu-id="1158b-124">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1158b-124">ReadWriteMailbox</span></span>|

<span data-ttu-id="1158b-125">Les quatre niveaux d’autorisations sont cumulatifs : l’autorisation **boîte aux lettres en lecture/écriture** inclut les autorisations de **élément en lecture/écriture**, **lire élément** et \*\* restreint\*\*, l’autorisation **élément en lecture/écriture** inclut **lire élément** et **restreint**et l’autorisation **lire élément** inclut **restreint**.</span><span class="sxs-lookup"><span data-stu-id="1158b-125">The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.</span></span> 

<span data-ttu-id="1158b-126">L’illustration suivante affiche les quatre niveaux d’autorisations et décrit les fonctionnalités proposées aux utilisateurs finaux, développeur et administrateur par chaque niveau.</span><span class="sxs-lookup"><span data-stu-id="1158b-126">The following figure shows the four levels of permissions and describes the capabilities offered to the end user, developer, and administrator by each tier.</span></span> <span data-ttu-id="1158b-127">Pour plus d’informations sur ces autorisations, voir [utilisateurs : problèmes de performances et de confidentialité](#end-users-privacy-and-performance-concerns), [développeurs : choix d’autorisation et les limites de l’utilisation de ressources](#developers-permission-choices-and-resource-usage-limits), et [comprendre les autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="1158b-127">For more information about these permissions, see [End users: privacy and performance concerns](#end-users-privacy-and-performance-concerns), [Developers: permission choices and resource usage limits](#developers-permission-choices-and-resource-usage-limits), and [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span> 


<span data-ttu-id="1158b-128">**Association du modèle d’autorisation à quatre niveaux à l’utilisateur final, au développeur et à l’administrateur**</span><span class="sxs-lookup"><span data-stu-id="1158b-128">**Relating the four-tier permission model to the end user, developer, and administrator**</span></span>

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1](../images/add-in-permission-tiers.png)


## <a name="appsource-add-in-integrity"></a><span data-ttu-id="1158b-130">AppSource : intégrité de complément</span><span class="sxs-lookup"><span data-stu-id="1158b-130">AppSource: add-in integrity</span></span>

<span data-ttu-id="1158b-131">[AppSource](https://appsource.microsoft.com) héberge des compléments pouvant être installés par les utilisateurs finals et les administrateurs.</span><span class="sxs-lookup"><span data-stu-id="1158b-131">[AppSource](https://appsource.microsoft.com) hosts add-ins that can be installed by end users and administrators.</span></span> <span data-ttu-id="1158b-132">AppSource applique les mesures suivantes pour maintenir l’intégrité de ces compléments Outlook :</span><span class="sxs-lookup"><span data-stu-id="1158b-132">AppSource enforces the following measures to maintain the integrity of these Outlook add-ins:</span></span>

- <span data-ttu-id="1158b-133">Oblige le serveur hôte d’un complément à toujours utiliser SSL (Secure Socket Layer) pour communiquer.</span><span class="sxs-lookup"><span data-stu-id="1158b-133">Requires the host server of an add-in to always use Secure Socket Layer (SSL) to communicate.</span></span>
    
- <span data-ttu-id="1158b-134">Oblige un développeur à fournir une preuve d’identité, un accord contractuel et une politique de confidentialité conforme pour soumettre les compléments.</span><span class="sxs-lookup"><span data-stu-id="1158b-134">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span> 
    
- <span data-ttu-id="1158b-135">Archive les compléments en mode lecture seule.</span><span class="sxs-lookup"><span data-stu-id="1158b-135">Archives add-ins in read-only mode.</span></span>
    
- <span data-ttu-id="1158b-136">Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.</span><span class="sxs-lookup"><span data-stu-id="1158b-136">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>
    

## <a name="end-users-privacy-and-performance-concerns"></a><span data-ttu-id="1158b-137">Utilisateurs : problèmes de confidentialité et de performance</span><span class="sxs-lookup"><span data-stu-id="1158b-137">End users: privacy and performance concerns</span></span>

<span data-ttu-id="1158b-138">Le modèle de sécurité résout les problèmes de sécurité, de confidentialité et de performance des utilisateurs des manières suivantes :</span><span class="sxs-lookup"><span data-stu-id="1158b-138">The security model addresses security, privacy, and performance concerns of end users in the following ways:</span></span>

- <span data-ttu-id="1158b-139">Les messages des utilisateurs qui sont protégés par la Gestion des droits relatifs à l’information (IRM) d’Outlook n’ont pas d’interaction avec les compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="1158b-139">End user's messages that are protected by Outlook's Information Rights Management (IRM) do not interact with Outlook add-ins.</span></span>
    
- <span data-ttu-id="1158b-140">Avant d’installer un complément d’AppSource, les utilisateurs finals peuvent voir l’accès dont peut disposer le complément, ainsi que les actions qu’il peut effectuer sur leurs données, et doivent explicitement confirmer qu’ils veulent poursuivre.</span><span class="sxs-lookup"><span data-stu-id="1158b-140">Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed.</span></span> <span data-ttu-id="1158b-141">Aucun complément Outlook n’est automatiquement transmis sur un ordinateur client sans une validation manuelle par l’utilisateur ou l’administrateur.</span><span class="sxs-lookup"><span data-stu-id="1158b-141">No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.</span></span>
    
- <span data-ttu-id="1158b-p105">L’octroi de l’autorisation **Restreint** permet au complément Outlook d’avoir un accès limité uniquement sur l’élément actuel. L’octroi de l’autorisation **Lire l’élément** permet au complément Outlook d’accéder à des informations d’identification personnelle, par exemple les noms et les adresses électroniques des expéditeurs et des destinataires, uniquement sur l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="1158b-p105">Granting the **restricted** permission allows the Outlook add-in to have limited access on only the current item. Granting the **read item** permission allows the Outlook add-in to access personal identifiable information, such as sender and recipient names and email addresses, on only the current item,.</span></span>
    
- <span data-ttu-id="1158b-p106">Un utilisateur final peut installer un complément Outlook uniquement pour lui-même. Les compléments de messagerie ayant une incidence sur l’organisation sont installés par un administrateur.</span><span class="sxs-lookup"><span data-stu-id="1158b-p106">An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.</span></span>
    
- <span data-ttu-id="1158b-146">Les utilisateurs peuvent installer des compléments Outlook qui activent des scénarios contextuels prisés par les utilisateurs tout en minimisant les risques de sécurité pour ces derniers.</span><span class="sxs-lookup"><span data-stu-id="1158b-146">End users can install Outlook add-ins that enable context-sensitive scenarios that are compelling to users while minimizing the users' security risks.</span></span>
    
- <span data-ttu-id="1158b-147">Les fichiers manifeste de compléments Outlook installés sont sécurisés dans le compte de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1158b-147">Manifest files of installed Outlook add-ins are secured in the user's email account.</span></span>
    
- <span data-ttu-id="1158b-148">Les données échangées avec des serveurs hébergeant des Compléments Office sont toujours chiffrées conformément au protocole SSL (Secure Socket Layer).</span><span class="sxs-lookup"><span data-stu-id="1158b-148">Data communicated with servers hosting Office Add-ins is always encrypted according to the Secure Socket Layer (SSL) protocol.</span></span>
    
- <span data-ttu-id="1158b-149">Applicable uniquement aux clients riches Outlook : les clients riches Outlook surveillent la performance des compléments Outlook installés, exercent un contrôle de gouvernance et désactivent les compléments Outlook qui dépassent les limites pour les aspects suivants :</span><span class="sxs-lookup"><span data-stu-id="1158b-149">Applicable to only the Outlook rich clients: The Outlook rich clients monitor the performance of installed Outlook add-ins, exercise governance control, and disable those Outlook add-ins that exceed limits in the following areas:</span></span>
    
  - <span data-ttu-id="1158b-150">Temps de réponse d’activation</span><span class="sxs-lookup"><span data-stu-id="1158b-150">Response time to activate</span></span>
    
  - <span data-ttu-id="1158b-151">Nombre de défaillances d’activation ou de réactivation</span><span class="sxs-lookup"><span data-stu-id="1158b-151">Number of failures to activate or reactivate</span></span>
    
  - <span data-ttu-id="1158b-152">Utilisation de la mémoire</span><span class="sxs-lookup"><span data-stu-id="1158b-152">Memory usage</span></span>
    
  - <span data-ttu-id="1158b-153">Utilisation du processeur</span><span class="sxs-lookup"><span data-stu-id="1158b-153">CPU usage</span></span>  

  <span data-ttu-id="1158b-p107">La gouvernance dissuade les attaques par déni de service et maintient les performances des compléments à un niveau raisonnable. La barre Entreprise indique aux utilisateurs les compléments Outlook que le client riche Outlook a désactivés sur la base d’un tel contrôle de gouvernance.</span><span class="sxs-lookup"><span data-stu-id="1158b-p107">Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.</span></span>

- <span data-ttu-id="1158b-156">À tout moment, les utilisateurs finals peuvent vérifier les autorisations demandées par les compléments Outlook installés, et désactiver ou activer ultérieurement tout complément Outlook dans le Centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="1158b-156">At any time, end users can verify the permissions requested by installed Outlook add-ins, and disable or subsequently enable any Outlook add-in in the Exchange Admin Center.</span></span>


## <a name="developers-permission-choices-and-resource-usage-limits"></a><span data-ttu-id="1158b-157">Développeurs : choix d’autorisations et limites d’utilisation des ressources.</span><span class="sxs-lookup"><span data-stu-id="1158b-157">Developers: permission choices and resource usage limits</span></span>

<span data-ttu-id="1158b-158">Le modèle de sécurité fournit aux développeurs des niveaux précis d’autorisations à choisir, et de strictes directives de performance à observer.</span><span class="sxs-lookup"><span data-stu-id="1158b-158">The security model provides developers granular levels of permissions to choose from, and strict performance guidelines to observe.</span></span>

### <a name="tiered-permissions-increases-transparency"></a><span data-ttu-id="1158b-159">Les autorisations à plusieurs niveaux augmentent la transparence</span><span class="sxs-lookup"><span data-stu-id="1158b-159">Tiered permissions increases transparency</span></span>

<span data-ttu-id="1158b-160">Les développeurs doivent suivre le modèle d’autorisations à plusieurs niveaux pour assurer la transparence et apaiser les inquiétudes des utilisateurs concernant ce que les compléments peuvent faire à leurs données et leur boîte aux lettres, en faisant la promotion indirecte de l’adoption du complément :</span><span class="sxs-lookup"><span data-stu-id="1158b-160">Developers should follow the tiered permissions model to provide transparency and alleviate users' concern about what add-ins can do to their data and mailbox, indirectly promoting add-in adoption:</span></span>

- <span data-ttu-id="1158b-161">Les développeurs demandent un niveau approprié d’autorisation pour un complément Outlook en fonction de la manière dont il doit être activé, et de son besoin de lire ou d’écrire certaines propriétés d’un élément, ou de créer et d’envoyer un élément.</span><span class="sxs-lookup"><span data-stu-id="1158b-161">Developers request an appropriate level of permission for an Outlook add-in, based on how the Outlook add-in should be activated, and its need to read or write certain properties of an item, or to create and send an item.</span></span>

- <span data-ttu-id="1158b-162">Les développeurs demandent une autorisation en utilisant l’élément [Permissions](../reference/manifest/permissions.md) dans le manifeste du complément Outlook, en affectant une valeur **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**, selon le cas.</span><span class="sxs-lookup"><span data-stu-id="1158b-162">Developers request permission by using the [Permissions](../reference/manifest/permissions.md) element in the manifest of the Outlook add-in, by assigning a value of **Restricted**, **ReadItem**, **ReadWriteItem** or **ReadWriteMailbox**, as appropriate.</span></span>

  > [!NOTE]
  > <span data-ttu-id="1158b-163">Notez que l’autorisation **ReadWriteItem** est disponible à partir du schéma de manifeste version 1.1.</span><span class="sxs-lookup"><span data-stu-id="1158b-163">Note that the **ReadWriteItem** permission is available starting in manifest schema v1.1.</span></span>

  <span data-ttu-id="1158b-164">L’exemple suivant demande l’autorisation **Lire l’élément**.</span><span class="sxs-lookup"><span data-stu-id="1158b-164">The following example requests the **read item** permission.</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- <span data-ttu-id="1158b-165">Les développeurs peuvent demander l'autorisation **restricted** si le complément Outlook s'active sur un type spécifique d'éléments Outlook (rendez-vous ou message), ou sur des entités extraites spécifiques (numéro de téléphone, adresse, URL) présentes dans le sujet ou dans le corps de l'élément.</span><span class="sxs-lookup"><span data-stu-id="1158b-165">Developers can request the **restricted** permission if the Outlook add-in activates on a specific type of Outlook items (appointment or message), or on specific extracted entities (phone number, address, URL) being present in the item's subject or body.</span></span> <span data-ttu-id="1158b-166">Par exemple, la règle suivante active le complément Outlook si une ou plusieurs des trois entités (numéro de téléphone, adresse postale ou URL) se trouvent dans l'objet ou le corps du message courant.</span><span class="sxs-lookup"><span data-stu-id="1158b-166">For example, the following rule activates the Outlook add-in if one or more of three entities - phone number, postal address, or URL - are found in the subject or body of the current message.</span></span>
    
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

- <span data-ttu-id="1158b-167">Les développeurs doivent demander le **lire élément** autorisation si le complément Outlook a besoin lire les propriétés de l’élément actif autre que les entités extrait par défaut, ou écrire des propriétés personnalisées définies par le complément, sur l’élément actif, mais nécessitent pas de lecture ou écrire à d’autres éléments ou création ou envoyer un message de boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1158b-167">Developers should request the **read item** permission if the Outlook add-in needs to read properties of the current item other than the default extracted entities, or write custom properties set by the add-in on the current item, but does not require reading or writing to other items, or creating or sending a message in the user's mailbox.</span></span> <span data-ttu-id="1158b-168">Par exemple, un développeur doit demander l’autorisation **lire élément** si un complément Outlook doit rechercher une entité comme une suggestion de réunion, une suggestion de tâche, une adresse e-mail ou un nom de contact dans le sujet ou le corps de l'élément, ou utilise une expression régulière pour se faire activer.</span><span class="sxs-lookup"><span data-stu-id="1158b-168">For example, a developer should request **read item** permission if an Outlook add-in needs to look for an entity like a meeting suggestion, task suggestion, email address, or contact name in the item's subject or body, or uses a regular expression to activate.</span></span>

- <span data-ttu-id="1158b-169">Les développeurs doivent demander l’autorisation **Lire/écrire dans l’élément** si le complément Outlook doit écrire dans les propriétés de l’élément composé, comme les noms des destinataires, les adresses de messagerie, le corps et l’objet, ou s’il a besoin d’ajouter ou de supprimer des pièces jointes d’élément.</span><span class="sxs-lookup"><span data-stu-id="1158b-169">Developers should request the **read/write item** permission if the Outlook add-in needs to write to properties of the composed item, such as recipient names, email addresses, body, and subject, or needs to add or remove item attachments.</span></span>

- <span data-ttu-id="1158b-170">Les développeurs demandent l’autorisation **Lire/écrire dans la boîte aux lettres** uniquement si le complément Outlook doit effectuer une ou plusieurs des actions suivantes à l’aide de la méthode [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) :</span><span class="sxs-lookup"><span data-stu-id="1158b-170">Developers request the **read/write mailbox** permission only if the Outlook add-in needs to do one or more of the following actions by using the [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method:</span></span>

  - <span data-ttu-id="1158b-171">Lire ou écrire des propriétés d’éléments dans la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1158b-171">Read or write to properties of items in the mailbox.</span></span>
  - <span data-ttu-id="1158b-172">Créer, lire, écrire ou envoyer des éléments dans la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1158b-172">Create, read, write, or send items in the mailbox.</span></span>
  - <span data-ttu-id="1158b-173">Créer, lire ou écrire dans des dossiers de la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1158b-173">Create, read, or write to folders in the mailbox.</span></span>


### <a name="resource-usage-tuning"></a><span data-ttu-id="1158b-174">Réglage de l’utilisation des ressources</span><span class="sxs-lookup"><span data-stu-id="1158b-174">Resource usage tuning</span></span>

<span data-ttu-id="1158b-p110">Les développeurs doivent connaître les limites de l’utilisation des ressources pour l’activation, incorporer le réglage des performances dans leur flux de travail de développement, afin de réduire le risque d’un complément peu performant refusant le service de l’hôte. Les développeurs doivent suivre les directives concernant la conception des règles d’activation telles que décrites dans [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Si un complément Outlook est destiné à être exécuté sur un client riche Outlook, les développeurs doivent vérifier que les performances du complément se situent dans les limites d’utilisation des ressources.</span><span class="sxs-lookup"><span data-stu-id="1158b-p110">Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.</span></span>


### <a name="other-measures-to-promote-user-security"></a><span data-ttu-id="1158b-177">Autres mesures visant à promouvoir la sécurité de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="1158b-177">Other measures to promote user security</span></span>

<span data-ttu-id="1158b-178">Les développeurs doivent connaître et planifier les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="1158b-178">Developers should be aware of and plan for the following as well:</span></span>

- <span data-ttu-id="1158b-179">Les développeurs ne peuvent pas utiliser de contrôles ActiveX dans les compléments car ils ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="1158b-179">Developers cannot use ActiveX controls in add-ins because they are not supported.</span></span>
    
- <span data-ttu-id="1158b-180">Les développeurs doivent procéder comme suit lorsqu’ils envoient un complément Outlook à AppSource :</span><span class="sxs-lookup"><span data-stu-id="1158b-180">Developers should do the following when submitting an Outlook add-in to AppSource:</span></span>
    
  - <span data-ttu-id="1158b-181">Produire un certificat SSL EV (Extended Validation) comme preuve d’identité.</span><span class="sxs-lookup"><span data-stu-id="1158b-181">Produce an Extended Validation (EV) SSL certificate as a proof of identity.</span></span>
    
  - <span data-ttu-id="1158b-182">Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.</span><span class="sxs-lookup"><span data-stu-id="1158b-182">Host the add-in they are submitting on a web server that supports SSL.</span></span>
    
  - <span data-ttu-id="1158b-183">Produire une stratégie de confidentialité conforme.</span><span class="sxs-lookup"><span data-stu-id="1158b-183">Produce a compliant privacy policy.</span></span>
    
  - <span data-ttu-id="1158b-184">Être prêts à signer un accord contractuel lors de la soumission du complément.</span><span class="sxs-lookup"><span data-stu-id="1158b-184">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>
    

## <a name="administrators-privileges"></a><span data-ttu-id="1158b-185">Administrateurs : privilèges</span><span class="sxs-lookup"><span data-stu-id="1158b-185">Administrators: privileges</span></span>

<span data-ttu-id="1158b-186">Le modèle de sécurité fournit les droits et les responsabilités suivants aux administrateurs :</span><span class="sxs-lookup"><span data-stu-id="1158b-186">The security model provides the following rights and responsibilities to administrators:</span></span>

- <span data-ttu-id="1158b-187">Peut empêcher les utilisateurs d’installer un complément Outlook, notamment les compléments sur AppSource.</span><span class="sxs-lookup"><span data-stu-id="1158b-187">Can prevent end users from installing any Outlook add-in, including add-ins from AppSource.</span></span>
    
- <span data-ttu-id="1158b-188">Peut désactiver ou activer tout complément Outlook sur le Centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="1158b-188">Can disable or enable any Outlook add-in on the Exchange Admin Center.</span></span>
    
- <span data-ttu-id="1158b-189">Applicable uniquement à Outlook sur Windows : peut remplacer les paramètres de seuil de performance par des paramètres du Registre Objet de stratégie de groupe (GPO).</span><span class="sxs-lookup"><span data-stu-id="1158b-189">Applicable to only Outlook on Windows: Can override performance threshold settings by GPO registry settings.</span></span>
    


## <a name="see-also"></a><span data-ttu-id="1158b-190">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1158b-190">See also</span></span>

- [<span data-ttu-id="1158b-191">Confidentialité et sécurité pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="1158b-191">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)    
- [<span data-ttu-id="1158b-192">API de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="1158b-192">Outlook add-in APIs</span></span>](apis.md)    
- [<span data-ttu-id="1158b-193">Limites pour l’activation et l’API JavaScript pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="1158b-193">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
