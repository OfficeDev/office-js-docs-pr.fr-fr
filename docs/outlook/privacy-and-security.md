---
title: Confidentialité, autorisations et sécurité pour les compléments Outlook
description: Découvrez comment gérer la confidentialité, les autorisations et la sécurité dans un complément Outlook.
ms.date: 08/09/2022
ms.localizationpriority: high
ms.openlocfilehash: a19284c6a8371deadcb3986978eabaf605189df6
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092874"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Confidentialité, autorisations et sécurité pour les compléments Outlook

Les utilisateurs finaux, les développeurs et les administrateurs peuvent appliquer les niveaux d’autorisation hiérarchisés du modèle de sécurité pour les compléments Outlook afin de contrôler les performances et la confidentialité.

Cet article décrit les autorisations que les compléments Outlook peuvent demander, et examine le modèle de sécurité selon les perspectives suivantes.

- **AppSource** : intégrité de complément

- **Utilisateurs** : problèmes de confidentialité et de performance

- **Développeurs** : choix d’autorisations et limites d’utilisation des ressources

- **Administrateurs**: privilèges pour définir des seuils de performances

## <a name="permissions-model"></a>Modèle d’autorisations

Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.

Le schéma de manifeste version 1.1 comprend quatre niveaux d’autorisation.

**Tableau 1. Niveaux d’autorisation d’un complément**

|**Niveau d’autorisation**|**Valeur dans le manifeste du complément Outlook**|
|:-----|:-----|
|Restricted|Restreint|
|Lire l’élément|ReadItem|
|Lire/écrire dans l’élément|ReadWriteItem|
|Lire/écrire dans la boîte aux lettres|ReadWriteMailbox|

Les quatre niveaux d’autorisations sont cumulatifs : l’autorisation **boîte aux lettres en lecture/écriture** inclut les autorisations de **élément en lecture/écriture**, **lire élément** et **restreint**, l’autorisation **élément en lecture/écriture** inclut **lire élément** et **restreint** et l’autorisation **lire élément** inclut **restreint**.

L’illustration suivante affiche les quatre niveaux d’autorisations et décrit les fonctionnalités proposées aux utilisateurs finaux, développeur et administrateur par chaque niveau. Pour plus d’informations sur ces autorisations, voir [utilisateurs : problèmes de performances et de confidentialité](#end-users-privacy-and-performance-concerns), [développeurs : choix d’autorisation et les limites de l’utilisation de ressources](#developers-permission-choices-and-resource-usage-limits), et [comprendre les autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).

**Association du modèle d’autorisation à quatre niveaux à l’utilisateur final, au développeur et à l’administrateur**

![Diagramme du modèle d’autorisations à quatre niveaux pour le schéma v1.1 des applications de messagerie.](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource : intégrité de complément

[AppSource](https://appsource.microsoft.com) héberge des compléments pouvant être installés par les utilisateurs finals et les administrateurs. AppSource applique les mesures suivantes pour maintenir l’intégrité de ces compléments Outlook.

- Oblige le serveur hôte d’un complément à toujours utiliser SSL (Secure Socket Layer) pour communiquer.

- Oblige un développeur à fournir une preuve d’identité, un accord contractuel et une politique de confidentialité conforme pour soumettre les compléments.

- Archive les compléments en mode lecture seule.

- Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.

## <a name="optional-connected-experiences"></a>Expériences connectées facultatives

Les utilisateurs finaux et les administrateurs informatiques peuvent désactiver [expériences connectées facultatives dans ](/deployoffice/privacy/optional-connected-experiences) les clients de bureau et mobiles Office. Pour les compléments Outlook, l’impact de la désactivation du paramètre **d’expériences connectées facultatives** dépend du client, mais signifie généralement que les compléments installés par l’utilisateur et l’accès à l’Office Store ne sont pas autorisés. Certains compléments Microsoft sont considérés comme essentiels ou stratégiques, et les compléments déployés par l’administrateur informatique d’une organisation via [Déploiement centralisé](/microsoft-365/admin/manage/centralized-deployment-of-add-ins) restent disponibles.

- Windows\*, Mac : le bouton **Obtenir des compléments** n’est pas affiché pour que les utilisateurs ne puissent plus gérer leurs compléments ni accéder à l’Office Store.
- Android, iOS : la boîte de dialogue **Obtenir des compléments** affiche uniquement les compléments déployés par l’administrateur.
- Navigateur : la disponibilité des compléments et l’accès au Store ne sont pas affectés de sorte que les utilisateurs puissent continuer à [gérer leurs compléments](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), y compris ceux déployés par l’administrateur.

  > [!NOTE]
  > \* Pour Windows, la prise en charge de cette expérience/comportement est disponible à partir de la version 2008 (build 13127.20296). Pour plus d’informations en fonction de votre version, consultez la page de l’historique des mises à jour de [Miicrosoft 365](/officeupdates/update-history-office365-proplus-by-date) et [comment trouver la version du client et le canal de mise à jour Office que vous utilisez](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

Pour obtenir des informations générales sur le comportement des compléments, consultez [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md#optional-connected-experiences).

## <a name="end-users-privacy-and-performance-concerns"></a>Utilisateurs : problèmes de confidentialité et de performance

Le modèle de sécurité résout les problèmes de sécurité, de confidentialité et de performance des utilisateurs des manières suivantes.

- Les messages de l’utilisateur final qui sont protégés par la gestion des droits relatifs à l’information (IRM) d’Outlook n’interagissent pas avec les compléments Outlook sur les clients non Windows.

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed. No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.

- L’octroi de l’autorisation **Restreint** permet au complément Outlook d’avoir un accès limité uniquement sur l’élément actuel.  L’octroi de l’autorisation **d’élément de lecture** permet au complément Outlook d’accéder aux informations d’identification personnelles, telles que les noms d’expéditeur et de destinataire et les adresses e-mail, uniquement sur l’élément actif.

- An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.

- Les utilisateurs peuvent installer des compléments Outlook qui activent des scénarios contextuels prisés par les utilisateurs tout en minimisant les risques de sécurité pour ces derniers.

- Les fichiers manifeste de compléments Outlook installés sont sécurisés dans le compte de messagerie de l’utilisateur.

- Les données échangées avec des serveurs hébergeant des Compléments Office sont toujours chiffrées conformément au protocole SSL (Secure Socket Layer).

- Applicable uniquement aux clients riches Outlook : les clients riches Outlook surveillent la performance des compléments Outlook installés, exercent un contrôle de gouvernance et désactivent les compléments Outlook qui dépassent les limites pour les aspects suivants.

  - Temps de réponse d’activation

  - Nombre de défaillances d’activation ou de réactivation

  - Utilisation de la mémoire

  - Utilisation du processeur  

  Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.

- À tout moment, les utilisateurs finals peuvent vérifier les autorisations demandées par les compléments Outlook installés, et désactiver ou activer ultérieurement tout complément Outlook dans le Centre d’administration Exchange.

## <a name="developers-permission-choices-and-resource-usage-limits"></a>Développeurs : choix d’autorisations et limites d’utilisation des ressources.

Le modèle de sécurité fournit aux développeurs des niveaux précis d’autorisations à choisir, et de strictes directives de performance à observer.

### <a name="tiered-permissions-increases-transparency"></a>Les autorisations à plusieurs niveaux augmentent la transparence

Les développeurs doivent suivre le modèle d’autorisations à plusieurs niveaux pour assurer la transparence et apaiser les inquiétudes des utilisateurs concernant ce que les compléments peuvent faire à leurs données et leur boîte aux lettres, en faisant la promotion indirecte de l’adoption du complément.

- Les développeurs demandent un niveau approprié d’autorisation pour un complément Outlook en fonction de la manière dont il doit être activé, et de son besoin de lire ou d’écrire certaines propriétés d’un élément, ou de créer et d’envoyer un élément.

- Les développeurs demandent une autorisation en utilisant l’élément [Permissions](/javascript/api/manifest/permissions) dans le manifeste du complément Outlook, en affectant une valeur **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**, selon le cas.

  > [!NOTE]
  > Notez que l’autorisation **ReadWriteItem** est disponible à partir du schéma de manifeste version 1.1.

  L’exemple suivant demande l’autorisation **Lire l’élément**.

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- Les développeurs peuvent demander l’autorisation **restreinte** si le complément Outlook s’active sur un type spécifique d’élément Outlook (rendez-vous ou message) ou sur des entités extraites spécifiques (numéro de téléphone, adresse, URL) présentes dans l’objet ou le corps de l’élément. Par exemple, la règle suivante active le complément Outlook si une ou plusieurs des trois entités (numéro de téléphone, adresse postale ou URL) se trouvent dans l'objet ou le corps du message courant.

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

- Les développeurs doivent demander l’autorisation **d’élément de lecture** si le complément Outlook doit lire les propriétés de l’élément actif autres que les entités extraites par défaut, ou écrire des propriétés personnalisées définies par le complément sur l’élément actif, mais ne nécessite pas de lecture ou d’écriture dans d’autres éléments, ni la création ou l’envoi d’un message dans la boîte aux lettres de l’utilisateur. Par exemple, un développeur doit demander l’autorisation **lire élément** si un complément Outlook doit rechercher une entité comme une suggestion de réunion, une suggestion de tâche, une adresse e-mail ou un nom de contact dans le sujet ou le corps de l'élément, ou utilise une expression régulière pour se faire activer.

- Les développeurs doivent demander l’autorisation **Lire/écrire dans l’élément** si le complément Outlook doit écrire dans les propriétés de l’élément composé, comme les noms des destinataires, les adresses de messagerie, le corps et l’objet, ou s’il a besoin d’ajouter ou de supprimer des pièces jointes d’élément.

- Les développeurs demandent l’autorisation **Lire/écrire dans la boîte aux lettres** uniquement si le complément Outlook doit effectuer une ou plusieurs des actions suivantes à l’aide de la méthode [mailbox.makeEWSRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

  - Lire ou écrire des propriétés d’éléments dans la boîte aux lettres.
  - Créer, lire, écrire ou envoyer des éléments dans la boîte aux lettres.
  - Créer, lire ou écrire dans des dossiers de la boîte aux lettres.

### <a name="resource-usage-tuning"></a>Réglage de l’utilisation des ressources

Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.

### <a name="other-measures-to-promote-user-security"></a>Autres mesures visant à promouvoir la sécurité de l’utilisateur

Les développeurs doivent connaître et planifier les éléments suivants.

- Les développeurs ne peuvent pas utiliser de contrôles ActiveX dans les compléments, car ils ne sont pas pris en charge.

- Les développeurs doivent procéder comme suit lorsqu’ils envoient un complément Outlook à AppSource.

  - Produire un certificat SSL EV (Extended Validation) comme preuve d’identité.

  - Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.

  - Produire une stratégie de confidentialité conforme.

  - Être prêts à signer un accord contractuel lors de la soumission du complément.

## <a name="administrators-privileges"></a>Administrateurs : privilèges

Le modèle de sécurité fournit les droits et les responsabilités suivants aux administrateurs.

- Peut empêcher les utilisateurs d’installer un complément Outlook, notamment les compléments sur AppSource.

- Peut désactiver ou activer tout complément Outlook sur le Centre d’administration Exchange.

- Applicable uniquement à Outlook sur Windows : peut remplacer les paramètres de seuil de performance par des paramètres du Registre Objet de stratégie de groupe (GPO).

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Contrôles de confidentialité pour Microsoft 365 Apps](/deployoffice/privacy/overview-privacy-controls)
- [API de complément Outlook](apis.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
