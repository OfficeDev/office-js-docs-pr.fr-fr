---
title: Confidentialité, autorisations et sécurité pour les compléments Outlook
description: Découvrez comment gérer la confidentialité, les autorisations et la sécurité dans un complément Outlook.
ms.date: 10/07/2020
localization_priority: Priority
ms.openlocfilehash: aa30b4c9aff9a07761d06ae538d56a01f2c30e0d
ms.sourcegitcommit: 4bfef315102bd5b4333ff9aeaa6537cffb5bca9e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2020
ms.locfileid: "48398415"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Confidentialité, autorisations et sécurité pour les compléments Outlook

Les utilisateurs finaux, les développeurs et les administrateurs peuvent appliquer les niveaux d’autorisation hiérarchisés du modèle de sécurité pour les compléments Outlook afin de contrôler les performances et la confidentialité.

Cet article décrit les autorisations que les compléments Outlook peuvent demander, et examine le modèle de sécurité selon les perspectives suivantes.

- **AppSource** : intégrité de complément

- **Utilisateurs** : problèmes de confidentialité et de performance

- **Développeurs** : choix d’autorisations et limites d’utilisation des ressources

- **Administrateurs**: privilèges pour définir des seuils de performances

## <a name="permissions-model"></a>Modèle d’autorisations

Comme la façon dont les clients perçoivent la sécurité des compléments peut avoir une incidence sur l’adoption de ces derniers, la sécurité des compléments Outlook repose sur un modèle d’autorisations à plusieurs niveaux. Un complément Outlook indique le niveau d’autorisations dont il a besoin, identifiant ainsi l’accès dont il peut disposer et les actions qu’il peut effectuer sur les données de la boîte aux lettres du client.

Le schéma de manifeste version 1.1 comprend quatre niveaux d’autorisation.

**Tableau 1. Niveaux d’autorisation d’un complément**

|**Niveau d’autorisation**|**Valeur dans le manifeste du complément Outlook**|
|:-----|:-----|
|Restricted|Restreint|
|Lire l’élément|ReadItem|
|Lire/écrire dans l’élément|ReadWriteItem|
|Lire/écrire dans la boîte aux lettres|ReadWriteMailbox|

Les quatre niveaux d’autorisations sont cumulatifs : l’autorisation **boîte aux lettres en lecture/écriture** inclut les autorisations de **élément en lecture/écriture**, **lire élément** et ** restreint**, l’autorisation **élément en lecture/écriture** inclut **lire élément** et **restreint**et l’autorisation **lire élément** inclut **restreint**.

L’illustration suivante affiche les quatre niveaux d’autorisations et décrit les fonctionnalités proposées aux utilisateurs finaux, développeur et administrateur par chaque niveau. Pour plus d’informations sur ces autorisations, voir [utilisateurs : problèmes de performances et de confidentialité](#end-users-privacy-and-performance-concerns), [développeurs : choix d’autorisation et les limites de l’utilisation de ressources](#developers-permission-choices-and-resource-usage-limits), et [comprendre les autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).

**Association du modèle d’autorisation à quatre niveaux à l’utilisateur final, au développeur et à l’administrateur**

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource : intégrité de complément

[AppSource](https://appsource.microsoft.com) héberge des compléments pouvant être installés par les utilisateurs finals et les administrateurs. AppSource applique les mesures suivantes pour maintenir l’intégrité de ces compléments Outlook.

- Oblige le serveur hôte d’un complément à toujours utiliser SSL (Secure Socket Layer) pour communiquer.

- Oblige un développeur à fournir une preuve d’identité, un accord contractuel et une politique de confidentialité conforme pour soumettre les compléments.

- Archive les compléments en mode lecture seule.

- Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.

## <a name="optional-connected-experiences"></a>Expériences connectées facultatives

Les utilisateurs finaux et les administrateurs informatiques peuvent désactiver [expériences connectées facultatives dans ](/deployoffice/privacy/optional-connected-experiences) les clients de bureau et mobiles Office. Pour les compléments Outlook, l’impact de la désactivation du paramètres **Expériences connectées optionnelles** dépend du client, mais les compléments installés par l’utilisateur et l’accès à Office Store ne sont généralement pas autorisés. Certains compléments Microsoft sont considérés comme essentiels ou stratégiques, et les compléments déployés par l’administrateur informatique d’une organisation via [Déploiement centralisé](../publish/centralized-deployment.md) restent disponibles.

- Windows\*, Mac : le bouton **Obtenir des compléments** ne s’affiche pas afin que les utilisateurs ne puissent plus gérer leurs compléments ni accéder à Office Store.
- Android, iOS : la boîte de dialogue **Obtenir des compléments** affiche uniquement les compléments déployés par l’administrateur.
- Navigateur : la disponibilité des compléments et l’accès au Store ne sont pas affectés de sorte que les utilisateurs puissent continuer à [gérer leurs compléments](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), y compris ceux déployés par l’administrateur.

  > [!NOTE]
  > \* Pour Windows, la prise en charge de cette expérience/ce comportement est disponible à partir de la version 2008 (build 13127.20296). Pour plus d’informations en fonction de votre version, consultez la page de l’historique des mises à jour de [Miicrosoft 365](/officeupdates/update-history-office365-proplus-by-date) et [comment trouver la version du client et le canal de mise à jour Office que vous utilisez](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

Pour obtenir des informations générales sur le comportement des compléments, consultez [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md#optional-connected-experiences).

## <a name="end-users-privacy-and-performance-concerns"></a>Utilisateurs : problèmes de confidentialité et de performance

Le modèle de sécurité résout les problèmes de sécurité, de confidentialité et de performance des utilisateurs des manières suivantes.

- Les messages des utilisateurs qui sont protégés par la Gestion des droits relatifs à l’information (IRM) d’Outlook n’ont pas d’interaction avec les compléments Outlook.

  > [!IMPORTANT]
  > - Les compléments s’activent sur les messages signés numériquement dans Outlook avec un abonnement Microsoft 365. Dans Windows, cette prise en charge a été introduite avec le build 8711.1000.
  >
  > - Démarrer avec Outlook build 13229.10000 sur Windows, les compléments peuvent désormais activer les éléments protégés par IRM. Pour plus d’informations sur cette fonctionnalité en mode aperçu, voir [Activation de complément sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).

- Avant d’installer un complément d’AppSource, les utilisateurs finals peuvent voir l’accès dont peut disposer le complément, ainsi que les actions qu’il peut effectuer sur leurs données, et doivent explicitement confirmer qu’ils veulent poursuivre. Aucun complément Outlook n’est automatiquement transmis sur un ordinateur client sans une validation manuelle par l’utilisateur ou l’administrateur.

- L’octroi de l’autorisation **Restreint** permet au complément Outlook d’avoir un accès limité uniquement sur l’élément actuel. L’octroi de l’autorisation **Lire l’élément** permet au complément Outlook d’accéder à des informations d’identification personnelle, par exemple les noms et les adresses électroniques des expéditeurs et des destinataires, uniquement sur l’élément actuel.

- Un utilisateur final peut installer un complément Outlook uniquement pour lui-même. Les compléments de messagerie ayant une incidence sur l’organisation sont installés par un administrateur.

- Les utilisateurs peuvent installer des compléments Outlook qui activent des scénarios contextuels prisés par les utilisateurs tout en minimisant les risques de sécurité pour ces derniers.

- Les fichiers manifeste de compléments Outlook installés sont sécurisés dans le compte de messagerie de l’utilisateur.

- Les données échangées avec des serveurs hébergeant des Compléments Office sont toujours chiffrées conformément au protocole SSL (Secure Socket Layer).

- Applicable uniquement aux clients riches Outlook : les clients riches Outlook surveillent la performance des compléments Outlook installés, exercent un contrôle de gouvernance et désactivent les compléments Outlook qui dépassent les limites pour les aspects suivants.

  - Temps de réponse d’activation

  - Nombre de défaillances d’activation ou de réactivation

  - Utilisation de la mémoire

  - Utilisation du processeur  

  La gouvernance dissuade les attaques par déni de service et maintient les performances des compléments à un niveau raisonnable. La barre Entreprise indique aux utilisateurs les compléments Outlook que le client riche Outlook a désactivés sur la base d’un tel contrôle de gouvernance.

- À tout moment, les utilisateurs finals peuvent vérifier les autorisations demandées par les compléments Outlook installés, et désactiver ou activer ultérieurement tout complément Outlook dans le Centre d’administration Exchange.

## <a name="developers-permission-choices-and-resource-usage-limits"></a>Développeurs : choix d’autorisations et limites d’utilisation des ressources.

Le modèle de sécurité fournit aux développeurs des niveaux précis d’autorisations à choisir, et de strictes directives de performance à observer.

### <a name="tiered-permissions-increases-transparency"></a>Les autorisations à plusieurs niveaux augmentent la transparence

Les développeurs doivent suivre le modèle d’autorisations à plusieurs niveaux pour assurer la transparence et apaiser les inquiétudes des utilisateurs concernant ce que les compléments peuvent faire à leurs données et leur boîte aux lettres, en faisant la promotion indirecte de l’adoption du complément.

- Les développeurs demandent un niveau approprié d’autorisation pour un complément Outlook en fonction de la manière dont il doit être activé, et de son besoin de lire ou d’écrire certaines propriétés d’un élément, ou de créer et d’envoyer un élément.

- Les développeurs demandent une autorisation en utilisant l’élément [Permissions](../reference/manifest/permissions.md) dans le manifeste du complément Outlook, en affectant une valeur **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**, selon le cas.

  > [!NOTE]
  > Notez que l’autorisation **ReadWriteItem** est disponible à partir du schéma de manifeste version 1.1.

  L’exemple suivant demande l’autorisation **Lire l’élément**.

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- Les développeurs peuvent demander l'autorisation **restricted** si le complément Outlook s'active sur un type spécifique d'éléments Outlook (rendez-vous ou message), ou sur des entités extraites spécifiques (numéro de téléphone, adresse, URL) présentes dans le sujet ou dans le corps de l'élément. Par exemple, la règle suivante active le complément Outlook si une ou plusieurs des trois entités (numéro de téléphone, adresse postale ou URL) se trouvent dans l'objet ou le corps du message courant.

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

- Les développeurs doivent demander le **lire élément** autorisation si le complément Outlook a besoin lire les propriétés de l’élément actif autre que les entités extrait par défaut, ou écrire des propriétés personnalisées définies par le complément, sur l’élément actif, mais nécessitent pas de lecture ou écrire à d’autres éléments ou création ou envoyer un message de boîte aux lettres de l’utilisateur. Par exemple, un développeur doit demander l’autorisation **lire élément** si un complément Outlook doit rechercher une entité comme une suggestion de réunion, une suggestion de tâche, une adresse e-mail ou un nom de contact dans le sujet ou le corps de l'élément, ou utilise une expression régulière pour se faire activer.

- Les développeurs doivent demander l’autorisation **Lire/écrire dans l’élément** si le complément Outlook doit écrire dans les propriétés de l’élément composé, comme les noms des destinataires, les adresses de messagerie, le corps et l’objet, ou s’il a besoin d’ajouter ou de supprimer des pièces jointes d’élément.

- Les développeurs demandent l’autorisation **Lire/écrire dans la boîte aux lettres** uniquement si le complément Outlook doit effectuer une ou plusieurs des actions suivantes à l’aide de la méthode [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

  - Lire ou écrire des propriétés d’éléments dans la boîte aux lettres.
  - Créer, lire, écrire ou envoyer des éléments dans la boîte aux lettres.
  - Créer, lire ou écrire dans des dossiers de la boîte aux lettres.

### <a name="resource-usage-tuning"></a>Réglage de l’utilisation des ressources

Les développeurs doivent connaître les limites de l’utilisation des ressources pour l’activation, incorporer le réglage des performances dans leur flux de travail de développement, afin de réduire le risque d’un complément peu performant refusant le service de l’hôte. Les développeurs doivent suivre les directives concernant la conception des règles d’activation telles que décrites dans [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Si un complément Outlook est destiné à être exécuté sur un client riche Outlook, les développeurs doivent vérifier que les performances du complément se situent dans les limites d’utilisation des ressources.

### <a name="other-measures-to-promote-user-security"></a>Autres mesures visant à promouvoir la sécurité de l’utilisateur

Les développeurs doivent connaître et planifier les éléments suivants.

- Les développeurs ne peuvent pas utiliser de contrôles ActiveX dans les compléments car ils ne sont pas pris en charge.

- Les développeurs doivent procéder comme suit lorsqu’ils envoient un complément Outlook à AppSource.

  - Produire un certificat SSL EV (Extended Validation) comme preuve d’identité.

  - Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.

  - Produire une stratégie de confidentialité conforme.

  - Être prêts à signer un accord contractuel lors de la soumission du complément.

## <a name="administrators-privileges"></a>Administrateurs : privilèges

Le modèle de sécurité fournit les droits et les responsabilités suivants aux administrateurs.

- Peut empêcher les utilisateurs d’installer un complément Outlook, notamment les compléments sur AppSource.

- Peut désactiver ou activer tout complément Outlook sur le Centre d’administration Exchange.

- Applicable uniquement à Outlook sur Windows : peut remplacer les paramètres de seuil de performance par des paramètres du Registre Objet de stratégie de groupe (GPO).

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Contrôles de confidentialité pour Microsoft 365 Apps](/deployoffice/privacy/overview-privacy-controls)
- [API de complément Outlook](apis.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
