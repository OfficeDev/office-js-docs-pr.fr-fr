---
title: Présentation des compléments Outlook
description: Les compléments Outlook sont des intégrations conçues par des tiers dans Outlook à l’aide de notre plate-forme web.
ms.date: 07/16/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 66945f445c89aefc0bf903c4febbc9f51d1f521b90e29144460df9e325826790
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091427"
---
# <a name="outlook-add-ins-overview"></a>Présentation des compléments Outlook

Les compléments Outlook sont des intégrations créées par des tiers dans Outlook à l’aide de notre plateforme web. Les compléments Outlook ont trois aspects clés :

- La même logique complémentaire et commerciale fonctionne sur les ordinateurs de bureau (Outlook sur Windows et Mac), sur le web (Microsoft 365 et Outlook.com) et sur les téléphones portables.
- Les compléments Outlook se composent d’un manifeste, qui décrit la manière dont le complément s’intègre dans Outlook (par exemple, un bouton ou un volet de tâches), ainsi que d’un code JavaScript/HTML, qui constitue l’interface utilisateur et la logique métier du complément.
- Les compléments Outlook peuvent être acquis à partir d’[AppSource](https://appsource.microsoft.com) ou [chargés séparément](sideload-outlook-add-ins-for-testing.md) par les utilisateurs finals ou les administrateurs.

Les compléments Outlook sont différents des compléments COM ou VSTO, qui sont de plus anciennes intégrations propres à Outlook s’exécutant sur Windows. Contrairement aux compléments COM, les compléments Outlook ne disposent d’aucun code installé physiquement sur l’appareil ou le client Outlook de l’utilisateur. Dans le cas d’un complément Outlook, Outlook lit le manifeste et raccorde des contrôles spécifiés dans l’interface utilisateur, puis charge le code JavaScript et HTML. Les composants web s’exécutent tous dans le contexte d’un navigateur dans un bac à sable (sandbox).

Les éléments Outlook qui prennent en charge les compléments incluent notamment les messages électroniques, les demandes de réunion, les réponses à des demandes de réunion, les annulations de réunion et les rendez-vous. Chaque complément définit le contexte dans lequel il est disponible, y compris les types d’éléments et si l’utilisateur lit ou compose un élément.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>Points d’extension

Les points d’extension correspondent à la manière dont les compléments sont intégrés à Outlook. Voici les méthodes possibles.

- Les compléments peuvent indiquer des boutons qui apparaissent dans les surfaces de commande dans les messages et les rendez-vous. Pour plus d’informations, voir [Commandes de complément pour Outlook](add-in-commands-for-outlook.md).

    **Complément avec boutons de commande dans le ruban**

    ![Complément de commande Forme sans interface utilisateur.](../images/uiless-command-shape.png)

- Les compléments peuvent désactiver les correspondances d’expressions régulières ou des entités détectées dans les messages et les rendez-vous. Pour plus d’informations, voir [Compléments Outlook contextuels](contextual-outlook-add-ins.md).

    **Complément contextuel pour une entité en surbrillance (adresse)**

    ![Montre une application contextuelle dans une carte.](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>Éléments de boîtes aux lettres disponibles pour les compléments

Les compléments Outlook s’activent lorsque l’utilisateur compose ou lit un message ou un rendez-vous, mais pas d’autres types d’éléments. Cependant, ils ne sont *pas* activés si l’élément de message actuel, en mode de composition ou de lecture, est l’un des éléments suivants :

- protégé par la Gestion des droits relatifs à l’information (IRM) ou chiffré par d’autres moyens de protection. Un message signé numériquement en est un exemple, puisque la signature numérique dépend de l’un de ces mécanismes ;

  > [!IMPORTANT]
  >
  > - Les compléments s’activent sur les messages signés numériquement dans Outlook avec un abonnement Microsoft 365. Dans Windows, cette prise en charge a été introduite avec le build 8711.1000.
  >
  > - Démarrer avec Outlook build 13229.10000 sur Windows, les compléments peuvent désormais activer les éléments protégés par IRM. Pour plus d’informations sur cette fonctionnalité en mode aperçu, reportez-vous à [Activation de complément sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).

- un rapport ou une notification de remise qui a la classe de message IPM.Report.* (notamment les rapports de remise et les notifications d’échec de remise, ainsi que les notifications de lecture, de non-lecture et de retard) ;

- un fichier .msg ou .eml joint à un autre message ;

- un fichier .msg ou .eml ouvert à partir du système de fichiers ;

- Dans une [boîte aux lettres de groupe](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), dans une boîte aux lettres partagée\*, dans la boîte aux lettres d’un autre utilisateur\*, dans une [boîte aux lettres d’archivage](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-features#archive-mailbox), ou dans un dossier public.

  > [!IMPORTANT]
  > \* La prise en charge des scénarios d’accès délégué (par exemple, les dossiers partagés à partir de la boîte aux lettres d’un autre utilisateur) a été introduite dans [ensemble de conditions requises 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md). La prise en charge des boîtes aux lettres partagées est désormais disponible en préversion. Pour plus d’informations, consultez [Activer les dossiers partagés et les scénarios de boîte aux lettres partagées](delegate-access.md).

- utilise un formulaire personnalisé.

- Créé via [Simple MAPI](https://support.microsoft.com/topic/a3d3f856-eaf6-b6d8-3617-186c0a1123c5). Simple MAPI est utilisé lorsqu'un utilisateur d'Office crée ou envoie un courriel à partir d'une application Office sur Windows pendant qu'Outlook est fermé. Par exemple, un utilisateur peut créer un courrier Outlook tout en travaillant dans Word, ce qui déclenche une fenêtre de composition Outlook sans lancer l’application Outlook complète. Toutefois, si Outlook est déjà en cours d’exécution lorsque l’utilisateur crée l’e-mail à partir de Word, ce n’est pas un scénario Simple MAPI. Les compléments Outlook fonctionnent donc dans le formulaire de composition tant que d’autres exigences d’activation sont remplies.

En général, Outlook peut activer des compléments sous forme de lecture pour les éléments dans le dossier Éléments envoyés, à l'exception des compléments qui s’activent en fonction des correspondances de chaînes d'entités connues. Pour plus d'informations sur les raisons de ce problème, reportez-vous à la rubrique "Prise en charge pour les entités connues" dans [Faire correspondre des chaînes dans un élément Outlook en tant qu'entités connues](match-strings-in-an-item-as-well-known-entities.md).

Il existe actuellement des considérations supplémentaires lors de la conception et de l’implémentation de compléments pour les clients mobiles. Pour plus d’informations, reportez-vous à [Ajouter une prise en charge mobile à un complément Outlook](add-mobile-support.md#compose-mode-and-appointments).

## <a name="supported-clients"></a>Clients pris en charge

Les add-ins Outlook sont pris en charge dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur le web pour Exchange 2013 sur site et versions ultérieures, Outlook sur iOS, Outlook sur Android, et Outlook sur le web et Outlook.com. Les fonctionnalités les plus récentes ne sont pas toutes prises en charge dans tous les [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) à la fois. Reportez-vous aux articles et références API relatives à ces fonctionnalités pour savoir dans quels applications elles peuvent ou non être prises en charge.

## <a name="get-started-building-outlook-add-ins"></a>Commencer à créer des compléments Outlook

Pour commencer à créer des compléments Outlook, procédez comme suit :

- [Démarrage rapide](../quickstarts/outlook-quickstart.md) : créer un volet Office simple.
- [Didacticiel](../tutorials/outlook-tutorial.md) : découvrez comment créer un complément qui insère des gists GitHub dans un nouveau message.

## <a name="see-also"></a>Voir aussi

- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
- [Instructions de conception pour les compléments Office](../design/add-in-design.md)
- [Gérer les licences de compléments pour Office et SharePoint](/office/dev/store/license-your-add-ins)
- [Publier votre complément Office](../publish/publish.md)
- [Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-the-office-store)
