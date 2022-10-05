---
title: Présentation des compléments Outlook
description: Les compléments Outlook sont des intégrations conçues par des tiers dans Outlook à l’aide de notre plate-forme web.
ms.date: 08/09/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: fd17728f840188fbedfdeba7d3ee8f97852d702a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467257"
---
# <a name="outlook-add-ins-overview"></a>Présentation des compléments Outlook

Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:

- La même logique complémentaire et commerciale fonctionne sur les ordinateurs de bureau (Outlook sur Windows et Mac), sur le web (Microsoft 365 et Outlook.com) et sur les téléphones portables.
- Les compléments Outlook se composent d’un manifeste, qui décrit la manière dont le complément s’intègre dans Outlook (par exemple, un bouton ou un volet de tâches), ainsi que d’un code JavaScript/HTML, qui constitue l’interface utilisateur et la logique métier du complément.
- Les compléments Outlook peuvent être acquis à partir d’[AppSource](https://appsource.microsoft.com) ou [chargés séparément](sideload-outlook-add-ins-for-testing.md) par les utilisateurs finals ou les administrateurs.

Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.

The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>Points d’extension

Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done.

- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

    **Complément avec boutons de commande dans le ruban**

    ![Commande de fonction de complément](../images/uiless-command-shape.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).

    **Complément contextuel pour une entité en surbrillance (adresse)**

    ![Montre une application contextuelle dans une carte.](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>Éléments de boîtes aux lettres disponibles pour les compléments

Les compléments Outlook s’activent lorsque l’utilisateur compose ou lit un message ou un rendez-vous, mais pas d’autres types d’éléments. Cependant, ils ne sont *pas* activés si l’élément de message actuel, en mode de composition ou de lecture, est l’un des éléments suivants :

- Protégé par la Gestion des droits relatifs à l’information (IRM) ou chiffré d’autres manières pour la protection et accessible à partir d’Outlook sur des clients non Windows. Un message signé de façon numérique constitue un exemple, puisque la signature numérique dépend de l’un de ces mécanismes ;

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- un rapport ou une notification de remise qui a la classe de message IPM.Report.* (notamment les rapports de remise et les notifications d’échec de remise, ainsi que les notifications de lecture, de non-lecture et de retard) ;

- un fichier .msg ou .eml joint à un autre message ;

- un fichier .msg ou .eml ouvert à partir du système de fichiers ;

- Dans une [boîte aux lettres de groupe](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), dans une boîte aux lettres partagée\*, dans la boîte aux lettres d’un autre utilisateur\*, dans une [boîte aux lettres d’archivage](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-client-and-compliance-&-security-feature-details?tabs=Archive-features#archive-mailbox), ou dans un dossier public.

  > [!IMPORTANT]
  > \* La prise en charge des scénarios d’accès délégué (par exemple, les dossiers partagés à partir de la boîte aux lettres d’un autre utilisateur) a été introduite dans [ensemble de conditions requises 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8). La prise en charge des boîtes aux lettres partagées est désormais en préversion dans Outlook sur Windows et sur Mac. Pour plus d’informations, consultez [Activer les dossiers partagés et les scénarios de boîte aux lettres partagées](delegate-access.md).

- utilise un formulaire personnalisé.

- Créé via Simple MAPI. Simple MAPI est utilisé lorsqu'un utilisateur d'Office crée ou envoie un courriel à partir d'une application Office sur Windows pendant qu'Outlook est fermé. Par exemple, un utilisateur peut créer un courrier Outlook tout en travaillant dans Word, ce qui déclenche une fenêtre de composition Outlook sans lancer l’application Outlook complète. Toutefois, si Outlook est déjà en cours d’exécution lorsque l’utilisateur crée l’e-mail à partir de Word, ce n’est pas un scénario Simple MAPI. Les compléments Outlook fonctionnent donc dans le formulaire de composition tant que d’autres exigences d’activation sont remplies.

En général, Outlook peut activer des compléments sous forme de lecture pour les éléments dans le dossier Éléments envoyés, à l'exception des compléments qui s’activent en fonction des correspondances de chaînes d'entités connues. Pour plus d’informations sur les raisons de ce comportement, voir [Prise en charge des entités connues](match-strings-in-an-item-as-well-known-entities.md#support-for-well-known-entities).

Il existe actuellement des considérations supplémentaires lors de la conception et de l’implémentation de compléments pour les clients mobiles. Pour plus d’informations, consultez [Ajouter une prise en charge mobile à un complément Outlook](add-mobile-support.md#compose-mode-and-appointments).

## <a name="supported-clients"></a>Clients pris en charge

Les add-ins Outlook sont pris en charge dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur le web pour Exchange 2013 sur site et versions ultérieures, Outlook sur iOS, Outlook sur Android, et Outlook sur le web et Outlook.com. Les fonctionnalités les plus récentes ne sont pas toutes prises en charge dans tous les [clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) à la fois. Reportez-vous aux articles et références API relatives à ces fonctionnalités pour savoir dans quels applications elles peuvent ou non être prises en charge.

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
