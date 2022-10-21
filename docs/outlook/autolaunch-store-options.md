---
title: Options de liste AppSource pour votre complément Outlook basé sur les événements
description: Découvrez les options de liste AppSource disponibles pour votre complément Outlook qui implémente l’activation basée sur les événements.
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: b8908fde484186fdb9bea9cda4520358278712f6
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664664"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>Options de liste AppSource pour votre complément Outlook basé sur les événements

Les compléments doivent être déployés par les administrateurs d’une organisation pour que les utilisateurs finaux puissent accéder à la fonctionnalité [d’activation basée sur les événements](autolaunch.md) . L’activation basée sur les événements est limitée si l’utilisateur final a acquis le complément directement à partir [d’AppSource](https://appsource.microsoft.com). Par exemple, si le complément Contoso inclut une fonction basée sur des événements, l’appel automatique du complément se produit uniquement si le complément a été installé pour l’utilisateur final par l’administrateur de son organisation. Sinon, l’appel automatique du complément est bloqué.

Un utilisateur final ou un administrateur peut acquérir des compléments via AppSource ou l’Office Store dans l’application. Si le scénario principal ou le workflow de votre complément nécessite une activation basée sur les événements, vous pouvez limiter votre complément au déploiement administrateur uniquement. Pour activer cette restriction, nous pouvons fournir des URL de code de version d’évaluation pour les compléments dans AppSource. Grâce aux codes de vol, seuls les utilisateurs finaux disposant de ces URL spéciales peuvent accéder à la liste. Voici un exemple d’URL.

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

Les utilisateurs et les administrateurs ne peuvent pas rechercher explicitement un complément par son nom dans AppSource ou l’Office Store in-app lorsqu’un code de version d’évaluation est activé pour celui-ci. En tant que créateur du complément, vous pouvez partager ces codes de version d’évaluation en privé avec les administrateurs de l’organisation pour le déploiement du complément.

> [!NOTE]
> Bien que les utilisateurs finaux puissent installer le complément à l’aide d’un code de version d’évaluation, le complément n’inclut pas l’activation basée sur les événements.

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="specify-a-flight-code"></a>Spécifier un code de version d’évaluation

Pour spécifier le code de version d’évaluation de votre complément, partagez-le dans les **Notes pour la certification** lorsque vous publiez votre complément. **Important** : les codes de vol respectent la casse.

![Exemple de demande de code de version d’évaluation dans l’écran Notes pour la certification pendant le processus de publication.](../images/outlook-publish-notes-for-certification.png)

## <a name="deploy-add-in-with-flight-code"></a>Déployer le complément avec le code de version d’évaluation

Une fois les codes de version d’évaluation définis, vous recevez l’URL de l’équipe de certification de l’application. Vous pouvez ensuite partager l’URL avec les administrateurs en privé.

Pour déployer le complément, l’administrateur peut effectuer les étapes suivantes.

- Connectez-vous à admin.microsoft.com ou AppSource.com avec votre compte d’administrateur Microsoft 365. Si l’authentification unique (SSO) est activée pour le complément, les informations d’identification de l’administrateur général sont nécessaires.
- Ouvrez l’URL du code de version d’évaluation dans un navigateur web.
- Dans la page de liste du complément, sélectionnez **Obtenir maintenant**. Vous devez être redirigé vers le portail d’application intégré.

## <a name="unrestricted-appsource-listing"></a>Liste AppSource illimitée

Si votre complément n’utilise pas l’activation basée sur les événements pour les scénarios critiques (autrement dit, votre complément fonctionne correctement sans appel automatique), envisagez de répertorier votre complément dans AppSource sans codes de version d’évaluation spéciaux. Si un utilisateur final obtient votre complément à partir d’AppSource, l’activation automatique ne se produit pas pour l’utilisateur. Toutefois, ils peuvent utiliser d’autres composants de votre complément, tels qu’un volet office ou une commande de fonction.

> [!IMPORTANT]
> Il s’agit d’une restriction temporaire. À l’avenir, nous prévoyons d’activer l’activation de complément basée sur les événements pour les utilisateurs finaux qui acquièrent directement votre complément.

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>Mettre à jour les compléments existants pour inclure l’activation basée sur les événements

Vous pouvez mettre à jour votre complément existant pour inclure l’activation basée sur les événements, puis le soumettre à nouveau pour validation et décider si vous souhaitez une liste AppSource restreinte ou illimitée.

Une fois le complément mis à jour approuvé, les administrateurs d’organisation qui ont précédemment déployé le complément reçoivent un message de mise à jour dans la section **Applications intégrées** du Centre d’administration. Le message informe l’administrateur des modifications apportées à l’activation basée sur les événements. Une fois que l’administrateur accepte les modifications, la mise à jour est déployée sur les utilisateurs finaux.

![Notifications de mise à jour d’application sur l’écran « Applications intégrées ».](../images/outlook-deploy-update-notification.png)

Pour les utilisateurs finaux qui ont installé le complément eux-mêmes, la fonctionnalité d’activation basée sur les événements ne fonctionnera pas même après la mise à jour du complément.

## <a name="admin-consent-for-installing-event-based-add-ins"></a>Administration consentement pour l’installation des compléments basés sur les événements

Chaque fois qu’un complément basé sur les événements est déployé à partir de l’écran **Applications intégrées** , l’administrateur obtient des détails sur les fonctionnalités d’activation basée sur les événements du complément dans l’Assistant déploiement. Les détails s’affichent dans la section **Autorisations et fonctionnalités des** applications. L’administrateur doit voir tous les événements dans lesquels le complément peut s’activer automatiquement.

![Écran « Accepter les demandes d’autorisations » lors du déploiement d’une nouvelle application.](../images/outlook-deploy-accept-permissions-requests.png)

De même, lorsqu’un complément existant est mis à jour vers des fonctionnalités basées sur les événements, l’administrateur voit l’état « Mise à jour en attente » sur le complément. Le complément mis à jour est déployé uniquement si l’administrateur accepte les modifications indiquées dans la section **Autorisations et fonctionnalités de l’application** , y compris l’ensemble des événements où le complément peut s’activer automatiquement.

Chaque fois que vous ajoutez une nouvelle fonction d’activation basée sur les événements à votre complément, les administrateurs voient le flux de mise à jour dans le portail d’administration et doivent donner leur consentement pour d’autres événements.

![Flux « Mises à jour » lors du déploiement d’une application mise à jour.](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
