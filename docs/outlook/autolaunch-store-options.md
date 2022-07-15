---
title: Options de liste AppSource pour votre complément Outlook basé sur les événements
description: Découvrez les options de liste AppSource disponibles pour votre complément Outlook qui implémente l’activation basée sur les événements.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d8d2c2e9960d2aef2d32ede6e20eb5f1db125a6c
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797679"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>Options de liste AppSource pour votre complément Outlook basé sur les événements

À l’heure actuelle, les compléments doivent être déployés par les administrateurs d’une organisation pour que les utilisateurs finaux puissent accéder à la fonctionnalité basée sur les événements. Nous limitons l’activation basée sur les événements si l’utilisateur final a acquis le complément directement à partir d’AppSource. Par exemple, si le complément Contoso inclut le `LaunchEvent` point d’extension avec au moins un point défini `LaunchEvent Type` sous le `LaunchEvents` nœud, l’appel automatique du complément se produit uniquement si le complément a été installé pour l’utilisateur final par l’administrateur de son organisation. Sinon, l’appel automatique du complément est bloqué. Consultez l’extrait suivant d’un exemple de manifeste de complément.

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

Un utilisateur final ou un administrateur peut acquérir des compléments via AppSource ou l’Office Store in-app. Si le scénario ou le flux de travail principal de votre complément nécessite une activation basée sur les événements, vous souhaiterez peut-être restreindre vos compléments disponibles pour le déploiement administrateur. Pour activer cette restriction, nous pouvons fournir des URL de code de version d’évaluation. Grâce aux codes de version d’évaluation, seuls les utilisateurs finaux disposant de ces URL spéciales peuvent accéder à la liste. Voici un exemple d’URL.

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

Les utilisateurs et les administrateurs ne peuvent pas rechercher explicitement un complément par son nom dans AppSource ou l’Office Store dans l’application lorsqu’un code de version d’évaluation est activé pour celui-ci. En tant que créateur du complément, vous pouvez partager en privé ces codes de vol avec les administrateurs de l’organisation pour le déploiement du complément.

> [!NOTE]
> Bien que les utilisateurs finaux puissent installer le complément à l’aide d’un code de version d’évaluation, le complément n’inclut pas l’activation basée sur les événements.

## <a name="specify-a-flight-code"></a>Spécifier un code de vol

Pour spécifier le code de version d’évaluation souhaité pour votre complément, partagez ces informations dans les **notes de certification** lorsque vous publiez votre complément. _**Important** :_ Les codes de vol respectent la casse.

![Capture d’écran montrant l’exemple de demande de code de version d’évaluation dans l’écran Notes pour la certification pendant le processus de publication.](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a>Déployer un complément avec du code de version d’évaluation

Une fois les codes de vol définis, vous recevrez l’URL de l’équipe de certification des applications. Vous pouvez ensuite partager l’URL avec les administrateurs en privé.

Pour déployer le complément, l’administrateur peut effectuer les étapes suivantes.

- Connectez-vous à admin.microsoft.com ou AppSource.com avec votre compte d’administrateur Microsoft 365. Si l’authentification unique (SSO) est activée pour le complément, des informations d’identification d’administrateur général sont nécessaires.
- Ouvrez l’URL du code de vol dans un navigateur web.
- Dans la page de liste des compléments, **sélectionnez Obtenir maintenant**. Vous devez être redirigé vers le portail d’application intégré.

## <a name="unrestricted-appsource-listing"></a>Description d’AppSource sans restriction

Si votre complément n’utilise pas l’activation basée sur les événements pour les scénarios critiques (autrement dit, votre complément fonctionne bien sans appel automatique), envisagez de répertorier votre complément dans AppSource sans codes de vol spéciaux. Si un utilisateur final obtient votre complément à partir d’AppSource, l’activation automatique ne se produit pas pour l’utilisateur. Toutefois, ils peuvent utiliser d’autres composants de votre complément, tels qu’un volet Office ou une commande de fonction.

> [!IMPORTANT]
> Il s’agit d’une restriction temporaire. À l’avenir, nous prévoyons d’activer l’activation de compléments basés sur des événements pour les utilisateurs finaux qui acquièrent directement votre complément.

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>Mettre à jour les compléments existants pour inclure l’activation basée sur les événements

Vous pouvez mettre à jour votre complément existant pour inclure l’activation basée sur les événements, puis le soumettre à nouveau pour validation et décider si vous souhaitez une liste AppSource restreinte ou non restreinte.

Une fois le complément mis à jour approuvé, les administrateurs de l’organisation qui ont déjà déployé le complément recevront un message de mise à jour dans la section **Applications intégrées** du Centre d’administration. Le message informe l’administrateur des modifications apportées à l’activation basée sur les événements. Une fois que l’administrateur a accepté les modifications, la mise à jour est déployée pour les utilisateurs finaux.

![Capture d’écran de la notification de mise à jour d’application sur l’écran « Applications intégrées ».](../images/outlook-deploy-update-notification.png)

Pour les utilisateurs finaux qui ont installé le complément seuls, la fonctionnalité d’activation basée sur les événements ne fonctionnera pas même après la mise à jour du complément.

## <a name="admin-consent-for-installing-event-based-add-ins"></a>Administration consentement pour l’installation de compléments basés sur des événements

Chaque fois qu’un complément basé sur les événements est déployé à partir de l’écran **Applications intégrées** , l’administrateur obtient des détails sur les fonctionnalités d’activation basées sur les événements du complément dans l’Assistant déploiement. Les détails s’affichent dans la section **Autorisations et fonctionnalités de l’application** . L’administrateur doit voir tous les événements où le complément peut s’activer automatiquement.

![Capture d’écran de l’écran « Accepter les demandes d’autorisations » lors du déploiement d’une nouvelle application.](../images/outlook-deploy-accept-permissions-requests.png)

De même, lorsqu’un complément existant est mis à jour vers des fonctionnalités basées sur des événements, l’administrateur voit un état « Mise à jour en attente » sur le complément. Le complément mis à jour est déployé uniquement si l’administrateur accepte les modifications indiquées dans la section **Autorisations et fonctionnalités** de l’application, y compris l’ensemble d’événements où le complément peut s’activer automatiquement.

Chaque fois que vous ajoutez des nouveautés `LaunchEvent Type` à votre complément, les administrateurs voient le flux de mise à jour dans le portail d’administration et doivent donner leur consentement pour des événements supplémentaires.

![Capture d’écran du flux « Mises à jour » lors du déploiement d’une application mise à jour.](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
