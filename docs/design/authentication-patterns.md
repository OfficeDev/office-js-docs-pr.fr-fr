---
title: Conception de lignes directrices relatives à l’authentification pour les compléments Office
ms.date: 07/30/2020
description: Découvrez comment concevoir visuellement une page de connexion ou d’inscription dans un complément Office.
localization_priority: Normal
ms.openlocfilehash: e75ab26f820e53ca0379467fba559ebcb727ee3a
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132143"
---
# <a name="authentication-patterns"></a>Modèles d’authentification

Des compléments peuvent exiger que des utilisateurs se connectent ou s’inscrivent pour pouvoir accéder à certaines fonctions et fonctionnalités. Des zones de saisie pour le nom d’utilisateur et le mot de passe, ou des boutons qui lancent des flux d’informations d’identification tiers sont des contrôles d’interface courants dans les expériences d’authentification. Une expérience d’authentification simple et efficace est une première étape importante pour inciter des utilisateurs à commencer à utiliser votre complément.

## <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:----|:----|
|Avant la connexion, décrivez la valeur de votre complément, ou montrez les fonctionnalités sans exiger de compte. |Attendez que des utilisateurs se connectent sans comprendre la valeur et les avantages de votre complément.|
|Guidez les utilisateurs dans les flux d’authentification à l’aide d’un bouton principal bien visible sur chaque écran. |Attirez l’attention sur les tâches secondaires et tertiaires avec des boutons et appels à l’action concurrents.|
|Utilisez des libellés de boutons clairs décrivant des tâches spécifiques telles que « Se connecter » ou « Créer un compte ». |Utilisez des étiquettes de boutons vagues telles que « Soumettre » ou « Commencer » pour guider les utilisateurs tout au long des flux d’authentification.|
|Utilisez une boîte de dialogue pour attirer l’attention des utilisateurs sur les formulaires d’authentification. |Enrichissez votre volet des tâches avec une première expérience d’exécution et des formulaires d’authentification.|
|Trouvez de petites efficacités dans le flux, comme la focalisation automatique sur des zones de saisie. |Ajoutez des étapes superflues à l’interaction, telles que l’obligation pour les utilisateurs de cliquer dans des champs de formulaire.|
|Offrir aux utilisateurs la possibilité de se déconnecter et de se réauthentifier. |Obligez les utilisateurs à se désinstaller pour changer d’identité.|

## <a name="authentication-flow"></a>Flux d’authentification

1. Mise en place de la première expérience d’exécution : positionnez votre bouton de connexion en tant qu’appel à l’action clair dans l’interface de première exécution de votre complément.

    ![Capture d’écran illustrant un volet Office de complément dans une application Office](../images/add-in-fre-value-placemat.png)

1. Boîte de dialogue pour le choix de fournisseur d’identité : affichez une liste claire de fournisseurs d’identité, dont un formulaire de nom d’utilisateur et de mot de passe, le cas échéant. Il se peut que l’interface utilisateur de votre complément se bloque lorsque la boîte de dialogue d’authentification est ouverte.

    ![Capture d’écran illustrant la boîte de dialogue choix du fournisseur d’identité dans une application Office](../images/add-in-auth-choices-dialog.png)

1. Connexion au fournisseur d’identité : le fournisseur d’identité aura sa propre interface utilisateur. Microsoft Azure Active Directory permet de personnaliser les pages de connexion et de panneau d’accès pour obtenir une apparence cohérente avec votre service. [En savoir plus](/azure/active-directory/fundamentals/customize-branding).

    ![Capture d’écran illustrant la boîte de dialogue de connexion du fournisseur d’identité dans une application Office](../images/add-in-auth-identity-sign-in.png)

1. Progression : indiquez la progression du chargement des paramètres et de l’interface utilisateur.

    ![Capture d’écran illustrant une boîte de dialogue avec un indicateur de progression dans une application Office](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> Lorsque vous utilisez le service d’identité de Microsoft vous avez la possibilité d’utiliser un bouton de connexion personnalisable à l’aide de thèmes lumineux et sombres. En savoir plus.

## <a name="single-sign-on-authentication-flow"></a>Flux d’authentification Sign-On unique

> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur la prise en charge de l’authentification unique, voir [ensembles Requirements sets](../reference/requirement-sets/identity-api-requirement-sets.md). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, voir [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Utilisez l’authentification unique pour une expérience utilisateur plus fluide. L’identité de l’utilisateur au sein d’Office (un compte Microsoft ou une identité Microsoft 365) est utilisée pour se connecter à votre complément. Par conséquent, les utilisateurs se connectent une seule fois. Cela permet d’éliminer les frictions dans l’expérience, en facilitant la prise en main pour vos clients.

1. Lorsqu’un complément est en cours d’installation, un utilisateur verra une fenêtre de consentement semblable à celle-ci :

    ![Capture d’écran illustrant la fenêtre de consentement dans une application Office lorsqu’un complément est en cours d’installation](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > L’éditeur du complément contrôle le logo, les chaînes et les étendues d’autorisation inclus dans la fenêtre de consentement. L’interface utilisateur est préconfigurée par Microsoft.

1. Le complément est chargé une fois que l’utilisateur a accepté. Il peut extraire et afficher les informations personnalisées nécessaires de l’utilisateur.

    ![Capture d’écran illustrant une application Office avec des boutons de complément affichés dans le ruban](../images/add-in-ribbon.png)

## <a name="see-also"></a>Voir aussi

- En savoir plus sur [le développement de compléments d’authentification unique](../develop/sso-in-office-add-ins.md)
