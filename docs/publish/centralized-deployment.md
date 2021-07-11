---
title: Publier des Office à l’aide d’un déploiement centralisé via le Centre d’administration Microsoft 365
description: Découvrez comment utiliser le déploiement centralisé pour déployer des add-ins internes ainsi que des add-ins fournis par des isvs.
ms.date: 03/22/2021
localization_priority: Normal
ms.openlocfilehash: b57e21f177fe66f03985ce6baee4d9eeda75d8bd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348789"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>Publier des Office à l’aide d’un déploiement centralisé via le Centre d’administration Microsoft 365

Le Centre d’administration Microsoft 365 permet aux administrateurs de déployer facilement des Office aux utilisateurs et groupes au sein de leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez utiliser le déploiement centralisé pour déployer des compléments internes, ainsi que des compléments fournis par des éditeurs de logiciels indépendants.

Le Centre d’administration Microsoft 365 prend actuellement en charge les scénarios suivants.

- Déploiement centralisé de nouveaux compléments et de ceux mis à jour pour des utilisateurs, des groupes ou une organisation.
- Déploiement sur plusieurs plateformes clientes, notamment Windows, Mac et le web. Pour Outlook, le déploiement vers iOS et Android est également pris en charge. (Toutefois, si l’installation des Excel, Outlook, Word et PowerPoint par les utilisateurs sur iPad est prise en charge, le déploiement centralisé vers iPad n’est pas pris en charge.) 
- Déploiement en anglais et pour les clients du monde entier.
- Déploiement de compléments hébergés sur le cloud.
- Déploiement de compléments hébergés au sein d’un pare-feu.
- Déploiement de compléments AppSource.
- Installation automatique d’un complément pour les utilisateurs au lancement de l’application Office.
- Suppression automatique d’un complément pour les utilisateurs si l’administrateur désactive ou supprime le complément, ou si les utilisateurs sont supprimés d’Azure Active Directory ou d’un groupe auprès duquel le complément a été déployé.

Le déploiement centralisé est la façon recommandée pour un administrateur Microsoft 365 de déployer des add-ins Office au sein d’une organisation, à condition que l’organisation réponde à toutes les conditions requises pour utiliser le déploiement centralisé. Pour plus d’informations sur la façon de déterminer si votre organisation peut utiliser le déploiement centralisé, voir Déterminer si le déploiement centralisé des [modules complémentaires](/office365/admin/manage/centralized-deployment-of-add-ins)fonctionne pour Microsoft 365 organisation.

> [!NOTE]
> Dans un environnement local sans connexion à Microsoft 365, ou pour déployer des applications SharePoint ou Office qui ciblent Office 2013, utilisez un catalogue d’applications [SharePoint.](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) Pour déployer des compléments COM/VSTO, utilisez ClickOnce ou Windows Installer, comme décrit dans la rubrique [Déploiement d’une solution Office](/visualstudio/vsto/deploying-an-office-solution).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Approche recommandée pour le déploiement des compléments Office

Envisagez de déployer des compléments Office dans une approche progressive pour vous assurer que le déploiement se déroule sans problème. Nous vous recommandons l’plan suivant.

1. Déployez le complément auprès d’un petit groupe de parties prenantes et de membres du service informatique. Si le déploiement réussit, passez à l’étape 2.

1. Déployez le complément auprès d’un groupe plus important de membres dans l’organisation qui utilisera le complément. Si le déploiement réussit, passez à l’étape 3.

1. Déployez le complément auprès du groupe entier de membres qui utilisera le complément.

Selon la taille de l’audience cible, vous pouvez ajouter des étapes à cette procédure ou en supprimer.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Publication d’un complément Office via le déploiement centralisé

Avant de commencer, confirmez que votre organisation répond à toutes les conditions requises pour utiliser le déploiement centralisé, comme décrit dans Determine [if Centralized Deployment of add-ins works for your Microsoft 365 organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

Si votre organisation répond à toutes les exigences, effectuez les étapes suivantes pour publier un Office par le biais d’un déploiement centralisé.

1. Connectez-vous Microsoft 365 avec votre compte d’éducation ou de travail.
1. Sélectionnez l’icône du lanceur d’applications située en haut à gauche et choisissez **Administrateur**.
1. Dans le menu de navigation, **sélectionnez Afficher plus,** puis **Paramètres**  >  **applications intégrées.**
1. En haut de la page, choisissez **Les modules.**
1. Si un message s’affiche en haut de la page pour annoncer la nouvelle Centre d’administration Microsoft 365, sélectionnez le message pour aller à l’aperçu du Centre d’administration (voir à propos de la [Centre d’administration Microsoft 365](/microsoft-365/admin/admin-overview/about-the-admin-center)).
1. Choisissez **Déployer un complément** en haut de la page.
1. Choisissez **Suivant** après avoir consulté la configuration requise.
1. Choisissez l’une des options suivantes dans la page **Déploiement** centralisé.

    - **Je veux ajouter un complément à partir de l’Office Store**.
    - **J’ai le fichier manifeste (.xml) sur cet appareil**. Pour cette option, sélectionnez **Parcourir** afin de recherche le fichier manifeste (.xml) que vous voulez utiliser.
    - **J’ai une URL pour le fichier manifeste**. Pour cette option, entrez l’URL du manifeste dans le champ disponible.

    ![Nouvelle boîte Add-In dialogue dans Centre d’administration Microsoft 365.](../images/new-add-in.png)

1. Si vous avez sélectionné l’option d’ajout d’un complément à partir de l’Office Store, sélectionnez le complément. Vous pouvez afficher les compléments disponibles via l’une des catégories suivantes : **Suggestions**, **Évaluation** ou **Nom**. Vous ne pouvez ajouter que des compléments gratuits de l’Office Store. L’ajout de compléments payants n’est pas actuellement pris en charge.

    > [!NOTE]
    > Avec l’option Office Store, les mises à jour et améliorations du complément sont automatiquement disponibles pour les utilisateurs sans intervention de votre part.

    ![Sélectionnez une boîte de dialogue de Centre d’administration Microsoft 365.](../images/select-an-add-in.png)

1. Choisissez **Continuer** après avoir passé en revue les détails du module, la politique de confidentialité et les termes du contrat de licence.

    ![Page de la sélection du Centre d’administration Microsoft 365.](../images/selected-add-in-admin-center.png)

1. Dans la page **Attribuer des utilisateurs,** sélectionnez **Tout le** monde, **Utilisateurs/Groupes** spécifiques ou **Moi seul.** Utilisez la zone de recherche pour trouver les utilisateurs et groupes vers lesquels vous voulez déployer le complément. Pour Outlook, vous pouvez également choisir la méthode de déploiement **Fixe,** **Disponible** ou **Facultatif**.

    ![Gérer les personnes qui disposent de la méthode d’accès et de déploiement dans Centre d’administration Microsoft 365.](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Les add-ins qui utilisent l' [sign-on unique (SSO)](../develop/sso-in-office-add-ins.md) invitent l’administrateur à consentir aux étendues répertoriées dans le manifeste du add-in.  Si le même service de backing est utilisé dans plusieurs applications (le même ID d’application Azure est utilisé avec ssO dans différents modules), les étendues de chaque application sont invités à donner leur consentement à chaque déploiement. Cette page affiche également la liste des autorisations dont le module a besoin.

1. Lorsque vous avez terminé, choisissez **Déployer.** Ce processus peut prendre jusqu’à trois minutes. Ensuite, terminez la procédure en appuyant sur **Suivant**. Vous voyez maintenant votre application avec d’autres Office applications.

    > [!NOTE]
    > Lorsqu’un administrateur choisit **Déployer,** le consentement est donné pour tous les utilisateurs.

    ![Liste des applications dans Centre d’administration Microsoft 365.](../images/citations.png)

> [!TIP]
> lorsque vous déployez un nouveau complément vers des utilisateurs et/ou des groupes de votre organisation, pensez à leur envoyer un courrier décrivant quand et comment utiliser le complément et incluez des liens vers un contenu pertinent de l’aide ou du forum aux questions, ou d’autres ressources de support.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Éléments à prendre en compte lors de l’octroi de l’accès à un complément

Les administrateurs peuvent affecter un complément à tout le monde ou à des utilisateurs et/ou groupes spécifiques au sein de l’organisation. La liste suivante décrit les implications de chaque option.

- **Tout le monde** : comme son nom l’indique, cette option affecte le complément à tous les utilisateurs du client. Utilisez-la avec parcimonie et uniquement pour les compléments qui sont réellement universels pour l’ensemble de votre organisation.

- **Utilisateurs** : si vous affectez un complément à un utilisateur particulier, vous devez mettre à jour les paramètres de déploiement centralisé pour le complément chaque fois que vous souhaitez l’affecter à des utilisateurs supplémentaires. De même, vous devez mettre à jour les paramètres de déploiement centralisé pour le complément chaque fois que vous souhaitez supprimer l’accès d’un utilisateur au complément.

- **Groupes** : si vous affectez un complément à un groupe, le complément est automatiquement affecté aux utilisateurs ajoutés au groupe. De même, quand un utilisateur est supprimé d’un groupe, il perd l’accès au complément. Dans les deux cas, aucune action supplémentaire n’est requise de la part Microsoft 365 administrateur.

En général, pour faciliter la maintenance, nous vous recommandons d’affecter des compléments à l’aide de groupes. Toutefois, dans les situations où vous souhaitez restreindre l’accès au complément à un très petit nombre d’utilisateurs, il peut être plus pratique d’affecter le complément à des utilisateurs spécifiques.

## <a name="add-in-states"></a>États de complément

Le tableau suivant décrit les différents états qui s’appliquent à un complément.

|État|Comment l’état se produit|Impact|
|-----|--------------------|------|
|**Actif**|Un administrateur a chargé le complément et l’a affecté à des utilisateurs et/ou groupes.|Les utilisateurs et/ou groupes auxquels le complément est affecté voient celui-ci dans les clients Office concernés.|
|**Désactivé**|Un administrateur a désactivé le complément.|Les utilisateurs et/ou groupes auxquels le complément est affecté ne peuvent plus y accéder. Si l’état du complément est modifié, passant de **Désactivé** à **Actif**, les utilisateurs et groupes y ont de nouveau accès.|
|**Deleted**|Un administrateur a supprimé le complément.|Les utilisateurs et/ou groupes auxquels le complément est affecté ne peuvent plus y accéder.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Mise à jour des compléments Office publiés via un déploiement centralisé

Après la publication d’un Office par le biais d’un déploiement centralisé, toutes les modifications apportées à l’application web du module sont automatiquement disponibles pour tous les utilisateurs une fois ces modifications implémentées dans l’application web. Les modifications apportées au fichier manifeste [XML](../develop/add-in-manifests.md) d’un add-in pour, par exemple, mettre à jour l’icône, le texte ou les commandes du module, se produisent comme suit :

- **Application** métier : si un administrateur a chargé explicitement un fichier manifeste (à partir de son appareil ou en pointant vers une URL) lors de l’implémentation du déploiement centralisé via le Centre d’administration Microsoft 365, il doit télécharger un nouveau fichier manifeste qui contient les modifications souhaitées. Une fois que le fichier manifeste a été chargé, le complément est mis à jour au démarrage suivant des applications Office concernées.

  > [!NOTE]
  > Un administrateur n’a pas besoin de supprimer un add-in LOB pour effectuer une mise à jour. Dans la section Des add-ins, l’administrateur peut simplement choisir le add-in LOB et appeler cette fonctionnalité en appuyant sur le bouton Mettre à jour du module de mise à jour présent dans le coin inférieur droit. 
  >
  > ![Screenshot shows the Update add-in dialog in Centre d’administration Microsoft 365.](../images/update-add-in-admin-center.png)

- **Office** Store : si un administrateur a sélectionné un module dans le Office Store lors de l’implémentation d’un déploiement centralisé via le Centre d’administration Microsoft 365, et que les mises à jour de ce dernier sont mises à jour dans le Office Store, celui-ci sera mis à jour ultérieurement via le déploiement centralisé. Le flux des mises à jour des add-ins du Store pour tous les utilisateurs peut prendre jusqu’à 24 heures. Après cette durée, la prochaine fois que les applications Office redémarrent pour ces utilisateurs, le add-in est mis à jour. Les utilisateurs peuvent également déclencher une actualisation manuelle pour obtenir la dernière version du add-in store en sélectionnant Insérer des modules d’ajout  >  **d’onglets**  >  **Admin Managed Tab Hit**  >  **Refresh**.

## <a name="end-user-experience-with-add-ins"></a>Expérience des utilisateurs finaux avec les compléments

Une fois qu’un complément a été publié via un déploiement centralisé, les utilisateurs finaux peuvent commencer à l’utiliser sur toutes les plateformes prises en charge par le complément.

Si le complément prend en charge les commandes de complément, celles-ci apparaissent dans le ruban de l’application Office pour tous les utilisateurs vers lesquels le complément est déployé. Dans l’exemple suivant, la commande **Recherche de citation** apparaît dans le ruban pour le complément **Citations**.

![Screenshot shows a section of the application Office ribbon with the Search Citation command highlighted in the Citations add-in.](../images/search-citation.png)

Si le complément ne prend pas en charge les commandes de complément, les utilisateurs peuvent l’ajouter à leur application Office en procédant comme suit :

1. Dans Word 2016 ou version ultérieure, Excel 2016 ou version ultérieure ou PowerPoint 2016 ou version ultérieure, choisissez **Insertion** > **Mes compléments**.
1. Sélectionnez l’onglet **Géré par l’administrateur** dans la fenêtre du complément.
1. Choisissez le complément, puis **Ajouter**.

    ![Capture d’écran illustrant l’onglet Géré par l’administrateur de la page Compléments Office d’une application Office. Le complément Citations apparaît sur l’onglet.](../images/office-add-ins-admin-managed.png)

Toutefois, pour Outlook 2016 ou version ultérieure, les utilisateurs peuvent procéder comme suit :

1. Dans Outlook, Choisissez **Accueil** > **Store**.
1. Sélectionnez l’élément **Géré par l’administrateur** dans la fenêtre du complément.
1. Choisissez le complément, puis **Ajouter**.

    ![Capture d’écran montrant la zone Géré par l’administrateur de la page Store de l’application Outlook.](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>Voir aussi

- [Déterminer si le déploiement centralisé des compléments fonctionne avec votre organisation Microsoft 365](/office365/admin/manage/centralized-deployment-of-add-ins)
