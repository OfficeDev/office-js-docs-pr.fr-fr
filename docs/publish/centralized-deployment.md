---
title: Publication de compléments Office à l’aide du déploiement centralisé via le centre d’administration Office 365
description: Découvrez comment utiliser le déploiement centralisé pour déployer des compléments internes ainsi que des compléments fournis par les éditeurs de logiciels indépendants.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 1410409fbd86be13da4551b2f140bd41fdaebbbf
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778675"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-office-365-admin-center"></a>Publication de compléments Office à l’aide du déploiement centralisé via le centre d’administration Office 365

The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.

Le centre d’administration Office 365 prend actuellement en charge les scénarios suivants.

- Déploiement centralisé de nouveaux compléments et de ceux mis à jour pour des utilisateurs, des groupes ou une organisation.
- Déploiement sur plusieurs plateformes client, y compris Windows, Mac et le Web. Pour Outlook, le déploiement sur iOS et Android est également pris en charge. (Toutefois, pendant que l’installation utilisateur des compléments Excel, Outlook, Word et PowerPoint sur iPad est prise en charge, le déploiement centralisé sur iPad n’est **pas** pris en charge.)
- Déploiement en anglais et pour les clients du monde entier.
- Déploiement de compléments hébergés sur le cloud.
- Déploiement de compléments hébergés au sein d’un pare-feu.
- Déploiement de compléments AppSource.
- Installation automatique d’un complément pour les utilisateurs au lancement de l’application Office.
- Suppression automatique d’un complément pour les utilisateurs si l’administrateur désactive ou supprime le complément, ou si les utilisateurs sont supprimés d’Azure Active Directory ou d’un groupe auprès duquel le complément a été déployé.

Le déploiement centralisé est la méthode recommandée pour le déploiement de compléments Office par un administrateur Office 365 dans une organisation, à condition que l’organisation remplisse toutes les conditions d’utilisation du déploiement centralisé. Pour savoir comment déterminer si votre organisation peut utiliser un déploiement centralisé, reportez-vous à [Déterminer si un déploiement centralisé de compléments est approprié pour votre organisation Office 365](/office365/admin/manage/centralized-deployment-of-add-ins).

> [!NOTE]
> In an on-premises environment with no connection to Office 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Approche recommandée pour le déploiement des compléments Office

Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:

1. Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.

2. Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.

3. Déployez le complément auprès du groupe entier de membres qui utilisera le complément.

Selon la taille de l’audience cible, vous pouvez ajouter des étapes à cette procédure ou en supprimer.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Publication d’un complément Office via le déploiement centralisé

Avant de commencer, vérifiez que votre organisation est conforme à toutes les conditions d’utilisation du déploiement centralisé, comme décrit dans la rubrique [Déterminer si un déploiement centralisé de compléments est approprié pour votre organisation Office 365](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

Si votre organisation répond à toutes les conditions requises, procédez comme suit pour publier un complément Office via un déploiement centralisé :

1. Connectez-vous à Office 365 avec votre compte professionnel ou scolaire.
2. Sélectionnez l’icône du lanceur d’applications située en haut à gauche et choisissez **Administrateur**.
3. Dans le menu de navigation, appuyez sur **Afficher plus**, puis choisissez **Paramètres** > **Services et compléments**.
4. Si un message annonçant le nouveau Centre d’administration Office 365 apparaît en haut de la page, cliquez dessus pour accéder à la préversion du Centre d’administration (reportez-vous à l’article [À propos du Centre d’administration Office 365](/microsoft-365/admin/admin-overview/about-the-admin-center)).
5. Choisissez **Déployer un complément** en haut de la page.
6. Choisissez **Suivant** après avoir consulté la configuration requise.
7. Dans la page **Déploiement centralisé**, choisissez l’une des options suivantes :

    - **Je veux ajouter un complément à partir de l’Office Store**.
    - **I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![Boîte de dialogue Nouveau complément dans le Centre d’administration Office 365](../images/new-add-in.png)

8. Si vous avez sélectionné l’option d’ajout d’un complément à partir de l’Office Store, sélectionnez le complément. Vous pouvez afficher les compléments disponibles via l’une des catégories suivantes : **Suggestions**, **Évaluation** ou **Nom**. Vous ne pouvez ajouter que des compléments gratuits de l’Office Store. L’ajout de compléments payants n’est pas actuellement pris en charge.

    > [!NOTE]
    > Avec l’option Office Store, les mises à jour et améliorations du complément sont automatiquement disponibles pour les utilisateurs sans intervention de votre part.

    ![Sélection d’une boîte de dialogue de complément dans le centre d’administration Office 365](../images/select-an-add-in.png)

9. Choisissez **Continuer** après avoir vérifié les détails du complément, la politique de confidentialité et les termes de la licence.

    ![Page de complément sélectionnée dans le centre d’administration Office 365](../images/selected-add-in-admin-center.png)

10. Sur la **page attribuer des utilisateurs** , choisissez **tout le monde**, **utilisateurs/groupes spécifiques**ou **moi seul**. Utilisez la zone de recherche pour trouver les utilisateurs et groupes vers lesquels vous voulez déployer le complément. Pour les compléments Outlook, vous pouvez également choisir la méthode de déploiement **fixe**, **disponible**ou **facultative**.

    ![Gérer les personnes ayant accès et méthode de déploiement dans le centre d’administration Office 365](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Un système d’[authentification unique (SSO) ](../develop/sso-in-office-add-ins.md) pour les compléments est actuellement en préversion, qui ne doit pas être utilisé pour des compléments en production.Lors du déploiement d’un complément utilisant une authentification unique, les utilisateurs et les groupes affectés sont également partagés avec des compléments partageant le même ID d’application Azure. Les modifications apportées aux affectations d’utilisateurs sont également appliquées à ces compléments. Les compléments connexes sont affichés sur cette page. Uniquement pour les compléments d’authentification unique, cette page affiche la liste des autorisations Microsoft Graph requises.

11. Lorsque vous avez terminé, choisissez **déployer**. Ce processus peut prendre jusqu’à trois minutes. Ensuite, terminez la procédure en appuyant sur **Suivant**. Votre complément apparaît à présent avec d’autres applications dans Office 365.

    > [!NOTE]
    > Lorsqu’un administrateur choisit **Deploy**, le consentement est donné pour tous les utilisateurs.

    ![Liste des applications dans le Centre d’administration Office 365](../images/citations.png)

> [!TIP]
> lorsque vous déployez un nouveau complément vers des utilisateurs et/ou des groupes de votre organisation, pensez à leur envoyer un courrier décrivant quand et comment utiliser le complément et incluez des liens vers un contenu pertinent de l’aide ou du forum aux questions, ou d’autres ressources de support.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Éléments à prendre en compte lors de l’octroi de l’accès à un complément

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:

- **Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

- **Utilisateurs** : si vous affectez un complément à un utilisateur particulier, vous devez mettre à jour les paramètres de déploiement centralisé pour le complément chaque fois que vous souhaitez l’affecter à des utilisateurs supplémentaires. De même, vous devez mettre à jour les paramètres de déploiement centralisé pour le complément chaque fois que vous souhaitez supprimer l’accès d’un utilisateur au complément.

- **Groupes** : si vous affectez un complément à un groupe, le complément est automatiquement affecté aux utilisateurs ajoutés au groupe. De même, quand un utilisateur est supprimé d’un groupe, il perd l’accès au complément. Dans les deux cas, aucune action supplémentaire n’est requise de votre part en tant qu’administrateur Office 365.

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.

## <a name="add-in-states"></a>États de complément

Le tableau suivant décrit les différents états qui s’appliquent à un complément.

|État|Comment l’état se produit|Impact|
|-----|--------------------|------|
|**Actif**|Un administrateur a chargé le complément et l’a affecté à des utilisateurs et/ou groupes.|Les utilisateurs et/ou groupes auxquels le complément est affecté voient celui-ci dans les clients Office concernés.|
|**Désactivé**|Un administrateur a désactivé le complément.|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.|
|**Deleted**|Un administrateur a supprimé le complément.|Les utilisateurs et/ou groupes auxquels le complément est affecté ne peuvent plus y accéder.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Mise à jour des compléments Office publiés via un déploiement centralisé

After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:

- **Complément métier** : si un administrateur a chargé explicitement un fichier manifeste lors de l’implémentation du déploiement centralisé via le Centre d’administration Office 365, il doit charger un nouveau fichier manifeste contenant les modifications souhaitées. Une fois que le fichier manifeste a été chargé, le complément est mis à jour au démarrage suivant des applications Office concernées.

  > [!NOTE]
  > Un administrateur n’a pas besoin de supprimer un complément LOB pour effectuer une mise à jour. Dans la section compléments, l’administrateur peut simplement choisir le complément LOB et appeler cette fonctionnalité en appuyant sur le bouton **mettre à jour le complément** présent dans le coin inférieur droit.
  > 
  > ![Capture d’écran illustrant la boîte de dialogue mettre à jour le complément dans le centre d’administration Office 365](../images/update-add-in-admin-center.png)

- **Complément de l’Office Store** : si un administrateur a sélectionné un complément dans l’Office Store lors de l’implémentation du déploiement centralisé via le Centre d’administration Office 365 et que le complément est mis à jour dans l’Office Store, le complément sera mis à jour plus tard via le déploiement centralisé. Le complément est mis à jour au démarrage suivant des applications Office concernées.

## <a name="end-user-experience-with-add-ins"></a>Expérience des utilisateurs finaux avec les compléments

Une fois qu’un complément a été publié via un déploiement centralisé, les utilisateurs finaux peuvent commencer à l’utiliser sur toutes les plateformes prises en charge par le complément.

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.

![Capture d’écran illustrant une section du ruban Office avec la commande Recherche de citation mise en évidence dans le complément Citations](../images/search-citation.png)

Si le complément ne prend pas en charge les commandes de complément, les utilisateurs peuvent l’ajouter à leur application Office en procédant comme suit :

1. Dans Word 2016 ou version ultérieure, Excel 2016 ou version ultérieure ou PowerPoint 2016 ou version ultérieure, choisissez **Insertion** > **Mes compléments**.
2. Sélectionnez l’onglet **Géré par l’administrateur** dans la fenêtre du complément.
3. Choisissez le complément, puis cliquez sur **Ajouter**.

    ![Screenshot shows the Admin Managed tab of the Office Add-ins page of an Office application. The Citations add-in is shown on the tab.](../images/office-add-ins-admin-managed.png)

Toutefois, pour Outlook 2016 ou version ultérieure, les utilisateurs peuvent procéder comme suit :

1. Dans Outlook, Choisissez **Accueil** > **Store**.
2. Sélectionnez l’élément **Géré par l’administrateur** dans la fenêtre du complément.
3. Choisissez le complément, puis **Ajouter**.

    ![Capture d’écran montrant la zone Géré par l’administrateur de la page Store de l’application Outlook.](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>Voir aussi

- [Déterminer si le déploiement centralisé des compléments fonctionne avec votre organisation Office 365](/office365/admin/manage/centralized-deployment-of-add-ins)
