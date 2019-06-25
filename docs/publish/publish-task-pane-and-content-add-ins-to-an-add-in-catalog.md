---
title: Publier des compléments du volet Office et de contenu dans un catalogue d’applications SharePoint
description: Pour rendre les compléments Office accessibles aux utilisateurs, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue d’applications pour leur organisation.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 20b97855ce50e3f70e602f511882761c6fd80655
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128558"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>Publier des compléments du volet Office et de contenu dans un catalogue d’applications SharePoint

Un catalogue d’applications est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Pour rendre les compléments Office accessibles aux utilisateurs dans leur organisation, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue d’applications pour leur organisation. Lorsqu’un administrateur enregistre un catalogue d’applications en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.

> [!IMPORTANT]
> - Les catalogues d’applications sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud `VersionOverrides` du [manifeste de complément](../develop/add-in-manifests.md), comme les commandes de complément.
> - Si vous ciblez un environnement de cloud ou hybride, nous vous recommandons d’[utiliser un déploiement centralisé via le centre d’administration Office 365](../publish/centralized-deployment.md) pour publier vos compléments.
> - Les catalogues d’applications dans SharePoint ne sont pas pris en charge par Office sur Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à [AppSource](/office/dev/store/submit-to-the-office-store).

## <a name="create-an-app-catalog"></a>Créer un catalogue d’applications

Suivez les étapes décrites dans l’une des sections suivantes pour créer un catalogue d’applications avec SharePoint Server local ou Office 365.

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>Création d’un catalogue d’applications pour SharePoint Server local

Pour créer le catalogue d’applications SharePoint, suivez les instructions de la section [Configurer le site de catalogue d'applications pour une application Web](https://docs.microsoft.com/fr-FR/sharepoint/administration/manage-the-app-catalog).

Une fois que vous avez créé le catalogue d’applications, suivez les étapes pour [publier un complément Office](#publish-an-office-add-in).

### <a name="to-create-an-app-catalog-on-office-365"></a>Pour créer catalogue d’applications Office 365

1. Aller au Centre d’administration Microsoft 365. Pour plus d’informations sur comment accéder au centre d’administration, voir [À propos du centre d’administration Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).

2. Dans la page Centre d’administration Microsoft 365, développez la liste des **centres d’administration**, puis sélectionnez **SharePoint**.

    > [!NOTE]
    > Vous devez utiliser le centre d’administration SharePoint classique pour créer le catalogue. Si c’est la première fois que vous accédez au centre d’administration SharePoint, sélectionnez **Centre d’administration SharePoint classique** dans le volet gauche.

3. Dans le volet Office situé à gauche, choisissez **Applications**.

4. Dans la page d’**applications**, choisissez **Catalogue d’applications**.
    > [!NOTE]
    > Si un catalogue d’applications est déjà créé et apparaît dans cette page, vous pouvez ignorer le reste de ces étapes et accéder à la section suivante de cet article pour publier votre complément dans le catalogue.

5. Dans la page **Site de catalogue d’applications**, cliquez sur **OK** pour accepter l’option par défaut et créer un site de catalogue d’applications.

6. Dans la page **Créer une collection de sites de catalogue d’applications**, indiquez le titre de votre site de catalogue d’applications.

7. Spécifiez l’**adresse du site web**.

8. Précisez qui est l’**administrateur **.

9. Choisissez 0 (zéro) comme **quota de ressources du serveur**. (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue d’applications.)

10. Sélectionnez **OK**.

## <a name="publish-an-office-add-in"></a>Publier un complément Office

Suivez les étapes décrites dans l’une des sections suivantes pour publier un complément Office dans un catalogue d’applications avec Office 365 ou avec SharePoint Server local.

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>Pour publier un complément Office dans un catalogue d’applications SharePoint sur Office 365

1. Aller au Centre d’administration Microsoft 365. Pour plus d’informations sur comment accéder au centre d’administration, voir [À propos du centre d’administration Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).
2. Dans la page Centre d’administration Microsoft 365, développez la liste des **centres d’administration**, puis sélectionnez **SharePoint**.
    > [!NOTE]
    > Vous devez utiliser le centre d’administration SharePoint classique pour créer le catalogue. Si c’est la première fois que vous accédez au centre d’administration SharePoint, sélectionnez **Centre d’administration SharePoint classique** dans le volet gauche.
3. Dans le volet Office situé à gauche, choisissez **Applications**.
4. Dans la page d’**applications**, choisissez **Catalogue d’applications**.
5. Choisissez **Distribuer des applications pour Office**.
6. Dans la page **Applications pour Office**, cliquez sur **Nouveau**.
7. Dans la boîte de dialogue **Ajouter un document**, sélectionnez le bouton **Choisir un fichier**.
8. Recherchez et spécifiez le fichier [manifeste](../develop/add-in-manifests.md) à télécharger, puis sélectionnez **Ouvrir**.
9. Dans la boîte de dialogue **Ajouter un document**, cliquez sur **OK**.

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>Pour publier un complément dans un catalogue d’applications avec SharePoint Server local

1. Ouvrez la page **Administration centrale**.
2. Dans le volet Office situé à gauche, choisissez **Applications**.
3. Dans la page **Applications**, sous **Gestion des applications**, sélectionnez **Gérer le catalogue d’applications**.
4. Dans la page **Gérer le catalogue d’applications**, vérifiez que vous avez sélectionné l’application web appropriée dans **Sélecteur d’applications web**.
5. Sélectionnez l’URL sous **URL du site** pour ouvrir le site du catalogue d’applications.
6. Choisissez **Distribuer des applications pour Office**.
7. Dans la page **Applications pour Office**, cliquez sur **Nouveau**.
8. Dans la boîte de dialogue **Ajouter un document**, sélectionnez le bouton **Choisir un fichier**.
9. Recherchez et spécifiez le fichier [manifeste](../develop/add-in-manifests.md) à télécharger, puis sélectionnez **Ouvrir**.
10. Dans la boîte de dialogue **Ajouter un document**, cliquez sur **OK**.

## <a name="insert-office-add-ins-from-the-app-catalog"></a>Insérer des compléments Office à partir du catalogue d’applications

Pour les applications Office en ligne, vous pouvez rechercher des compléments Office à partir du catalogue d’applications en procédant comme suit.

1. Ouvrez l’application Office en ligne (Excel, PowerPoint ou Word).
2. Créer ou ouvrir un document.
3. Sélectionnez **Insérer** > **des compléments**.
4. Dans la boîte de dialogue Compléments Office, choisissez l’onglet **MON ORGANISATION** Les compléments Office sont alors affichés.
5. Choisissez un complément Office, puis **Ajouter**.

Pour les applications Office sur le bureau, vous pouvez rechercher des compléments Office à partir du catalogue d’applications en procédant comme suit.

1. Ouvrir l’application de bureau Office (Excel, Word ou PowerPoint)
2. Accédez à **Fichier** > **Options** > **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité** > **Catalogues de compléments approuvés**.
3. Entrez l’URL du catalogue d’applications SharePoint dans la zone **URL du catalogue**, puis sélectionnez **Ajouter un catalogue**.
    Utilisez la forme la plus courte de l’URL. Par exemple, si l’URL du catalogue d’applications Office est :
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    Spécifiez simplement l’URL de la collection de sites parente :
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. Fermez puis rouvrez l’application Office. 
5. Sélectionnez **Insertion** > **Mes compléments**.
4. Dans la boîte de dialogue Compléments Office, choisissez l’onglet **MON ORGANISATION** Les compléments Office sont alors affichés.
5. Choisissez un complément Office, puis **Ajouter**.

Par ailleurs, un administrateur peut spécifier un catalogue d’applications sur SharePoint à l’aide d’une stratégie de groupe. Pour plus d’informations, reportez-vous à la section relative à l’[utilisation d’une stratégie de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des compléments Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
