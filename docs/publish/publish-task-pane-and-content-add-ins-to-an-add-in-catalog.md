---
title: Publication de compléments du volet Office et de contenu dans un catalogue SharePoint
description: Pour rendre les compléments Office accessibles aux utilisateurs, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation.
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432242"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>Publication de compléments du volet Office et de contenu dans un catalogue SharePoint

Un catalogue de compléments est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Pour rendre les compléments Office accessibles aux utilisateurs dans leur organisation, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation. Lorsqu’un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.

> [!IMPORTANT]
> - Les catalogues de compléments sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud `VersionOverrides` du [manifeste de complément](../develop/add-in-manifests.md), comme les commandes de complément.
> - Si vous ciblez un environnement de cloud ou hybride, nous vous recommandons d’[utiliser un déploiement centralisé via le centre d’administration Office 365](../publish/centralized-deployment.md) pour publier vos compléments.
> - Les catalogues SharePoint ne sont pas pris en charge dans Office pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à [AppSource](/office/dev/store/submit-to-the-office-store).   

## <a name="create-an-add-in-catalog"></a>Création d’un catalogue de compléments

Suivez les étapes décrites dans l’une des sections suivantes pour créer un catalogue de compléments sur SharePoint ou Office 365.

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a>Création d’un catalogue de compléments sur SharePoint local

> [!NOTE]
> L’interface utilisateur dans SharePoint local fait toujours référence aux compléments en tant qu’**applications**.

1. Accédez au **site Administration centrale**.

2. Dans le volet Office situé à gauche, cliquez sur **Applications**.

3. Sur la page **Applications**, sous **Gestion des applications**, sélectionnez **	Gérer le catalogue d’applications**.

4. Sur la page **Gérer le catalogue d’applications**, vérifiez que vous avez sélectionné l’application web appropriée dans **Sélecteur d’applications web**.

5. Choisissez  **Afficher les paramètres du site**.

6. Sur la page  **Paramètre du site**, choisissez  **Administrateurs de collections de sites** pour spécifier les administrateurs de collection de sites, puis choisissez **OK**.

7. Pour accorder des autorisations de site aux utilisateurs, choisissez  **Autorisations de site**, puis choisissez  **Accorder des autorisations**.

8. Dans la boîte de dialogue  **Partager le site de catalogue d’applications**, spécifiez des utilisateurs de site, définissez les autorisations appropriées pour ces derniers, puis éventuellement d’autres options, puis choisissez  **Partager**.

9. Pour ajouter un complément au catalogue de compléments Office, choisissez **Applications pour Office**.

### <a name="to-create-an-app-catalog-on-office-365"></a>Pour créer catalogue d’applications Office 365

SharePoint l’appelle un catalogue d’« Applications », mais vous pouvez également enregistrer des compléments Office dans le catalogue.

1. Aller au Centre d’administration Microsoft 365. Pour plus d’informations sur comment accéder au centre d’administration, voir [À propos du centre d’administration Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).

2. Dans la page Centre d’administration Microsoft 365, développez la liste des **centres d’administration**, puis sélectionnez **SharePoint**.

    > [!NOTE]
    > Vous devez utiliser le centre d’administration SharePoint classique pour créer le catalogue. Si c’est la première fois que vous accédez au centre d’administration SharePoint, sélectionnez **Centre d’administration SharePoint classique** dans le volet gauche.

3. Dans le volet Office situé à gauche, choisissez **Applications**.

4. Dans la page d’**applications**, choisissez **Catalogue d’applications**.
    > [!NOTE]
    > Si un catalogue d’applications est déjà créé et apparaît dans cette page, vous pouvez ignorer le reste de ces étapes et accéder à la section suivante de cet article pour publier votre complément dans le catalogue.

5. Dans la page **Site de catalogue d’applications**, cliquez sur **OK** pour accepter l’option par défaut et créer un site de catalogue.

6. Dans la page **Créer une collection de sites de catalogue d’applications**, indiquez le titre de votre site de catalogue.

7. Spécifiez l’**adresse du site web**.

8. Précisez qui est l’**administrateur **.

9. Choisissez 0 (zéro) comme **quota de ressources du serveur**. (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue d’applications.)

10. Sélectionnez **OK**.

Le catalogue d’applications est créé.

## <a name="publish-an-add-in-to-an-app-catalog"></a>Publication d’un complément dans un catalogue d’applications

Pour publier un complément dans un catalogue d’applications existant, procédez comme suit.

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

    Les compléments de contenu et de volet Office de ce catalogue sont désormais disponibles dans la boîte de dialogue **Compléments Office**. Pour y accéder, choisissez **Mes compléments** sous l’onglet **Insérer**, puis choisissez **MON ORGANISATION**.

## <a name="end-user-experience-with-the-add-in-catalog"></a>Expérience des utilisateurs finaux avec le catalogue des compléments

Les utilisateurs finaux peuvent accéder au catalogue des compléments dans une application Office en procédant comme suit :

1. Dans l’application Office, accédez à **Fichier**  >  **Options**  >  **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité**  >  **Catalogues de compléments approuvés**.

2. Spécifiez l’URL de la _collection de sites SharePoint parente_ du catalogue de compléments. 

    Par exemple, si l’URL du catalogue de compléments Office est :

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    Spécifiez simplement l’URL de la collection de sites parente :

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. Fermez puis rouvrez l’application Office. Le catalogue de compléments est disponible dans la boîte de dialogue **Compléments Office**.

Par ailleurs, un administrateur peut spécifier un catalogue de compléments Office sur SharePoint à l’aide d’une stratégie de groupe. Pour plus d’informations, reportez-vous à la section relative à l’[utilisation d’une stratégie de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des compléments Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
